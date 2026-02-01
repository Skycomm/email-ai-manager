"""
Teams Communications Agent - Handles all Microsoft Teams interactions.

This agent:
- Formats and sends notifications to Teams
- Parses user commands from replies
- Manages conversation threads
- Sends daily digests
"""

import logging
import re
from datetime import datetime
from typing import Dict, Any, List, Optional

from .base import BaseAgent
from ..db import Database
from ..models import EmailRecord, CommandType, EmailState
from ..integrations.mcp_teams import TeamsClient

logger = logging.getLogger(__name__)


TEAMS_COMMS_SYSTEM_PROMPT = """You are the Teams Communications Agent, responsible for formatting notifications and parsing user commands.

Your job is to:
1. Create clear, well-formatted Teams messages for email notifications
2. Parse user replies to extract commands
3. Handle conversation threading

Keep messages concise but informative. Use formatting (bold, lists) to improve readability.
"""


class TeamsCommsAgent(BaseAgent):
    """
    Agent for handling Microsoft Teams communications.
    """

    def __init__(self, db: Database, teams_client: Optional[TeamsClient] = None):
        super().__init__(
            db=db,
            name="teams_comms",
            system_prompt=TEAMS_COMMS_SYSTEM_PROMPT
        )
        self.teams = teams_client or TeamsClient()
        self._last_check_timestamp: Optional[datetime] = None
        self._processed_message_ids: set = set()

    async def process(self, *args, **kwargs) -> Any:
        """Main process method - delegates to specific operations."""
        pass

    async def notify_email(self, email: EmailRecord) -> Optional[str]:
        """
        Send a notification about an email to Teams.

        Args:
            email: The email to notify about

        Returns:
            Teams message ID if sent successfully
        """
        message_id = self.teams.send_email_notification(email)

        if message_id:
            self.log_action(
                "teams_notification_sent",
                email_id=email.id,
                details={
                    "teams_message_id": message_id,
                    "subject": email.subject
                }
            )

        return message_id

    def _clean_email_body(self, body: str) -> str:
        """
        Clean up email body for display - extract readable text from HTML.

        Args:
            body: Raw email body (may be HTML)

        Returns:
            Cleaned text suitable for Teams display
        """
        if not body:
            return "(No content)"

        text = body

        # Remove HTML comments
        text = re.sub(r'<!--.*?-->', '', text, flags=re.DOTALL)

        # Remove style and script tags entirely
        text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.DOTALL | re.IGNORECASE)
        text = re.sub(r'<script[^>]*>.*?</script>', '', text, flags=re.DOTALL | re.IGNORECASE)

        # Replace common block elements with newlines
        text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'</p>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'</div>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'</tr>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'</li>', '\n', text, flags=re.IGNORECASE)

        # Remove all remaining HTML tags
        text = re.sub(r'<[^>]+>', '', text)

        # Decode common HTML entities
        text = text.replace('&nbsp;', ' ')
        text = text.replace('&amp;', '&')
        text = text.replace('&lt;', '<')
        text = text.replace('&gt;', '>')
        text = text.replace('&quot;', '"')
        text = text.replace('&#39;', "'")

        # Remove zero-width characters (used for preview hacking)
        text = re.sub(r'[\u200b\u200c\u200d\ufeff]', '', text)

        # Collapse multiple whitespace/newlines
        text = re.sub(r'\n\s*\n', '\n\n', text)
        text = re.sub(r' +', ' ', text)

        # Trim and limit length for Teams
        text = text.strip()
        if len(text) > 3000:
            text = text[:3000] + '\n\n... (truncated)'

        return text

    async def send_full_email(self, email: EmailRecord) -> Optional[str]:
        """
        Send the full email content to Teams.

        Args:
            email: The email to display

        Returns:
            Teams message ID if sent
        """
        # Clean up the email body for readable display
        body_content = self._clean_email_body(email.body_full or email.body_preview)

        content = f"""<h4>Full Email Content</h4>
<hr>
<p><b>From:</b> {email.sender_name or email.sender_email} &lt;{email.sender_email}&gt;</p>
<p><b>To:</b> {', '.join(email.to_recipients)}</p>
<p><b>Subject:</b> {email.subject}</p>
<p><b>Received:</b> {email.received_at.strftime('%Y-%m-%d %H:%M')}</p>
<hr>
<pre style="white-space:pre-wrap;font-family:inherit;">{body_content}</pre>"""

        if email.has_attachments:
            content += "<hr><p><i>📎 This email has attachments (not shown)</i></p>"

        return self.teams.send_notification(content)

    async def send_digest(
        self,
        stats: Dict[str, Any],
        pending_emails: List[EmailRecord]
    ) -> Optional[str]:
        """
        Send the daily digest to Teams.

        Args:
            stats: Email statistics
            pending_emails: Emails still awaiting action

        Returns:
            Teams message ID if sent
        """
        spam_count = stats.get("spam_filtered", 0)
        message_id = self.teams.send_daily_digest(stats, pending_emails, spam_count)

        if message_id:
            self.log_action(
                "daily_digest_sent",
                details={
                    "total_emails": stats.get("total_emails", 0),
                    "pending_count": len(pending_emails)
                }
            )

        return message_id

    async def check_for_commands(self) -> List[Dict[str, Any]]:
        """
        Check for new commands from Teams replies.

        Returns:
            List of parsed commands with associated emails
        """
        commands = []

        # Commands that don't require an email lookup (handled by coordinator)
        BATCH_COMMANDS = [CommandType.DISMISS_ALL, CommandType.REVIEW, CommandType.KEEP, CommandType.ARCHIVE_ALL]
        # Numbered commands (like "more 3", "spam 5", "mute 2") are looked up by coordinator via summary_email_mapping
        NUMBERED_COMMANDS = [CommandType.MORE, CommandType.SPAM, CommandType.MUTE]

        try:
            # Get recent messages
            messages = self.teams.get_recent_replies(limit=50)

            for message in messages:
                if not message or not isinstance(message, dict):
                    continue
                msg_id = message.get("id", "")

                # Skip already processed
                if msg_id in self._processed_message_ids:
                    continue

                # Skip messages from the bot itself
                if self._is_bot_message(message):
                    continue

                # Parse command from message
                text = self._extract_text(message)
                if not text:
                    continue

                command_type, parameter = self.teams.parse_command(text)

                if command_type == CommandType.UNKNOWN:
                    continue

                # Batch commands don't need an email lookup
                if command_type in BATCH_COMMANDS:
                    commands.append({
                        "email": None,
                        "command_type": command_type.value,
                        "parameter": parameter,
                        "message_id": msg_id
                    })
                    self.log_action(
                        "batch_command_received",
                        user_command=command_type.value,
                        details={"parameter": parameter}
                    )
                    self._processed_message_ids.add(msg_id)
                    continue

                # Numbered commands (e.g., "more 3", "spam 5") pass through to coordinator
                # Coordinator will look up the email via summary_email_mapping
                if command_type in NUMBERED_COMMANDS and parameter and parameter.isdigit():
                    commands.append({
                        "email": None,
                        "command_type": command_type.value,
                        "parameter": parameter,
                        "message_id": msg_id
                    })
                    self.log_action(
                        "numbered_command_received",
                        user_command=command_type.value,
                        details={"number": parameter}
                    )
                    self._processed_message_ids.add(msg_id)
                    continue

                # Find associated email for regular commands
                email = self._find_email_for_command(message, parameter)

                if email:
                    commands.append({
                        "email": email,
                        "command_type": command_type.value,
                        "parameter": parameter,
                        "message_id": msg_id
                    })

                    self.log_action(
                        "command_received",
                        email_id=email.id,
                        user_command=command_type.value,
                        details={"parameter": parameter}
                    )

                self._processed_message_ids.add(msg_id)

        except Exception as e:
            logger.error(f"Error checking for commands: {e}")
            self.log_action("command_check_error", error=str(e), success=False)

        return commands

    async def send_confirmation(
        self,
        email: EmailRecord,
        action: str,
        success: bool = True
    ) -> Optional[str]:
        """
        Send a confirmation message after an action.

        Args:
            email: The email that was acted upon
            action: Description of the action
            success: Whether it succeeded

        Returns:
            Teams message ID if sent
        """
        if success:
            emoji = "✅"
            status = "completed"
        else:
            emoji = "❌"
            status = "failed"

        content = f"""
        <p>{emoji} <b>{action.title()}</b> {status}</p>
        <p><i>Email:</i> {email.subject}</p>
        """

        return self.teams.send_notification(content)

    async def send_error(
        self,
        email: EmailRecord,
        error_message: str
    ) -> Optional[str]:
        """
        Send an error notification.

        Args:
            email: The related email
            error_message: Error description

        Returns:
            Teams message ID if sent
        """
        content = f"""
        <p>❌ <b>Error Processing Email</b></p>
        <p><i>Subject:</i> {email.subject}</p>
        <p><i>Error:</i> {error_message}</p>
        <p><i>The email has been flagged for manual review.</i></p>
        """

        return self.teams.send_notification(content)

    def _is_bot_message(self, message: Dict[str, Any]) -> bool:
        """Check if a message was sent by the bot."""
        # Check for bot indicators in the message
        from_info = message.get("from", {})
        user = from_info.get("user", {}) or from_info.get("application", {})

        # If it's an application, it's likely the bot
        if from_info.get("application"):
            return True

        return False

    def _extract_text(self, message: Dict[str, Any]) -> str:
        """Extract plain text from a Teams message."""
        body = message.get("body", {})
        content = body.get("content", "")
        content_type = body.get("contentType", "text")

        if content_type == "html":
            # Simple HTML stripping
            text = re.sub(r"<[^>]+>", "", content)
            text = text.replace("&nbsp;", " ").replace("&amp;", "&")
            return text.strip()

        return content.strip()

    def _find_email_for_command(
        self,
        message: Dict[str, Any],
        parameter: Optional[str]
    ) -> Optional[EmailRecord]:
        """
        Find the email associated with a command.

        Uses:
        - Approval token if provided
        - Thread context
        - Most recent awaiting email
        """
        # If parameter looks like an approval token
        if parameter and len(parameter) == 6:
            email = self.db.get_email_by_approval_token(parameter)
            if email:
                return email

        # Try to find by thread context (reply to notification)
        # This would require tracking Teams thread -> email mapping

        # Fall back to most recent awaiting approval
        pending = self.db.get_emails_by_state(
            EmailState.AWAITING_APPROVAL,
            limit=1
        )
        if pending:
            return pending[0]

        return None

    def close(self):
        """Clean up resources."""
        self.teams.close()
