"""
Teams-specific MCP operations.

Higher-level wrapper for Microsoft Teams interactions.
"""

import logging
import re
from datetime import datetime
from typing import List, Optional, Dict, Any, Tuple

from .mcp_client import MCPClient, MCPClientError
from ..models import EmailRecord, CommandType
from ..config import settings

logger = logging.getLogger(__name__)


class TeamsClient:
    """High-level Teams operations using MCP."""

    def __init__(self, mcp_client: Optional[MCPClient] = None):
        self.mcp = mcp_client or MCPClient()
        self._channel_id = settings.teams_channel_id
        self._chat_id = settings.teams_chat_id
        self._team_id: Optional[str] = None

    def send_notification(
        self,
        content: str,
        use_channel: bool = True
    ) -> Optional[str]:
        """
        Send a notification to Teams.

        Args:
            content: HTML content to send
            use_channel: If True, send to channel; otherwise send to chat

        Returns:
            Message ID if sent successfully
        """
        try:
            if use_channel and self._channel_id:
                if not self._team_id:
                    self._discover_team_id()

                result = self.mcp.send_channel_message(
                    team_id=self._team_id,
                    channel_id=self._channel_id,
                    content=content,
                    content_type="html"
                )
            elif self._chat_id:
                result = self.mcp.send_chat_message(
                    chat_id=self._chat_id,
                    content=content,
                    content_type="html"
                )
            else:
                logger.warning("No Teams channel or chat configured")
                return None

            return result.get("id")

        except MCPClientError as e:
            logger.error(f"Failed to send Teams notification: {e}")
            return None

    def send_email_notification(self, email: EmailRecord) -> Optional[str]:
        """
        Send a formatted email notification to Teams.

        Args:
            email: The email record to notify about

        Returns:
            Teams message ID if sent
        """
        # Determine priority emoji
        priority_emoji = {
            1: "🚨",
            2: "⚡",
            3: "📧",
            4: "📬",
            5: "📭"
        }.get(email.priority, "📧")

        # Determine category emoji
        category_emoji = {
            "urgent": "🔴",
            "action_required": "💼",
            "fyi": "ℹ️",
            "meeting": "📅",
            "spam_candidate": "🗑️",
            "forward_candidate": "↪️"
        }.get(email.category.value if email.category else "", "📧")

        # Build notification content
        content = f"""
        <h3>{priority_emoji} New Email Requiring Action</h3>
        <hr>
        <p><b>From:</b> {email.sender_name or email.sender_email} &lt;{email.sender_email}&gt;</p>
        <p><b>Subject:</b> {email.subject}</p>
        <p><b>Priority:</b> {email.priority}/5 | <b>Category:</b> {category_emoji} {email.category.value if email.category else 'Unknown'}</p>
        <hr>
        <h4>📝 Summary:</h4>
        <p>{email.summary or email.body_preview[:300]}</p>
        """

        if email.current_draft:
            content += f"""
            <hr>
            <h4>✉️ Draft Reply:</h4>
            <blockquote>{email.current_draft}</blockquote>
            """

        content += f"""
        <hr>
        <p><b>Token:</b> <code>[{email.approval_token}]</code></p>
        <p>
        Reply with:<br>
        • <code>approve</code> or <code>{email.approval_token}</code> - Send this reply<br>
        • <code>edit: [your changes]</code> - Modify the draft<br>
        • <code>rewrite</code> - Generate a new draft<br>
        • <code>ignore</code> - Skip, no reply needed<br>
        • <code>more</code> - Show full email<br>
        • <code>spam</code> - Mark as spam
        </p>
        """

        message_id = self.send_notification(content)

        if message_id:
            logger.info(f"Sent Teams notification for email '{email.subject}'")

        return message_id

    def send_daily_digest(
        self,
        stats: Dict[str, Any],
        pending_emails: List[EmailRecord],
        spam_filtered: int = 0
    ) -> Optional[str]:
        """
        Send daily digest to Teams.

        Args:
            stats: Email statistics
            pending_emails: Emails still awaiting action
            spam_filtered: Number of spam emails filtered

        Returns:
            Teams message ID if sent
        """
        today = datetime.utcnow().strftime("%B %d, %Y")

        content = f"""
        <h3>📊 Email Daily Digest - {today}</h3>
        <hr>
        <p>
        📬 <b>Processed:</b> {stats.get('total_emails', 0)} emails<br>
        ✅ <b>Auto-handled:</b> {stats.get('auto_handled', 0)} (FYI/spam)<br>
        📝 <b>Drafts sent:</b> {stats.get('emails_sent', 0)}<br>
        ⏳ <b>Awaiting action:</b> {len(pending_emails)}
        </p>
        """

        if pending_emails:
            content += "<hr><h4>🚨 Needs Attention:</h4><ol>"
            for email in pending_emails[:5]:  # Top 5
                priority_label = "URGENT" if email.priority <= 2 else ""
                content += f"<li>[{priority_label}] {email.subject} - {email.sender_email}</li>"
            if len(pending_emails) > 5:
                content += f"<li>... and {len(pending_emails) - 5} more</li>"
            content += "</ol>"

        if spam_filtered > 0:
            content += f"<hr><p>🗑️ <b>Spam filtered:</b> {spam_filtered} emails</p>"

        return self.send_notification(content)

    def get_recent_replies(
        self,
        since_message_id: Optional[str] = None,
        limit: int = 50
    ) -> List[Dict[str, Any]]:
        """
        Get recent replies from Teams channel/chat.

        Args:
            since_message_id: Only get messages after this ID
            limit: Maximum messages to fetch

        Returns:
            List of message dictionaries
        """
        try:
            if self._channel_id:
                if not self._team_id:
                    self._discover_team_id()

                messages = self.mcp.list_channel_messages(
                    team_id=self._team_id,
                    channel_id=self._channel_id,
                    top=limit
                )
            elif self._chat_id:
                messages = self.mcp.list_chat_messages(
                    chat_id=self._chat_id,
                    top=limit
                )
            else:
                return []

            return messages

        except MCPClientError as e:
            logger.error(f"Failed to get Teams replies: {e}")
            return []

    def parse_command(self, message_text: str) -> Tuple[CommandType, Optional[str]]:
        """
        Parse a user command from a Teams message.

        Args:
            message_text: The raw message text

        Returns:
            Tuple of (CommandType, optional parameter)
        """
        text = message_text.strip().lower()

        # Direct commands
        if text in ["approve", "send", "yes", "y"]:
            return CommandType.APPROVE, None

        if text in ["ignore", "skip", "no", "n"]:
            return CommandType.IGNORE, None

        if text == "rewrite":
            return CommandType.REWRITE, None

        if text == "more":
            return CommandType.MORE, None

        if text == "spam":
            return CommandType.SPAM, None

        if text in ["done", "delete"]:
            return CommandType.DELETE, None

        # Token-based approval (6-char hex)
        if re.match(r"^[a-f0-9]{6}$", text):
            return CommandType.APPROVE, text

        # Edit command with content
        edit_match = re.match(r"^edit:\s*(.+)$", text, re.IGNORECASE | re.DOTALL)
        if edit_match:
            return CommandType.EDIT, edit_match.group(1).strip()

        # Forward command
        forward_match = re.match(r"^forward\s+(?:to\s+)?(.+)$", text, re.IGNORECASE)
        if forward_match:
            return CommandType.FORWARD, forward_match.group(1).strip()

        return CommandType.UNKNOWN, message_text

    def _discover_team_id(self):
        """Discover the team ID for the configured channel."""
        try:
            teams = self.mcp.list_joined_teams()
            for team in teams:
                channels = self.mcp.list_team_channels(team.get("id", ""))
                for channel in channels:
                    if channel.get("id") == self._channel_id:
                        self._team_id = team.get("id")
                        logger.info(f"Discovered team ID: {self._team_id}")
                        return

            logger.warning(f"Could not find team for channel {self._channel_id}")
        except MCPClientError as e:
            logger.error(f"Failed to discover team ID: {e}")

    def close(self):
        """Close underlying connections."""
        self.mcp.close()
