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
        self._team_id: Optional[str] = getattr(settings, 'teams_team_id', None)

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
        Get recent replies from Teams channel/chat, including thread replies.

        Args:
            since_message_id: Only get messages after this ID
            limit: Maximum messages to fetch

        Returns:
            List of message dictionaries (includes both top-level and thread replies)
        """
        all_messages = []
        try:
            if self._channel_id:
                if not self._team_id:
                    self._discover_team_id()

                # If we still don't have a team ID, can't fetch messages
                if not self._team_id:
                    logger.warning("No team ID available for channel")
                    return []

                # Get top-level messages
                messages = self.mcp.list_channel_messages(
                    team_id=self._team_id,
                    channel_id=self._channel_id,
                    top=limit
                )

                # Also fetch thread replies for recent messages
                for msg in messages[:10]:  # Check threads on last 10 messages
                    if not msg:
                        continue
                    msg_id = msg.get("id")
                    if msg_id:
                        # Add the parent message
                        all_messages.append(msg)
                        # Fetch thread replies
                        try:
                            replies = self._get_thread_replies(msg_id)
                            for reply in replies:
                                if reply:
                                    # Tag with parent message ID for context
                                    reply["_parent_message_id"] = msg_id
                                    all_messages.append(reply)
                        except Exception as e:
                            logger.debug(f"Could not fetch replies for {msg_id}: {e}")

            elif self._chat_id:
                messages = self.mcp.list_chat_messages(
                    chat_id=self._chat_id,
                    top=limit
                )
                all_messages = messages
            else:
                return []

            return all_messages

        except MCPClientError as e:
            logger.error(f"Failed to get Teams replies: {e}")
            return []

    def _get_thread_replies(self, parent_message_id: str) -> List[Dict[str, Any]]:
        """Fetch replies to a specific message thread."""
        try:
            result = self.mcp.call_tool("list-channel-message-replies", {
                "team_id": self._team_id,
                "channel_id": self._channel_id,
                "message_id": parent_message_id
            })
            if result is None:
                logger.debug(f"No thread replies for message {parent_message_id}")
                return []
            if isinstance(result, list):
                replies = [r for r in result if r is not None]
                if replies:
                    logger.info(f"Found {len(replies)} thread replies for message {parent_message_id}")
                return replies
            replies = result.get("replies", result.get("value", []))
            if replies is None:
                return []
            filtered_replies = [r for r in replies if r is not None]
            if filtered_replies:
                logger.info(f"Found {len(filtered_replies)} thread replies for message {parent_message_id}")
            return filtered_replies
        except Exception as e:
            logger.warning(f"Failed to get thread replies for {parent_message_id}: {e}")
            return []

    def parse_command(self, message_text: str) -> Tuple[CommandType, Optional[str]]:
        """
        Parse a user command from a Teams message.
        Supports both exact commands and conversational phrases.

        Args:
            message_text: The raw message text

        Returns:
            Tuple of (CommandType, optional parameter)
        """
        text = message_text.strip().lower()

        # Direct commands
        if text in ["approve", "send", "yes", "y", "ok", "looks good", "send it"]:
            return CommandType.APPROVE, None

        if text in ["ignore", "skip", "no", "n", "pass", "not now", "later"]:
            return CommandType.IGNORE, None

        if text in ["rewrite", "try again", "redo"]:
            return CommandType.REWRITE, None

        if text in ["more", "show more", "full email", "details"]:
            return CommandType.MORE, None

        if text == "spam":
            return CommandType.SPAM, None

        if text in ["done", "delete"]:
            return CommandType.DELETE, None

        # Spam batch commands
        if text in ["dismiss all", "dismiss_all", "archive all", "clear spam"]:
            return CommandType.DISMISS_ALL, None

        if text == "review":
            return CommandType.REVIEW, None

        # Keep command with parameter (index or keyword)
        keep_match = re.match(r"^keep\s+(.+)$", text, re.IGNORECASE)
        if keep_match:
            return CommandType.KEEP, keep_match.group(1).strip()

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

        # Conversational spam/junk detection
        spam_phrases = [
            "junk", "is junk", "all junk", "this is junk",
            "is spam", "this is spam", "mark as spam",
            "trash", "garbage", "delete this", "not interested",
            "unsubscribe", "stop sending", "don't want this"
        ]
        for phrase in spam_phrases:
            if phrase in text:
                return CommandType.SPAM, None

        # Conversational ignore detection
        ignore_phrases = [
            "don't need to reply", "no reply needed", "no action",
            "not important", "can ignore", "skip this",
            "doesn't need", "don't care",
            # Notification-related (user doesn't want alerts for this type)
            "don't need to be notified", "dont need to be notified",
            "don't notify", "dont notify", "no notification",
            "stop notifying", "don't alert", "dont alert",
            "don't need notification", "dont need notification"
        ]
        for phrase in ignore_phrases:
            if phrase in text:
                return CommandType.IGNORE, None

        # Conversational approval
        approve_phrases = [
            "looks good", "send that", "go ahead", "that works",
            "perfect", "good to go", "ship it"
        ]
        for phrase in approve_phrases:
            if phrase in text:
                return CommandType.APPROVE, None

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
