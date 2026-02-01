"""
Teams-specific MCP operations.

Higher-level wrapper for Microsoft Teams interactions.
"""

import logging
import re
import hashlib
from datetime import datetime, timedelta
from typing import List, Optional, Dict, Any, Tuple

from .mcp_client import MCPClient, MCPClientError
from ..models import EmailRecord, CommandType
from ..config import settings

logger = logging.getLogger(__name__)


def _generate_dedup_key(sender_email: str, subject: str) -> str:
    """
    Generate a deduplication key for similar notifications.
    Groups emails by sender domain + subject pattern (removing numbers/timestamps).

    Examples:
        - "VPN Connection Down Alert #123" and "VPN Connection Down Alert #456" -> same key
        - "Netflix promo 1" and "Netflix promo 2" -> same key
    """
    # Extract domain from sender
    domain = sender_email.split('@')[-1].lower() if '@' in sender_email else sender_email.lower()

    # Normalize subject - remove numbers, timestamps, IDs
    normalized_subject = subject.lower()
    # Remove numbers (IDs, counts, etc.)
    normalized_subject = re.sub(r'\d+', '#', normalized_subject)
    # Remove common timestamp patterns
    normalized_subject = re.sub(r'\d{1,2}:\d{2}(:\d{2})?', '', normalized_subject)
    normalized_subject = re.sub(r'\d{1,2}/\d{1,2}/\d{2,4}', '', normalized_subject)
    # Remove extra whitespace
    normalized_subject = ' '.join(normalized_subject.split())

    # Create hash for consistent key length
    key_input = f"{domain}:{normalized_subject}"
    return hashlib.md5(key_input.encode()).hexdigest()[:16]


class TeamsClient:
    """High-level Teams operations using MCP."""

    def __init__(self, mcp_client: Optional[MCPClient] = None, db=None):
        self.mcp = mcp_client or MCPClient()
        self.db = db  # Database for persisting pending notifications
        self._channel_id = settings.teams_channel_id
        self._chat_id = settings.teams_chat_id
        self._team_id: Optional[str] = getattr(settings, 'teams_team_id', None)
        # In-memory cache of pending notifications (dedup_key -> {message_id, count, emails, last_updated})
        # This is loaded from DB on startup if db is provided
        self._pending_notifications: Dict[str, Dict[str, Any]] = {}
        self._load_pending_notifications()

    def _load_pending_notifications(self):
        """Load pending notifications from database."""
        if self.db:
            try:
                import json
                data = self.db.get_setting("pending_notifications")
                if data:
                    self._pending_notifications = json.loads(data)
                    # Clean up old entries (> 24 hours)
                    now = datetime.utcnow()
                    to_remove = []
                    for key, info in self._pending_notifications.items():
                        last_updated = datetime.fromisoformat(info.get("last_updated", "2000-01-01"))
                        if (now - last_updated) > timedelta(hours=24):
                            to_remove.append(key)
                    for key in to_remove:
                        del self._pending_notifications[key]
                    if to_remove:
                        self._save_pending_notifications()
                        logger.info(f"Cleaned up {len(to_remove)} old pending notifications")
            except Exception as e:
                logger.warning(f"Could not load pending notifications: {e}")
                self._pending_notifications = {}

    def _save_pending_notifications(self):
        """Save pending notifications to database."""
        if self.db:
            try:
                import json
                self.db.set_setting("pending_notifications", json.dumps(self._pending_notifications))
            except Exception as e:
                logger.warning(f"Could not save pending notifications: {e}")

    def clear_pending_notification(self, dedup_key: str):
        """Clear a pending notification (called when user responds)."""
        if dedup_key in self._pending_notifications:
            del self._pending_notifications[dedup_key]
            self._save_pending_notifications()
            logger.info(f"Cleared pending notification: {dedup_key}")

    def clear_pending_for_email(self, email: EmailRecord):
        """Clear pending notification for an email."""
        dedup_key = _generate_dedup_key(email.sender_email, email.subject)
        self.clear_pending_notification(dedup_key)

    def update_message(
        self,
        message_id: str,
        content: str
    ) -> bool:
        """
        Update an existing Teams message.

        Args:
            message_id: ID of the message to update
            content: New HTML content

        Returns:
            True if update succeeded
        """
        try:
            if self._chat_id:
                self.mcp.update_chat_message(
                    chat_id=self._chat_id,
                    message_id=message_id,
                    content=content,
                    content_type="html"
                )
            elif self._channel_id:
                if not self._team_id:
                    self._discover_team_id()
                self.mcp.update_channel_message(
                    team_id=self._team_id,
                    channel_id=self._channel_id,
                    message_id=message_id,
                    content=content,
                    content_type="html"
                )
            return True
        except MCPClientError as e:
            logger.warning(f"Failed to update Teams message: {e}")
            return False

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

    def send_fyi_notification_deduped(
        self,
        email: EmailRecord,
        content_builder: callable = None
    ) -> Tuple[Optional[str], bool]:
        """
        Send an FYI notification with deduplication.

        If a similar notification (same sender domain + subject pattern) already exists
        and hasn't been responded to, UPDATE that message instead of creating a new one.

        Args:
            email: The email record
            content_builder: Optional function to build content, receives (email, count, email_ids)

        Returns:
            Tuple of (message_id, was_updated) - was_updated is True if existing message was updated
        """
        dedup_key = _generate_dedup_key(email.sender_email, email.subject)

        # Check if we have a pending notification for this pattern
        existing = self._pending_notifications.get(dedup_key)

        if existing:
            # Update the existing notification
            count = existing.get("count", 1) + 1
            email_ids = existing.get("email_ids", [])
            email_ids.append(email.id)
            message_id = existing.get("message_id")

            # Track status history for alerts (up/down/changed/etc.)
            status_history = existing.get("status_history", [])
            current_status = self._extract_alert_status(email.subject)
            if current_status:
                status_history.append({
                    "status": current_status,
                    "time": datetime.utcnow().isoformat()
                })

            # Build updated content showing the count and status summary
            if content_builder:
                content = content_builder(email, count, email_ids, status_history)
            else:
                content = self._build_deduped_fyi_content(email, count, email_ids, status_history)

            # Try to update the existing message
            if message_id and self.update_message(message_id, content):
                # Update tracking
                self._pending_notifications[dedup_key] = {
                    "message_id": message_id,
                    "count": count,
                    "email_ids": email_ids,
                    "last_updated": datetime.utcnow().isoformat(),
                    "subject_pattern": email.subject,
                    "sender_domain": email.sender_email.split('@')[-1] if '@' in email.sender_email else email.sender_email,
                    "status_history": status_history
                }
                self._save_pending_notifications()
                logger.info(f"Updated existing notification (count={count}): {email.subject}")
                return message_id, True
            else:
                # Couldn't update, fall through to create new
                logger.warning(f"Could not update existing message, creating new one")

        # No existing notification or update failed - create new
        # Initialize status history for alerts
        initial_status = self._extract_alert_status(email.subject)
        status_history = [{"status": initial_status, "time": datetime.utcnow().isoformat()}] if initial_status else []

        if content_builder:
            content = content_builder(email, 1, [email.id], status_history)
        else:
            content = self._build_deduped_fyi_content(email, 1, [email.id], status_history)

        message_id = self.send_notification(content)

        if message_id:
            # Track this notification for deduplication
            self._pending_notifications[dedup_key] = {
                "message_id": message_id,
                "count": 1,
                "email_ids": [email.id],
                "last_updated": datetime.utcnow().isoformat(),
                "subject_pattern": email.subject,
                "sender_domain": email.sender_email.split('@')[-1] if '@' in email.sender_email else email.sender_email,
                "status_history": status_history
            }
            self._save_pending_notifications()
            logger.info(f"Created new FYI notification: {email.subject}")

        return message_id, False

    def _extract_alert_status(self, subject: str) -> Optional[str]:
        """
        Extract alert status from subject line.
        Returns 'up', 'down', 'changed', etc. or None if not an alert.
        """
        subject_lower = subject.lower()

        # Check for up/down/online/offline status
        if any(word in subject_lower for word in ['is up', 'is online', 'is back', 'restored', 'recovered']):
            return 'up'
        elif any(word in subject_lower for word in ['is down', 'is offline', 'went down', 'failed', 'unreachable']):
            return 'down'
        elif any(word in subject_lower for word in ['changed', 'connectivity changed', 'status changed']):
            # Try to determine direction from subject
            if 'up' in subject_lower or 'online' in subject_lower:
                return 'up'
            elif 'down' in subject_lower or 'offline' in subject_lower:
                return 'down'
            return 'changed'

        return None

    def _summarize_status_history(self, status_history: List[Dict[str, Any]]) -> str:
        """
        Create a human-readable summary of status changes.
        E.g., "Went up/down 5 times, currently UP"
        """
        if not status_history:
            return ""

        # Count transitions
        up_count = sum(1 for s in status_history if s.get("status") == "up")
        down_count = sum(1 for s in status_history if s.get("status") == "down")
        changed_count = sum(1 for s in status_history if s.get("status") == "changed")

        # Get current status (last known)
        current = status_history[-1].get("status", "unknown") if status_history else "unknown"
        current_display = current.upper() if current in ["up", "down"] else current

        total_changes = len(status_history)

        if total_changes <= 1:
            return f"Status: {current_display}"
        elif up_count > 0 and down_count > 0:
            return f"⚡ Flapped {total_changes}x (↑{up_count} ↓{down_count}) → Currently {current_display}"
        elif up_count > 0:
            return f"🔄 {up_count} recoveries → Currently {current_display}"
        elif down_count > 0:
            return f"⚠️ {down_count} failures → Currently {current_display}"
        else:
            return f"🔄 Changed {total_changes} times → Currently {current_display}"

    def _build_deduped_fyi_content(
        self,
        email: EmailRecord,
        count: int,
        email_ids: List[str],
        status_history: Optional[List[Dict[str, Any]]] = None
    ) -> str:
        """Build content for a deduplicated FYI notification."""
        sender = email.sender_name or email.sender_email
        status_history = status_history or []

        if count > 1:
            status_summary = self._summarize_status_history(status_history)
            status_line = f'<p style="background:#fff3cd;padding:6px;border-radius:4px;"><b>{status_summary}</b></p>' if status_summary else ""

            content = f"""<div style="border-left: 3px solid #ffa500; padding-left: 10px;">
<p><b>ℹ️ FYI ({count}x)</b> - Similar alerts grouped</p>
<p><b>From:</b> {sender}</p>
<p><b>Latest:</b> {email.subject}</p>
{status_line}
<hr>
<p>{email.body_preview[:200]}...</p>
<hr>
<p><code>ignore</code> - dismiss all • <code>mute {email.sender_email}</code> - stop these alerts</p>
</div>"""
        else:
            content = f"""<div style="border-left: 3px solid #17a2b8; padding-left: 10px;">
<p><b>ℹ️ FYI</b></p>
<p><b>From:</b> {sender}</p>
<p><b>Subject:</b> {email.subject}</p>
<hr>
<p>{email.body_preview[:200]}...</p>
<hr>
<p><code>ignore</code> - dismiss • <code>mute {email.sender_email}</code> - stop these</p>
</div>"""

        return content

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
                # Also fetch quote replies for recent messages
                for msg in messages[:10]:  # Check last 10 messages for replies
                    if not msg:
                        continue
                    msg_id = msg.get("id")
                    if msg_id:
                        all_messages.append(msg)
                        # Fetch quote replies
                        try:
                            replies = self.mcp.list_chat_message_replies(
                                chat_id=self._chat_id,
                                message_id=msg_id
                            )
                            for reply in replies:
                                if reply:
                                    reply["_parent_message_id"] = msg_id
                                    all_messages.append(reply)
                        except Exception as e:
                            logger.debug(f"Could not fetch chat replies for {msg_id}: {e}")
                # Add any remaining messages we didn't process for replies
                for msg in messages[10:]:
                    if msg:
                        all_messages.append(msg)
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

        # EXPLICIT confirmation required to actually send emails
        # This is the ONLY way to send via Teams
        if text in ["confirm send", "confirm_send", "confirmsend"]:
            return CommandType.CONFIRM_SEND, None

        # Direct commands - these NO LONGER send emails, just prompt for confirmation
        if text in ["approve", "send", "yes", "y", "ok", "looks good", "send it"]:
            return CommandType.APPROVE, None

        if text in ["ignore", "skip", "no", "n", "pass", "not now", "later"]:
            return CommandType.IGNORE, None

        if text in ["rewrite", "try again", "redo"]:
            return CommandType.REWRITE, None

        if text in ["more", "show more", "full email", "details"]:
            return CommandType.MORE, None

        # "more [#]" - get details for numbered email from summary
        more_num_match = re.match(r"^more\s+(\d+)$", text, re.IGNORECASE)
        if more_num_match:
            return CommandType.MORE, more_num_match.group(1)

        if text == "spam":
            return CommandType.SPAM, None

        # "spam [#]" - mark numbered email as spam
        spam_num_match = re.match(r"^spam\s+(\d+)$", text, re.IGNORECASE)
        if spam_num_match:
            return CommandType.SPAM, spam_num_match.group(1)

        # Mute command - never show emails from this sender again
        # Supports: "mute", "mute sender@domain.com", "mute 3" (by number)
        mute_match = re.match(r"^mute\s*(.*)$", text, re.IGNORECASE)
        if mute_match:
            param = mute_match.group(1).strip() if mute_match.group(1) else None
            return CommandType.MUTE, param

        if text in ["done", "delete"]:
            return CommandType.DELETE, None

        # Archive all - acknowledge all emails in the morning summary
        if text in ["archive all", "archiveall", "ack all", "acknowledge all", "done all", "clear all"]:
            return CommandType.ARCHIVE_ALL, None

        # Spam batch commands (for spam digest, not morning summary)
        if text in ["dismiss all", "dismiss_all", "clear spam"]:
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

        # Follow-up command - "followup", "followup 3", "followup tomorrow", "followup 2d"
        followup_match = re.match(r"^(?:followup|follow up|remind|reminder)\s*(.*)$", text, re.IGNORECASE)
        if followup_match:
            param = followup_match.group(1).strip() if followup_match.group(1) else None
            return CommandType.FOLLOWUP, param

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
            "don't need notification", "dont need notification",
            # Common typos
            "dont need to be nitified", "don't need to be nitified",
            "dont. need", "dont need"
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
