"""
Email-specific MCP operations.

Higher-level wrapper around MCP client for email operations.
"""

import logging
from datetime import datetime, timedelta
from typing import List, Optional, Dict, Any

from .mcp_client import MCPClient, MCPClientError
from ..models import EmailRecord
from ..config import settings

logger = logging.getLogger(__name__)


class EmailClient:
    """High-level email operations using MCP."""

    def __init__(self, mcp_client: Optional[MCPClient] = None):
        self.mcp = mcp_client or MCPClient()

    def fetch_new_emails(
        self,
        mailbox: Optional[str] = None,
        since_days: int = 7,
        max_emails: int = 50
    ) -> List[Dict[str, Any]]:
        """
        Fetch new emails from the mailbox.

        Args:
            mailbox: Email address of mailbox (uses default if not specified)
            since_days: Only fetch emails from the last N days
            max_emails: Maximum number of emails to fetch

        Returns:
            List of email message dictionaries
        """
        mailbox = mailbox or settings.mailbox_email

        try:
            # Calculate date filter for last N days
            since_date = datetime.utcnow() - timedelta(days=since_days)
            filter_query = f"receivedDateTime ge {since_date.strftime('%Y-%m-%dT%H:%M:%SZ')}"

            # Fetch recent emails from inbox
            messages = self.mcp.list_mail_messages(
                mailbox=mailbox,
                folder="inbox",
                top=max_emails,
                filter_query=filter_query
            )

            logger.info(f"Fetched {len(messages)} emails from last {since_days} days from {mailbox}")
            return messages

        except MCPClientError as e:
            logger.error(f"Failed to fetch emails: {e}")
            return []

    def get_email_details(
        self,
        message_id: str,
        mailbox: Optional[str] = None
    ) -> Optional[Dict[str, Any]]:
        """
        Get full details of an email.

        Args:
            message_id: MS365 message ID
            mailbox: Mailbox email address

        Returns:
            Full email message or None if not found
        """
        mailbox = mailbox or settings.mailbox_email

        try:
            return self.mcp.get_mail_message(message_id, mailbox)
        except MCPClientError as e:
            logger.error(f"Failed to get email {message_id}: {e}")
            return None

    def send_reply(
        self,
        email: EmailRecord,
        reply_body: str,
        reply_all: bool = False
    ) -> bool:
        """
        Send a reply to an email.

        Args:
            email: The email record to reply to
            reply_body: The reply content (HTML)
            reply_all: Whether to reply to all recipients

        Returns:
            True if sent successfully
        """
        try:
            # Format recipients
            to_recipients = [{"emailAddress": {"address": email.sender_email}}]

            if reply_all:
                # Add other recipients
                for recipient in email.to_recipients:
                    if recipient != settings.mailbox_email:
                        to_recipients.append({"emailAddress": {"address": recipient}})

            # Build reply subject
            subject = email.subject
            if not subject.lower().startswith("re:"):
                subject = f"Re: {subject}"

            # Send the reply
            result = self.mcp.send_mail(
                to=to_recipients,
                subject=subject,
                body=reply_body,
                body_type="HTML",
                sender_email=email.mailbox
            )

            logger.info(f"Sent reply to {email.sender_email} for '{email.subject}'")
            return True

        except MCPClientError as e:
            logger.error(f"Failed to send reply: {e}")
            return False

    def forward_email(
        self,
        email: EmailRecord,
        forward_to: str,
        comment: Optional[str] = None
    ) -> bool:
        """
        Forward an email to another recipient.

        Args:
            email: The email to forward
            forward_to: Email address to forward to
            comment: Optional comment to add

        Returns:
            True if forwarded successfully
        """
        try:
            # Build forward body
            forward_body = ""
            if comment:
                forward_body = f"<p>{comment}</p><hr>"

            forward_body += f"""
            <p><b>---------- Forwarded message ----------</b></p>
            <p><b>From:</b> {email.sender_name or email.sender_email} &lt;{email.sender_email}&gt;</p>
            <p><b>Date:</b> {email.received_at.strftime('%Y-%m-%d %H:%M')}</p>
            <p><b>Subject:</b> {email.subject}</p>
            <hr>
            {email.body_full or email.body_preview}
            """

            # Send forwarded email
            result = self.mcp.send_mail(
                to=[{"emailAddress": {"address": forward_to}}],
                subject=f"Fwd: {email.subject}",
                body=forward_body,
                body_type="HTML",
                sender_email=email.mailbox
            )

            logger.info(f"Forwarded email '{email.subject}' to {forward_to}")
            return True

        except MCPClientError as e:
            logger.error(f"Failed to forward email: {e}")
            return False

    def archive_email(
        self,
        message_id: str,
        mailbox: Optional[str] = None
    ) -> bool:
        """
        Move an email to archive.

        Args:
            message_id: MS365 message ID
            mailbox: Mailbox email address

        Returns:
            True if archived successfully
        """
        # This would use the move-mail-message MCP tool
        # Implementation depends on MCP server capabilities
        logger.info(f"Archive email {message_id} (not yet implemented)")
        return True

    def delete_email(
        self,
        message_id: str,
        mailbox: Optional[str] = None
    ) -> bool:
        """
        Delete an email.

        Args:
            message_id: MS365 message ID
            mailbox: Mailbox email address

        Returns:
            True if deleted successfully
        """
        # This would use the delete-mail-message MCP tool
        logger.info(f"Delete email {message_id} (not yet implemented)")
        return True

    def parse_email_to_record(
        self,
        message: Dict[str, Any],
        mailbox: str
    ) -> EmailRecord:
        """
        Convert an MCP email message to an EmailRecord.

        Args:
            message: Raw email message from MCP
            mailbox: Mailbox this came from

        Returns:
            EmailRecord instance
        """
        # Extract sender info
        sender = message.get("from", {}).get("emailAddress", {})
        sender_email = sender.get("address", "unknown@unknown.com")
        sender_name = sender.get("name")

        # Extract recipients
        to_recipients = [
            r.get("emailAddress", {}).get("address", "")
            for r in message.get("toRecipients", [])
        ]
        cc_recipients = [
            r.get("emailAddress", {}).get("address", "")
            for r in message.get("ccRecipients", [])
        ]

        # Parse received time
        received_str = message.get("receivedDateTime", "")
        try:
            received_at = datetime.fromisoformat(received_str.replace("Z", "+00:00"))
        except (ValueError, TypeError):
            received_at = datetime.utcnow()

        # Extract body
        body = message.get("body", {})
        body_content = body.get("content", "")
        body_preview = message.get("bodyPreview", body_content[:500])

        return EmailRecord.create(
            message_id=message.get("id", ""),
            mailbox=mailbox,
            thread_id=message.get("conversationId"),
            sender_email=sender_email,
            sender_name=sender_name,
            to_recipients=to_recipients,
            cc_recipients=cc_recipients,
            subject=message.get("subject", "(No Subject)"),
            body_preview=body_preview,
            body_full=body_content,
            received_at=received_at,
            has_attachments=message.get("hasAttachments", False),
            importance=message.get("importance", "normal"),
        )

    def close(self):
        """Close underlying connections."""
        self.mcp.close()
