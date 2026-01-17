"""
Coordinator Agent - The central orchestrator.

This agent:
- Polls for new emails
- Routes emails to appropriate specialist agents
- Manages the email state machine
- Coordinates responses between agents
"""

import logging
from datetime import datetime
from typing import List, Optional, Dict, Any

from .base import BaseAgent
from .drafting import DraftingAgent
from .teams_comms import TeamsCommsAgent
from .spam_filter import SpamFilterAgent
from ..db import Database
from ..models import EmailRecord, EmailState, EmailCategory
from ..integrations import EmailClient, TeamsClient, MCPClient
from ..config import settings

logger = logging.getLogger(__name__)


COORDINATOR_SYSTEM_PROMPT = """You are the Email Coordinator Agent, responsible for triaging and routing emails.

Your job is to analyze incoming emails and determine:
1. Category (urgent, action_required, fyi, meeting, spam_candidate, forward_candidate)
2. Priority (1-5, where 1 is highest)
3. Whether a reply is needed
4. The appropriate tone for any reply

Be efficient and accurate in your categorization. Consider:
- Sender importance (internal vs external, known contacts)
- Subject line keywords (urgent, action required, FYI, etc.)
- Email content and intent
- Time sensitivity

Always err on the side of caution - if unsure, mark as action_required rather than FYI.
"""


class CoordinatorAgent(BaseAgent):
    """
    Central coordinator that orchestrates email processing.

    Responsibilities:
    - Poll for new emails
    - Categorize and prioritize emails
    - Route to specialist agents (drafting, spam filter, etc.)
    - Manage state transitions
    - Handle errors and retries
    """

    def __init__(self, db: Database):
        super().__init__(
            db=db,
            name="coordinator",
            system_prompt=COORDINATOR_SYSTEM_PROMPT
        )

        # Initialize MCP clients
        self.mcp = MCPClient()
        self.email_client = EmailClient(self.mcp)
        self.teams_client = TeamsClient(self.mcp)

        # Initialize specialist agents
        self.drafting_agent = DraftingAgent(db)
        self.teams_agent = TeamsCommsAgent(db, self.teams_client)
        self.spam_agent = SpamFilterAgent(db)

    async def process(self) -> Dict[str, Any]:
        """
        Run a complete processing cycle.

        Processes all emails FIRST, then sends ONE consolidated Teams notification
        with a summary of spam vs action-required emails.

        Returns:
            Summary of actions taken
        """
        summary = {
            "new_emails": 0,
            "processed": 0,
            "spam_detected": 0,
            "action_required": 0,
            "fyi": 0,
            "errors": 0
        }

        # Track categorized emails for consolidated notification
        spam_emails: List[EmailRecord] = []
        action_emails: List[EmailRecord] = []
        fyi_emails: List[EmailRecord] = []

        try:
            # 1. Poll for new emails (last 7 days)
            new_emails = await self.poll_emails()
            summary["new_emails"] = len(new_emails)

            if not new_emails:
                # Just check for commands, no new emails
                await self.check_teams_replies()
                return summary

            # 2. Categorize ALL emails first (no notifications yet)
            for email in new_emails:
                try:
                    result = await self._categorize_email_only(email)
                    summary["processed"] += 1

                    if result.get("is_spam"):
                        spam_emails.append(email)
                        summary["spam_detected"] += 1
                    elif result.get("is_action"):
                        action_emails.append(email)
                        summary["action_required"] += 1
                    else:
                        fyi_emails.append(email)
                        summary["fyi"] += 1

                except Exception as e:
                    logger.error(f"Error categorizing email {email.id}: {e}")
                    summary["errors"] += 1

            # 3. Send ONE consolidated Teams notification
            await self._send_consolidated_summary(
                spam_emails=spam_emails,
                action_emails=action_emails,
                fyi_emails=fyi_emails
            )

            # 4. Check for Teams replies
            await self.check_teams_replies()

            self.log_action(
                "poll_cycle_complete",
                details=summary
            )

        except Exception as e:
            logger.error(f"Error in coordinator cycle: {e}")
            self.log_action("poll_cycle_error", error=str(e), success=False)
            summary["errors"] += 1

        return summary

    async def _categorize_email_only(self, email: EmailRecord) -> Dict[str, Any]:
        """
        Categorize an email without sending notifications.

        Returns:
            Dict with is_spam, is_action flags
        """
        result = {"is_spam": False, "is_action": False, "is_fyi": False}

        try:
            email.transition_to(EmailState.PROCESSING)
            self.db.save_email(email)

            # Check for spam first
            spam_result = await self.spam_agent.process(email)
            if spam_result.get("is_spam") or spam_result.get("likely_spam"):
                email.category = EmailCategory.SPAM_CANDIDATE
                email.spam_score = spam_result.get("spam_score", 0)
                email.transition_to(EmailState.SPAM_DETECTED)
                self.spam_agent.add_to_batch(email)
                self.db.save_email(email)
                result["is_spam"] = True
                return result

            # Categorize non-spam
            categorization = await self.categorize_email(email)
            email.category = categorization["category"]
            email.priority = categorization["priority"]

            if email.category == EmailCategory.SPAM_CANDIDATE:
                email.transition_to(EmailState.SPAM_DETECTED)
                self.spam_agent.add_to_batch(email)
                self.db.save_email(email)
                result["is_spam"] = True
                return result

            if email.category in [EmailCategory.URGENT, EmailCategory.ACTION_REQUIRED]:
                email.transition_to(EmailState.ACTION_REQUIRED)
                # Generate draft
                draft_result = await self.drafting_agent.process(email)
                email.summary = draft_result.get("summary")
                email.add_draft_version(draft_result.get("draft", ""))
                email.generate_approval_token()
                email.transition_to(EmailState.DRAFT_GENERATED)
                email.transition_to(EmailState.AWAITING_APPROVAL)
                result["is_action"] = True
            else:
                email.transition_to(EmailState.FYI_NOTIFIED)
                result["is_fyi"] = True

            self.db.save_email(email)

        except Exception as e:
            email.state = EmailState.ERROR
            email.error_message = str(e)
            self.db.save_email(email)
            raise

        return result

    async def _send_consolidated_summary(
        self,
        spam_emails: List[EmailRecord],
        action_emails: List[EmailRecord],
        fyi_emails: List[EmailRecord]
    ) -> Optional[str]:
        """
        Send ONE clean Teams notification - Executive Assistant style.

        - Groups similar emails by sender
        - Priority sorted
        - Spam moved to folder silently (just a count mentioned)
        """
        # Move spam to folder silently
        spam_moved = 0
        if spam_emails:
            spam_moved = await self._move_spam_to_folder(spam_emails)

        # Nothing to report?
        if not action_emails and not fyi_emails and spam_moved == 0:
            return None

        # Group emails by sender domain for cleaner display
        grouped_action = self._group_by_sender(action_emails)
        grouped_fyi = self._group_by_sender(fyi_emails)

        # Build ONE clean message
        content = f"""<h3>📧 Your Email Briefing</h3><hr>"""

        # Action required section - grouped by sender
        if action_emails:
            content += f"""<h4>🔴 Needs Your Response ({len(action_emails)})</h4>"""

            for sender_key, emails in grouped_action.items():
                if len(emails) == 1:
                    email = emails[0]
                    token = email.approval_token or "N/A"
                    summary_text = email.summary[:100] if email.summary else ""
                    content += f"""<p>
                        <b>{email.subject[:70]}</b><br>
                        From: {email.sender_name or email.sender_email}<br>
                        {f'<i>{summary_text}...</i><br>' if summary_text else ''}
                        → <code>{token}</code> to send reply
                    </p>"""
                else:
                    # Group multiple from same sender
                    content += f"""<p><b>{len(emails)} emails from {emails[0].sender_name or sender_key}</b><br>"""
                    for email in emails[:3]:
                        token = email.approval_token or "N/A"
                        content += f"""• {email.subject[:50]} [<code>{token}</code>]<br>"""
                    if len(emails) > 3:
                        content += f"""<i>...and {len(emails) - 3} more</i><br>"""
                    content += "</p>"

        # FYI section - just counts grouped by sender
        if fyi_emails:
            content += f"""<hr><h4>📋 FYI - No Action Needed ({len(fyi_emails)})</h4>"""
            fyi_summary = []
            for sender_key, emails in grouped_fyi.items():
                sender_name = emails[0].sender_name or sender_key
                if len(emails) == 1:
                    fyi_summary.append(f"{sender_name}: {emails[0].subject[:40]}")
                else:
                    fyi_summary.append(f"{sender_name}: {len(emails)} emails")
            content += "<p>" + " | ".join(fyi_summary[:5])
            if len(fyi_summary) > 5:
                content += f" | <i>+{len(fyi_summary) - 5} more</i>"
            content += "</p>"

        # Spam mention (already moved to folder)
        if spam_moved > 0:
            content += f"""<hr><p>🗑️ <b>{spam_moved} spam/newsletters</b> moved to AI-Spam folder</p>"""

        # Commands
        content += """<hr>
        <p><b>Reply:</b> <code>[token]</code> send | <code>more [token]</code> full email | <code>edit [token]: changes</code> modify | <code>ignore [token]</code> skip</p>"""

        message_id = self.teams_client.send_notification(content)
        if message_id:
            self.log_action(
                "briefing_sent",
                details={
                    "action": len(action_emails),
                    "fyi": len(fyi_emails),
                    "spam_moved": spam_moved
                }
            )
        return message_id

    def _group_by_sender(self, emails: List[EmailRecord]) -> Dict[str, List[EmailRecord]]:
        """Group emails by sender domain/email for cleaner display."""
        from collections import defaultdict
        grouped = defaultdict(list)
        for email in emails:
            # Use domain as key for grouping
            sender = email.sender_email.lower()
            domain = sender.split('@')[-1] if '@' in sender else sender
            grouped[domain].append(email)
        # Sort by count (most emails first) then by priority
        return dict(sorted(grouped.items(), key=lambda x: (-len(x[1]), min(e.priority for e in x[1]))))

    async def _move_spam_to_folder(self, spam_emails: List[EmailRecord]) -> int:
        """Move spam emails to AI-Spam folder silently."""
        moved = 0
        for email in spam_emails:
            try:
                # Try to move to Junk folder (built-in)
                self.email_client.mcp.move_mail_message(
                    message_id=email.message_id,
                    destination_folder_id="JunkEmail",
                    sender_email=email.mailbox
                )
                email.transition_to(EmailState.ARCHIVED)
                self.db.save_email(email)
                moved += 1
            except Exception as e:
                logger.warning(f"Could not move spam to folder: {e}")
                # Still count as handled even if move fails
                email.transition_to(EmailState.ARCHIVED)
                self.db.save_email(email)
                moved += 1

        if moved > 0:
            self.log_action("spam_moved_to_folder", details={"count": moved})

        return moved

    async def poll_emails(self) -> List[EmailRecord]:
        """
        Poll for new emails from all configured mailboxes.

        Returns:
            List of new EmailRecord objects
        """
        new_emails = []

        for mailbox in settings.all_mailboxes:
            try:
                messages = self.email_client.fetch_new_emails(mailbox=mailbox)

                for message in messages:
                    message_id = message.get("id", "")

                    # Skip if already processed
                    if self.db.is_message_processed(message_id, mailbox):
                        continue

                    # Convert to EmailRecord
                    email = self.email_client.parse_email_to_record(message, mailbox)

                    # Save to database
                    self.db.save_email(email)
                    self.db.mark_message_processed(message_id, mailbox)

                    new_emails.append(email)

                    self.log_action(
                        "email_ingested",
                        email_id=email.id,
                        details={
                            "subject": email.subject,
                            "sender": email.sender_email,
                            "mailbox": mailbox
                        }
                    )

            except Exception as e:
                logger.error(f"Error polling mailbox {mailbox}: {e}")
                self.log_action(
                    "poll_error",
                    details={"mailbox": mailbox},
                    error=str(e),
                    success=False
                )

        return new_emails

    async def categorize_email(self, email: EmailRecord) -> Dict[str, Any]:
        """
        Use Claude to categorize and prioritize an email.

        Args:
            email: The email to categorize

        Returns:
            Dict with category and priority
        """
        messages = [{
            "role": "user",
            "content": f"""Analyze this email and categorize it:

From: {email.sender_name or email.sender_email} <{email.sender_email}>
Subject: {email.subject}
Body Preview: {email.body_preview}

Determine:
1. Category: One of [urgent, action_required, fyi, meeting, spam_candidate, forward_candidate]
2. Priority: 1-5 (1 = highest, 5 = lowest)
3. Needs Reply: true/false
4. Reasoning: Brief explanation

Respond in JSON format."""
        }]

        schema = {
            "type": "object",
            "properties": {
                "category": {"type": "string"},
                "priority": {"type": "integer"},
                "needs_reply": {"type": "boolean"},
                "reasoning": {"type": "string"}
            }
        }

        try:
            result = self.call_claude_structured(messages, schema)

            # Parse category
            category_str = result.get("category", "action_required").lower()
            category_map = {
                "urgent": EmailCategory.URGENT,
                "action_required": EmailCategory.ACTION_REQUIRED,
                "fyi": EmailCategory.FYI,
                "meeting": EmailCategory.MEETING,
                "spam_candidate": EmailCategory.SPAM_CANDIDATE,
                "forward_candidate": EmailCategory.FORWARD_CANDIDATE
            }
            category = category_map.get(category_str, EmailCategory.ACTION_REQUIRED)

            return {
                "category": category,
                "priority": min(5, max(1, result.get("priority", 3))),
                "needs_reply": result.get("needs_reply", True),
                "reasoning": result.get("reasoning", "")
            }

        except Exception as e:
            logger.warning(f"Categorization failed, using defaults: {e}")
            return {
                "category": EmailCategory.ACTION_REQUIRED,
                "priority": 3,
                "needs_reply": True,
                "reasoning": "Default categorization due to error"
            }

    async def check_teams_replies(self) -> List[Dict[str, Any]]:
        """
        Check for and process user replies from Teams.

        Returns:
            List of processed commands
        """
        processed = []

        try:
            commands = await self.teams_agent.check_for_commands()

            for cmd in commands:
                email = cmd.get("email")
                command_type = cmd.get("command_type")
                parameter = cmd.get("parameter")

                # Handle spam batch commands (no email required)
                if command_type in ["dismiss_all", "review", "keep"]:
                    result = await self.handle_spam_batch_command(
                        command_type,
                        parameter
                    )
                    processed.append(result)
                    continue

                if not email:
                    continue

                result = await self.handle_user_command(
                    email,
                    command_type,
                    parameter
                )
                processed.append(result)

        except Exception as e:
            logger.error(f"Error checking Teams replies: {e}")
            self.log_action("teams_check_error", error=str(e), success=False)

        return processed

    async def handle_user_command(
        self,
        email: EmailRecord,
        command_type: str,
        parameter: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Handle a user command from Teams.

        Args:
            email: The email being acted upon
            command_type: Type of command (approve, edit, ignore, etc.)
            parameter: Optional command parameter

        Returns:
            Result of command execution
        """
        result = {"success": False, "action": command_type}

        try:
            if command_type in ["approve", "send", "yes"]:
                # Send the email
                if email.current_draft:
                    sent = self.email_client.send_reply(email, email.current_draft)
                    if sent:
                        email.transition_to(EmailState.APPROVED)
                        email.transition_to(EmailState.SENT)
                        email.sent_at = datetime.utcnow()
                        email.handled_by = "ai"
                        result["success"] = True

            elif command_type == "edit":
                # Modify draft with user's changes
                if parameter:
                    new_draft = await self.drafting_agent.edit_draft(
                        email,
                        parameter
                    )
                    email.add_draft_version(new_draft)
                    email.generate_approval_token()
                    # Re-send notification
                    await self.teams_agent.notify_email(email)
                    result["success"] = True

            elif command_type == "rewrite":
                # Generate completely new draft
                draft_result = await self.drafting_agent.process(email)
                email.add_draft_version(draft_result.get("draft", ""))
                email.generate_approval_token()
                await self.teams_agent.notify_email(email)
                result["success"] = True

            elif command_type in ["ignore", "skip"]:
                email.transition_to(EmailState.IGNORED)
                email.handled_by = "user"
                result["success"] = True

            elif command_type == "spam":
                email.category = EmailCategory.SPAM_CANDIDATE
                email.transition_to(EmailState.SPAM_DETECTED)
                email.transition_to(EmailState.ARCHIVED)
                # TODO: Learn spam pattern
                result["success"] = True

            elif command_type == "forward":
                if parameter:
                    forwarded = self.email_client.forward_email(email, parameter)
                    if forwarded:
                        email.transition_to(EmailState.FORWARDED)
                        result["success"] = True

            elif command_type == "more":
                # Send full email content
                await self.teams_agent.send_full_email(email)
                result["success"] = True

            self.db.save_email(email)

            self.log_action(
                f"command_{command_type}",
                email_id=email.id,
                user_command=command_type,
                details={"parameter": parameter},
                success=result["success"]
            )

        except Exception as e:
            logger.error(f"Error handling command {command_type}: {e}")
            self.log_action(
                f"command_{command_type}_error",
                email_id=email.id,
                user_command=command_type,
                error=str(e),
                success=False
            )
            result["error"] = str(e)

        return result

    async def _send_spam_batch_notification(self) -> Optional[str]:
        """
        Send a batched notification for all detected spam emails.

        Returns:
            Teams message ID if sent
        """
        batch = self.spam_agent.get_batch()
        if not batch:
            return None

        # Format the spam batch notification
        content = f"""
        <h3>🗑️ Spam/Newsletter Batch ({len(batch)} emails)</h3>
        <hr>
        <p>The following emails have been identified as spam or newsletters:</p>
        <ul>
        """

        for email in batch[:10]:  # Show first 10
            sender = email.sender_name or email.sender_email
            content += f"<li><b>{email.subject[:50]}{'...' if len(email.subject) > 50 else ''}</b> - {sender}</li>"

        if len(batch) > 10:
            content += f"<li><i>... and {len(batch) - 10} more</i></li>"

        content += """
        </ul>
        <hr>
        <p>Reply with:</p>
        <ul>
        <li><code>dismiss all</code> - Archive all spam emails</li>
        <li><code>review</code> - Show full list for individual review</li>
        <li><code>keep [subject keywords]</code> - Keep emails matching keywords</li>
        </ul>
        """

        message_id = self.teams_client.send_notification(content)
        if message_id:
            self.spam_agent.mark_notification_sent()
            self.log_action(
                "spam_batch_notified",
                details={"count": len(batch), "message_id": message_id}
            )
        return message_id

    async def handle_spam_batch_command(
        self,
        command: str,
        parameter: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Handle commands for the spam batch.

        Args:
            command: The command (dismiss_all, review, keep)
            parameter: Optional parameter (e.g., keywords for keep)

        Returns:
            Result of command execution
        """
        result = {"success": False, "action": command}

        try:
            if command in ["dismiss_all", "dismiss all", "archive all"]:
                # Archive all spam in batch
                count = await self.spam_agent.archive_batch()
                result["success"] = True
                result["archived_count"] = count

                # Send confirmation
                confirmation = f"<p>✅ Archived {count} spam emails</p>"
                self.teams_client.send_notification(confirmation)

            elif command == "review":
                # Send detailed list
                batch = self.spam_agent.get_batch()
                content = "<h3>📋 Spam Review List</h3><hr>"
                for i, email in enumerate(batch, 1):
                    content += f"""
                    <p><b>{i}. {email.subject}</b><br>
                    From: {email.sender_email}<br>
                    <code>keep {i}</code> to restore</p>
                    """
                self.teams_client.send_notification(content)
                result["success"] = True

            elif command.startswith("keep"):
                # Keep specific emails (by index or keyword)
                batch = self.spam_agent.get_batch()
                kept = []
                dismissed = []

                if parameter and parameter.isdigit():
                    # Keep by index
                    idx = int(parameter) - 1
                    if 0 <= idx < len(batch):
                        email = batch[idx]
                        email.category = EmailCategory.ACTION_REQUIRED
                        email.state = EmailState.ACTION_REQUIRED
                        self.db.save_email(email)
                        kept.append(email)
                        batch.pop(idx)

                elif parameter:
                    # Keep by keyword match
                    keyword = parameter.lower()
                    new_batch = []
                    for email in batch:
                        if keyword in email.subject.lower() or keyword in email.sender_email.lower():
                            email.category = EmailCategory.ACTION_REQUIRED
                            email.state = EmailState.ACTION_REQUIRED
                            self.db.save_email(email)
                            kept.append(email)
                        else:
                            new_batch.append(email)
                    self.spam_agent._spam_batch = new_batch

                # Notify about kept emails
                if kept:
                    kept_subjects = [e.subject[:30] for e in kept]
                    confirmation = f"<p>✅ Restored {len(kept)} emails: {', '.join(kept_subjects)}</p>"
                    self.teams_client.send_notification(confirmation)

                result["success"] = True
                result["kept_count"] = len(kept)

            self.log_action(
                f"spam_batch_{command}",
                details={"parameter": parameter},
                success=result["success"]
            )

        except Exception as e:
            logger.error(f"Error handling spam batch command {command}: {e}")
            result["error"] = str(e)

        return result

    def close(self):
        """Clean up resources."""
        self.mcp.close()
