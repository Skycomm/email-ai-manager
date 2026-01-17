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

    async def process(self) -> Dict[str, Any]:
        """
        Run a complete processing cycle.

        Returns:
            Summary of actions taken
        """
        summary = {
            "new_emails": 0,
            "processed": 0,
            "drafts_generated": 0,
            "notifications_sent": 0,
            "errors": 0
        }

        try:
            # 1. Poll for new emails
            new_emails = await self.poll_emails()
            summary["new_emails"] = len(new_emails)

            # 2. Process each new email
            for email in new_emails:
                try:
                    result = await self.process_email(email)
                    summary["processed"] += 1
                    if result.get("draft_generated"):
                        summary["drafts_generated"] += 1
                    if result.get("notification_sent"):
                        summary["notifications_sent"] += 1
                except Exception as e:
                    logger.error(f"Error processing email {email.id}: {e}")
                    summary["errors"] += 1

            # 3. Check for Teams replies
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

    async def process_email(self, email: EmailRecord) -> Dict[str, Any]:
        """
        Process a single email through the pipeline.

        Args:
            email: The email to process

        Returns:
            Processing result summary
        """
        result = {
            "draft_generated": False,
            "notification_sent": False
        }

        try:
            # Transition to processing
            email.transition_to(EmailState.PROCESSING)
            self.db.save_email(email)

            # 1. Categorize and prioritize
            categorization = await self.categorize_email(email)
            email.category = categorization["category"]
            email.priority = categorization["priority"]

            # 2. Route based on category
            if email.category == EmailCategory.SPAM_CANDIDATE:
                # TODO: Route to spam filter agent
                email.transition_to(EmailState.SPAM_DETECTED)

            elif email.category == EmailCategory.FYI:
                # Just notify, no reply needed
                email.transition_to(EmailState.FYI_NOTIFIED)

            elif email.category in [EmailCategory.URGENT, EmailCategory.ACTION_REQUIRED]:
                # Generate draft and request approval
                email.transition_to(EmailState.ACTION_REQUIRED)

                # Generate summary and draft
                draft_result = await self.drafting_agent.process(email)
                email.summary = draft_result.get("summary")
                email.add_draft_version(draft_result.get("draft", ""))
                email.generate_approval_token()

                email.transition_to(EmailState.DRAFT_GENERATED)
                email.transition_to(EmailState.AWAITING_APPROVAL)

                result["draft_generated"] = True

            elif email.category == EmailCategory.MEETING:
                # Calendar-related, may need response
                email.transition_to(EmailState.ACTION_REQUIRED)

            elif email.category == EmailCategory.FORWARD_CANDIDATE:
                email.transition_to(EmailState.FORWARD_SUGGESTED)

            # 3. Send Teams notification
            teams_msg_id = await self.teams_agent.notify_email(email)
            if teams_msg_id:
                email.teams_message_id = teams_msg_id
                result["notification_sent"] = True

            # Save final state
            self.db.save_email(email)

            self.log_action(
                "email_processed",
                email_id=email.id,
                details={
                    "category": email.category.value if email.category else None,
                    "priority": email.priority,
                    "state": email.state.value
                }
            )

        except Exception as e:
            email.state = EmailState.ERROR
            email.error_message = str(e)
            email.retry_count += 1
            self.db.save_email(email)

            self.log_action(
                "email_processing_error",
                email_id=email.id,
                error=str(e),
                success=False
            )
            raise

        return result

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

    def close(self):
        """Clean up resources."""
        self.mcp.close()
