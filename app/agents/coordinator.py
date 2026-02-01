"""
Coordinator Agent - The central orchestrator.

This agent:
- Polls for new emails
- Routes emails to appropriate specialist agents
- Manages the email state machine
- Coordinates responses between agents
"""

import logging
import re
from collections import defaultdict
from datetime import datetime, date, timedelta, timezone
from typing import List, Optional, Dict, Any
from zoneinfo import ZoneInfo

from .base import BaseAgent
from .drafting import DraftingAgent
from .teams_comms import TeamsCommsAgent
from .spam_filter import SpamFilterAgent
from .calendar import CalendarAgent, is_meeting_email
from .rules import RulesAgent
from ..db import Database
from ..models import EmailRecord, EmailState, EmailCategory, SpamRule, RuleAction
from ..integrations import EmailClient, TeamsClient, MCPClient
from ..config import settings

logger = logging.getLogger(__name__)


def to_local_time(dt: Optional[datetime]) -> Optional[datetime]:
    """Convert a datetime to the configured local timezone (default Perth, UTC+8)."""
    if dt is None:
        return None
    try:
        tz = ZoneInfo(settings.timezone)
        # If dt is naive (no timezone), assume it's UTC
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(tz)
    except Exception:
        # Fallback: just add 8 hours for Perth if zoneinfo fails
        if dt.tzinfo is None:
            return dt + timedelta(hours=8)
        return dt


def format_local_time(dt: Optional[datetime], fmt: str = '%Y-%m-%d %H:%M') -> str:
    """Format a datetime in the local timezone."""
    if dt is None:
        return 'Unknown'
    local_dt = to_local_time(dt)
    return local_dt.strftime(fmt) if local_dt else 'Unknown'


def get_outlook_deep_link(message_id: str) -> str:
    """Generate an Outlook web deep link for a specific email."""
    # URL-encode the message ID for the deep link
    import urllib.parse
    encoded_id = urllib.parse.quote(message_id, safe='')
    return f"https://outlook.office365.com/mail/inbox/id/{encoded_id}"


def is_vip_sender(email: EmailRecord) -> bool:
    """Check if sender is a VIP based on config."""
    sender_email = email.sender_email.lower()
    sender_domain = sender_email.split('@')[-1] if '@' in sender_email else ""

    # Check VIP email addresses
    for vip_email in settings.vip_senders:
        if vip_email.lower() == sender_email:
            return True

    # Check VIP domains
    for vip_domain in settings.vip_domains:
        if vip_domain.lower() == sender_domain:
            return True

    return False


def is_internal_sender(email: EmailRecord) -> bool:
    """Check if sender is from an internal domain."""
    sender_email = email.sender_email.lower()
    sender_domain = sender_email.split('@')[-1] if '@' in sender_email else ""

    for internal_domain in settings.internal_domains:
        if internal_domain.lower() == sender_domain:
            return True

    return False


def is_alert_email(email: EmailRecord) -> bool:
    """
    Check if an email is an alert/monitoring notification that needs immediate attention.

    Alert emails from monitoring systems (Meraki, UptimeRobot, etc.) or with alert-related
    subjects get sent immediately with deduplication, while regular FYI emails wait for
    the morning summary.
    """
    sender_email = email.sender_email.lower()
    sender_domain = sender_email.split('@')[-1] if '@' in sender_email else ""
    subject_lower = email.subject.lower()

    # Check if sender domain is an alert/monitoring service
    for alert_domain in settings.alert_sender_domains:
        if alert_domain.lower() in sender_domain:
            return True

    # Check if subject contains alert patterns
    for pattern in settings.alert_subject_patterns:
        if pattern.lower() in subject_lower:
            return True

    return False


def check_auto_send_eligible(email: EmailRecord) -> bool:
    """
    Check if an email reply can be auto-sent without approval.

    Requirements:
    - auto_send_enabled must be True
    - Category must be in auto_send_categories
    - Priority must be >= auto_send_max_priority (lower priority = higher number)
    - If auto_send_internal_only, sender must be internal
    - If external_domain_require_approval, external recipients always need approval
    """
    if not settings.auto_send_enabled:
        return False

    # External domains always require approval if setting is enabled
    if settings.external_domain_require_approval and not is_internal_sender(email):
        return False

    # Check category
    if email.category:
        category_value = email.category.value
        if category_value not in settings.auto_send_categories:
            return False
    else:
        return False

    # Check priority (higher number = lower priority = safer to auto-send)
    if email.priority < settings.auto_send_max_priority:
        return False

    # Check internal-only restriction (redundant if external_domain_require_approval is True, but kept for clarity)
    if settings.auto_send_internal_only and not is_internal_sender(email):
        return False

    return True


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
        self.teams_client = TeamsClient(self.mcp, db=db)  # Pass DB for notification deduplication

        # Initialize specialist agents
        self.drafting_agent = DraftingAgent(db)
        self.teams_agent = TeamsCommsAgent(db, self.teams_client)
        self.spam_agent = SpamFilterAgent(db)
        self.calendar_agent = CalendarAgent(db, self.mcp)
        self.rules_agent = RulesAgent(db)

        # Mapping for numbered email references in summaries (e.g., "more 3")
        # Maps number -> email_id for the current summary
        # Now persisted in database for restart resilience
        self.summary_email_mapping: Dict[int, str] = self.db.get_summary_mapping()

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
            "spam_deleted": 0,
            "newsletters": 0,
            "action_required": 0,
            "fyi": 0,
            "auto_sent": 0,
            "muted": 0,
            "errors": 0
        }

        # Track categorized emails for notifications
        hard_spam_emails: List[EmailRecord] = []  # Auto-delete, no notification
        newsletter_emails: List[EmailRecord] = []  # Group in ONE message
        action_emails: List[EmailRecord] = []  # ONE message per email
        fyi_emails: List[EmailRecord] = []  # Grouped summary
        auto_sent_emails: List[EmailRecord] = []

        try:
            # 1. Poll for new emails (last 7 days)
            new_emails = await self.poll_emails()
            summary["new_emails"] = len(new_emails)

            # 1b. Also get any existing emails stuck in 'new' state (reprocessing)
            pending_new = self.db.get_emails_by_state(EmailState.NEW, limit=50)
            new_email_ids = {e.id for e in new_emails}
            reprocess_emails = [e for e in pending_new if e.id not in new_email_ids]

            if reprocess_emails:
                logger.info(f"Found {len(reprocess_emails)} emails to reprocess from 'new' state")
                new_emails.extend(reprocess_emails)
                summary["new_emails"] = len(new_emails)

            if not new_emails:
                # Even with no new emails, check for morning summary and Teams replies
                logger.info("No new emails, checking morning summary...")
                sent_summary = await self._check_morning_summary()
                if sent_summary:
                    logger.info("Morning summary was sent")
                await self.check_teams_replies()
                return summary

            # 2. Categorize ALL emails first (no notifications yet)
            for email in new_emails:
                try:
                    # Check if sender is muted first
                    if self.db.is_sender_muted(email.sender_email):
                        email.transition_to(EmailState.ARCHIVED)
                        self.db.save_email(email)
                        summary["muted"] += 1
                        logger.info(f"Muted sender, auto-archived: {email.sender_email}")
                        continue

                    result = await self._categorize_email_only(email)
                    summary["processed"] += 1

                    if result.get("is_hard_spam"):
                        hard_spam_emails.append(email)
                        summary["spam_deleted"] += 1
                    elif result.get("is_newsletter"):
                        newsletter_emails.append(email)
                        summary["newsletters"] += 1
                    elif result.get("is_auto_sent"):
                        auto_sent_emails.append(email)
                        summary["auto_sent"] += 1
                    elif result.get("is_action"):
                        action_emails.append(email)
                        summary["action_required"] += 1
                    else:
                        fyi_emails.append(email)
                        summary["fyi"] += 1

                except Exception as e:
                    logger.error(f"Error categorizing email {email.id}: {e}")
                    summary["errors"] += 1

            # 3. Handle spam - auto-delete (no notification)
            if hard_spam_emails:
                await self._delete_spam(hard_spam_emails)

            # 4. Send notifications based on new structure:
            #    - Newsletters: ONE grouped message
            #    - Action items: ONE message per email (for reply threading)
            #    - FYI: Grouped summary
            await self._send_notifications(
                newsletter_emails=newsletter_emails,
                action_emails=action_emails,
                fyi_emails=fyi_emails,
                auto_sent_emails=auto_sent_emails
            )

            # 5. Check for Teams replies
            await self.check_teams_replies()

            # 6. Check for pending follow-up reminders
            await self.check_followup_reminders()

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
            Dict with is_hard_spam, is_newsletter, is_action, is_auto_sent flags
        """
        result = {
            "is_hard_spam": False,  # Delete without notification
            "is_newsletter": False,  # Group in one message
            "is_action": False,
            "is_fyi": False,
            "is_auto_sent": False,
            "is_meeting": False,
            "meeting_info": None
        }

        try:
            email.transition_to(EmailState.PROCESSING)

            # Check VIP status early
            email.is_vip = is_vip_sender(email)
            if email.is_vip:
                logger.info(f"VIP sender detected: {email.sender_email}")

            # Fetch thread context if this is part of a conversation
            if email.thread_id:
                email.thread_context = await self._fetch_thread_context(email)

            self.db.save_email(email)

            # Check for spam/newsletter first (VIP senders skip spam check)
            if not email.is_vip:
                spam_result = await self.spam_agent.process(email)

                # Newsletter - group in one message
                if spam_result.get("is_newsletter"):
                    email.category = EmailCategory.NEWSLETTER
                    email.spam_score = spam_result.get("spam_score", 0)
                    email.transition_to(EmailState.FYI_NOTIFIED)
                    self.db.save_email(email)
                    result["is_newsletter"] = True
                    return result

                # Hard spam - delete without notification
                if spam_result.get("is_spam"):
                    email.category = EmailCategory.SPAM_CANDIDATE
                    email.spam_score = spam_result.get("spam_score", 0)
                    email.transition_to(EmailState.SPAM_DETECTED)
                    self.db.save_email(email)
                    result["is_hard_spam"] = True
                    return result

            # Check for meeting emails
            if is_meeting_email(email):
                meeting_result = await self.calendar_agent.process(email)
                result["is_meeting"] = meeting_result.get("is_meeting", False)
                result["meeting_info"] = meeting_result

                if result["is_meeting"]:
                    email.category = EmailCategory.MEETING

                    # If meeting was auto-responded, treat as handled
                    if meeting_result.get("auto_responded"):
                        email.transition_to(EmailState.SENT)
                        email.handled_by = "ai_auto"
                        result["is_auto_sent"] = True
                        self.db.save_email(email)
                        return result

                    # Generate meeting-specific draft
                    if meeting_result.get("meeting_type") == "invite":
                        draft = self.calendar_agent.suggest_meeting_response(
                            email,
                            {"has_conflict": meeting_result.get("has_conflict"),
                             "conflicting_events": meeting_result.get("conflict_details")}
                        )
                        email.summary = f"Meeting invite: {meeting_result.get('suggested_action', 'review')}"
                        if meeting_result.get("has_conflict"):
                            email.summary += f" (conflict detected)"
                        email.add_draft_version(draft)
                        email.generate_approval_token()
                        email.transition_to(EmailState.ACTION_REQUIRED)
                        email.transition_to(EmailState.DRAFT_GENERATED)
                        email.transition_to(EmailState.AWAITING_APPROVAL)
                        result["is_action"] = True
                        self.db.save_email(email)
                        return result
                    else:
                        # Meeting response or update - FYI
                        email.transition_to(EmailState.FYI_NOTIFIED)
                        result["is_fyi"] = True
                        self.db.save_email(email)
                        return result

            # Categorize non-spam, non-meeting
            categorization = await self.categorize_email(email)
            email.category = categorization["category"]
            email.priority = categorization["priority"]

            # VIP senders get priority boost
            if email.is_vip and email.priority > 1:
                email.priority = max(1, email.priority - 1)
                logger.info(f"VIP priority boost: {email.priority}")

            if email.category == EmailCategory.SPAM_CANDIDATE:
                email.transition_to(EmailState.SPAM_DETECTED)
                self.spam_agent.add_to_batch(email)
                self.db.save_email(email)
                result["is_spam"] = True
                return result

            if email.category in [EmailCategory.URGENT, EmailCategory.ACTION_REQUIRED]:
                email.transition_to(EmailState.ACTION_REQUIRED)
                # Generate draft with thread context
                draft_result = await self.drafting_agent.process(email)
                email.summary = draft_result.get("summary")
                email.add_draft_version(draft_result.get("draft", ""))
                email.generate_approval_token()
                email.transition_to(EmailState.DRAFT_GENERATED)

                # NEVER auto-send - always require manual approval via web dashboard
                email.auto_send_eligible = False
                email.transition_to(EmailState.AWAITING_APPROVAL)
                result["is_action"] = True
            else:
                # Generate summary for FYI emails too (but no draft)
                try:
                    summary_result = await self.drafting_agent.generate_summary_only(email)
                    email.summary = summary_result.get("summary")
                except Exception as e:
                    logger.warning(f"Could not generate FYI summary: {e}")
                email.transition_to(EmailState.FYI_NOTIFIED)
                result["is_fyi"] = True

            self.db.save_email(email)

            # Apply LLM-based routing rules (move to folder, etc.)
            try:
                rule_result = await self._apply_email_rules(email)
                if rule_result.get("matched"):
                    logger.info(f"Email matched rule: {rule_result.get('rule_name')} -> {rule_result.get('action')}")
            except Exception as e:
                logger.warning(f"Error applying email rules: {e}")

        except Exception as e:
            email.state = EmailState.ERROR
            email.error_message = str(e)
            self.db.save_email(email)
            raise

        return result

    async def _apply_email_rules(self, email: EmailRecord) -> Dict[str, Any]:
        """
        Apply LLM-based email rules (move to folder, forward, etc.).

        Args:
            email: The email to process

        Returns:
            Dict with matched rule info and action taken
        """
        result = {"matched": False}

        try:
            # Get matching rules from the rules agent
            matches = await self.rules_agent.evaluate_all_rules(email, min_confidence=60)

            if not matches:
                return result

            # Apply the first matching rule (highest priority)
            rule, match_result = matches[0]
            result["matched"] = True
            result["rule_name"] = rule.name
            result["confidence"] = match_result["confidence"]
            result["action"] = rule.action.value

            # Execute the rule action
            if rule.action == RuleAction.MOVE_TO_FOLDER:
                # Move email to the specified folder in MS365
                success = self._move_email_to_folder(email, rule.action_value)
                result["action_success"] = success
                if success:
                    logger.info(f"Moved email '{email.subject[:40]}' to folder: {rule.action_value}")

            elif rule.action == RuleAction.ARCHIVE:
                # Archive the email immediately
                email.transition_to(EmailState.ARCHIVED)
                email.handled_by = "rule"
                self.db.save_email(email)
                result["action_success"] = True

            elif rule.action == RuleAction.FORWARD:
                # Forward the email
                forwarded = self.email_client.forward_email(email, rule.action_value)
                result["action_success"] = forwarded
                if forwarded:
                    email.transition_to(EmailState.FORWARDED)
                    self.db.save_email(email)

            elif rule.action == RuleAction.SET_PRIORITY:
                # Override email priority
                try:
                    email.priority = int(rule.action_value)
                    self.db.save_email(email)
                    result["action_success"] = True
                except ValueError:
                    result["action_success"] = False

            elif rule.action == RuleAction.NOTIFY:
                # Send a custom Teams notification
                self.teams_client.send_notification(
                    f"<p>📧 <b>Rule Match:</b> {rule.name}</p>"
                    f"<p><b>Email:</b> {email.subject}</p>"
                    f"<p>{rule.action_value}</p>"
                )
                result["action_success"] = True

            self.log_action(
                "rule_applied",
                email_id=email.id,
                details={
                    "rule_id": rule.id,
                    "rule_name": rule.name,
                    "action": rule.action.value,
                    "action_value": rule.action_value,
                    "confidence": match_result["confidence"]
                }
            )

        except Exception as e:
            logger.error(f"Error applying email rules: {e}")
            result["error"] = str(e)

        return result

    def _move_email_to_folder(self, email: EmailRecord, folder_name: str) -> bool:
        """
        Move an email to a specific folder in MS365.

        Args:
            email: The email to move
            folder_name: Name of the destination folder

        Returns:
            True if successful, False otherwise
        """
        try:
            # First, try to find or create the folder
            folder_id = self._get_or_create_folder(email.mailbox, folder_name)
            if not folder_id:
                logger.warning(f"Could not find or create folder: {folder_name}")
                return False

            # Move the email to the folder
            self.mcp.move_mail_message(
                message_id=email.message_id,
                destination_folder_id=folder_id,
                sender_email=email.mailbox
            )
            return True

        except Exception as e:
            logger.error(f"Failed to move email to folder {folder_name}: {e}")
            return False

    def _get_or_create_folder(self, mailbox: str, folder_name: str) -> Optional[str]:
        """
        Get folder ID by name, creating it if it doesn't exist.

        Args:
            mailbox: The mailbox email
            folder_name: Name of the folder

        Returns:
            Folder ID or None if failed
        """
        try:
            # List existing folders
            folders = self.mcp.list_mail_folders(sender_email=mailbox)

            # Look for matching folder
            for folder in folders:
                if folder.get("displayName", "").lower() == folder_name.lower():
                    return folder.get("id")

            # Folder doesn't exist - create it
            # Note: The MCP server may not support folder creation
            # In that case, we'll log a warning and return None
            logger.warning(f"Folder '{folder_name}' not found. Please create it manually in Outlook.")
            return None

        except Exception as e:
            logger.error(f"Error getting folder {folder_name}: {e}")
            return None

    async def _fetch_thread_context(self, email: EmailRecord) -> Optional[str]:
        """
        Fetch conversation history for thread-aware responses.

        Args:
            email: Email with thread_id

        Returns:
            Summary of previous messages in thread, or None
        """
        if not email.thread_id:
            return None

        try:
            # Get previous messages in this conversation
            thread_messages = self.mcp.get_conversation_messages(
                mailbox=email.mailbox,
                conversation_id=email.thread_id,
                top=5
            )

            if not thread_messages or len(thread_messages) <= 1:
                return None

            # Build context summary (excluding current email)
            context_parts = []
            for msg in thread_messages:
                msg_id = msg.get("id", "")
                if msg_id == email.message_id:
                    continue  # Skip current email

                sender = msg.get("from", {}).get("emailAddress", {})
                sender_name = sender.get("name") or sender.get("address", "Unknown")
                subject = msg.get("subject", "")
                preview = msg.get("bodyPreview", "")[:150]

                context_parts.append(f"- {sender_name}: {preview}...")

            if not context_parts:
                return None

            thread_context = f"Previous messages in this thread ({len(context_parts)}):\n" + "\n".join(context_parts[:3])
            logger.debug(f"Thread context for {email.subject}: {len(context_parts)} messages")
            return thread_context

        except Exception as e:
            logger.warning(f"Could not fetch thread context: {e}")
            return None

    async def _delete_spam(self, spam_emails: List[EmailRecord]) -> int:
        """Delete hard spam emails - no notification."""
        deleted = 0
        for email in spam_emails:
            try:
                # Move to Deleted Items folder
                self.email_client.mcp.move_mail_message(
                    message_id=email.message_id,
                    destination_folder_id="DeletedItems",
                    sender_email=email.mailbox
                )
                email.transition_to(EmailState.ARCHIVED)
                self.db.save_email(email)
                deleted += 1
                logger.info(f"Deleted spam: {email.subject[:50]} from {email.sender_email}")
            except Exception as e:
                logger.warning(f"Could not delete spam: {e}")
                email.transition_to(EmailState.ARCHIVED)
                self.db.save_email(email)
                deleted += 1

        if deleted > 0:
            self.log_action("spam_deleted", details={"count": deleted})
        return deleted

    def _is_morning_summary_time(self) -> bool:
        """
        Check if it's time to send the morning summary.
        Returns True if current hour >= configured morning hour AND
        we haven't sent a summary today yet.

        Uses the configured timezone (default Australia/Perth) not UTC.
        """
        # Get current time in configured timezone
        now_local = to_local_time(datetime.now(timezone.utc))
        morning_hour = settings.teams_morning_summary_hour

        if now_local.hour < morning_hour:
            return False

        # Check if we already sent today's summary (use local date)
        today_str = now_local.date().isoformat()
        last_summary = self.db.get_setting("last_morning_summary_date")

        if last_summary == today_str:
            return False  # Already sent today

        return True

    def _mark_morning_summary_sent(self):
        """Mark that we've sent today's morning summary."""
        # Use local date for tracking
        now_local = to_local_time(datetime.now(timezone.utc))
        today_str = now_local.date().isoformat()
        self.db.set_setting("last_morning_summary_date", today_str)

    def _auto_archive_old_fyi(self) -> int:
        """
        Auto-archive old FYI/newsletter emails.
        Called after morning summary to clean up old emails.
        Returns the count of archived emails.
        """
        archive_hours = settings.fyi_auto_archive_hours
        return self.db.archive_old_fyi_emails(older_than_hours=archive_hours)

    async def _check_morning_summary(self) -> bool:
        """
        Check if it's time to send the morning summary and send it if so.
        Called on every poll cycle regardless of whether there are new emails.

        Returns:
            True if morning summary was sent
        """
        if not self._is_morning_summary_time():
            return False

        # Get ALL FYI emails from last 24 hours for the morning summary
        all_fyi_24h = self.db.get_fyi_emails_last_24h(limit=50)

        # Separate newsletters from regular FYI
        pending_newsletters = [e for e in all_fyi_24h if e.category == EmailCategory.NEWSLETTER]
        pending_fyi = [e for e in all_fyi_24h if e.category != EmailCategory.NEWSLETTER]

        # Get auto-sent emails from last 24 hours
        all_auto_sent = self.db.get_auto_sent_emails_last_24h(limit=20)

        if pending_newsletters or pending_fyi or all_auto_sent:
            logger.info(f"Sending morning summary: {len(pending_newsletters)} newsletters, {len(pending_fyi)} fyi, {len(all_auto_sent)} auto-sent")
            await self._send_morning_summary(pending_newsletters, pending_fyi, all_auto_sent)
            self._mark_morning_summary_sent()

            # Auto-archive old FYI emails after morning summary
            archived_count = self._auto_archive_old_fyi()
            if archived_count > 0:
                logger.info(f"Auto-archived {archived_count} old FYI/newsletter emails")

            return True

        # No emails to summarize, but still mark as sent so we don't keep checking
        logger.info("Morning summary time but no emails to summarize")
        self._mark_morning_summary_sent()

        # Still run auto-archive even if no summary sent
        archived_count = self._auto_archive_old_fyi()
        if archived_count > 0:
            logger.info(f"Auto-archived {archived_count} old FYI/newsletter emails")

        return False

    async def _send_notifications(
        self,
        newsletter_emails: List[EmailRecord],
        action_emails: List[EmailRecord],
        fyi_emails: List[EmailRecord],
        auto_sent_emails: Optional[List[EmailRecord]] = None
    ) -> None:
        """
        Send Teams notifications:
        - Action items: Immediately, ONE message per email (for reply threading)
        - Newsletters/FYI/Auto-sent: Grouped in morning summary at 7am
        """
        auto_sent_emails = auto_sent_emails or []

        # 1. Send ONE message per action-required email IMMEDIATELY
        for email in sorted(action_emails, key=lambda e: e.priority):
            await self._send_action_email_notification(email)

        # 2. Check if it's time for the morning summary (newsletters, FYI, auto-sent)
        # This uses the dedicated method which queries last 24h of emails
        await self._check_morning_summary()

        # 3. ALL FYI/newsletter emails wait for morning summary - no immediate notifications
        # Only Important/Action emails get sent immediately (handled above)
        held_count = len(newsletter_emails) + len(fyi_emails) + len(auto_sent_emails)
        if held_count > 0:
            logger.info(f"Holding {held_count} emails for morning summary (newsletters: {len(newsletter_emails)}, fyi: {len(fyi_emails)}, auto-sent: {len(auto_sent_emails)})")

    async def _send_morning_summary(
        self,
        newsletter_emails: List[EmailRecord],
        fyi_emails: List[EmailRecord],
        auto_sent_emails: List[EmailRecord]
    ) -> Optional[str]:
        """
        Send the 7am morning summary with all non-action emails.
        Combines newsletters, FYI, and auto-sent into ONE message.
        Each email gets a number for easy reference (more 1, spam 2, etc.)
        """
        total = len(newsletter_emails) + len(fyi_emails) + len(auto_sent_emails)
        if total == 0:
            return None

        # Clear previous mapping and build new one
        self.summary_email_mapping = {}
        email_num = 0

        content = f"""<h3>☀️ Morning Summary</h3><hr>"""

        # Newsletters section - with numbered references
        if newsletter_emails:
            content += f"""<h4>📰 Newsletters ({len(newsletter_emails)})</h4>"""
            for email in newsletter_emails[:8]:
                email_num += 1
                self.summary_email_mapping[email_num] = email.id
                sender_name = email.sender_name or email.sender_email.split('@')[0]
                preview = (email.summary or email.body_preview or "")[:100].replace('\n', ' ').strip()
                outlook_link = get_outlook_deep_link(email.message_id)

                content += f"""<p style="margin:4px 0;"><b>{email_num}. {sender_name}:</b> {email.subject[:55]}</p>"""
                if preview:
                    content += f"""<p style="margin:0 0 4px 18px;color:#666;font-size:0.9em;">{preview}...</p>"""
                content += f"""<p style="margin:0 0 10px 18px;font-size:0.85em;"><code>more {email_num}</code> • <a href="{outlook_link}">View in Outlook</a></p>"""

            if len(newsletter_emails) > 8:
                content += f"""<p><i>...+{len(newsletter_emails) - 8} more</i></p>"""

        # Auto-sent section - show what was replied (no numbers needed, just FYI)
        if auto_sent_emails:
            content += f"""<h4>✅ Auto-Replied ({len(auto_sent_emails)})</h4>"""
            for email in auto_sent_emails[:5]:
                draft_snippet = ""
                if email.current_draft:
                    lines = [l.strip() for l in email.current_draft.split('\n') if l.strip()]
                    for line in lines:
                        if not line.lower().startswith(('hi ', 'hello', 'dear ')):
                            draft_snippet = line[:80]
                            break
                content += f"""<p style="margin:4px 0;">• <b>{email.sender_email.split('@')[0]}:</b> {email.subject[:45]}</p>"""
                if draft_snippet:
                    content += f"""<p style="margin:0 0 8px 18px;color:#28a745;font-size:0.9em;">↳ "{draft_snippet}..."</p>"""
            if len(auto_sent_emails) > 5:
                content += f"""<p><i>...+{len(auto_sent_emails) - 5} more</i></p>"""

        # FYI section - with numbered references
        if fyi_emails:
            content += f"""<h4>📬 FYI - No Action Needed ({len(fyi_emails)})</h4>"""
            for email in fyi_emails[:8]:
                email_num += 1
                self.summary_email_mapping[email_num] = email.id
                sender_short = email.sender_name or email.sender_email.split('@')[0]
                preview = (email.summary or email.body_preview or "")[:100].replace('\n', ' ').strip()
                outlook_link = get_outlook_deep_link(email.message_id)

                content += f"""<p style="margin:4px 0;"><b>{email_num}. {sender_short}:</b> {email.subject[:50]}</p>"""
                if preview:
                    content += f"""<p style="margin:0 0 4px 18px;color:#666;font-size:0.9em;">{preview}...</p>"""
                content += f"""<p style="margin:0 0 10px 18px;font-size:0.85em;"><code>more {email_num}</code> • <a href="{outlook_link}">View in Outlook</a></p>"""

            if len(fyi_emails) > 8:
                content += f"""<p><i>...+{len(fyi_emails) - 8} more</i></p>"""

        content += """<hr>
<p><b>Commands:</b> <code>more [#]</code> details • <code>spam [#]</code> delete • <code>mute [#]</code> silence sender • <code>archive all</code> clear</p>"""

        # Persist the mapping to database for restart resilience
        self.db.save_summary_mapping(self.summary_email_mapping)
        logger.info(f"Saved summary mapping with {len(self.summary_email_mapping)} entries")

        message_id = self.teams_client.send_notification(content)
        if message_id:
            self.log_action(
                "morning_summary_sent",
                details={
                    "newsletters": len(newsletter_emails),
                    "fyi": len(fyi_emails),
                    "auto_sent": len(auto_sent_emails)
                }
            )

            # Mark these emails as notified (transition to ACKNOWLEDGED so they don't appear again)
            for email in newsletter_emails + fyi_emails:
                try:
                    email.transition_to(EmailState.ACKNOWLEDGED)
                    self.db.save_email(email)
                except Exception:
                    pass  # Already in a terminal state, that's fine

        return message_id

    async def _send_newsletter_summary(self, newsletter_emails: List[EmailRecord]) -> Optional[str]:
        """Send ONE grouped message for all newsletters with quick summaries. (Legacy - now using morning summary)"""
        content = f"""<h3>📰 Newsletters & Promotional ({len(newsletter_emails)})</h3><hr>"""

        for email in newsletter_emails[:15]:
            sender_name = email.sender_name or email.sender_email.split('@')[0]
            subject_short = email.subject[:60] + ('...' if len(email.subject) > 60 else '')
            # Quick one-line summary
            preview = (email.body_preview or "")[:80].replace('\n', ' ').strip()
            if preview:
                preview = f" - {preview}..."

            content += f"""<p style="margin:4px 0;">• <b>{sender_name}:</b> {subject_short}</p>"""

        if len(newsletter_emails) > 15:
            content += f"""<p><i>...+{len(newsletter_emails) - 15} more</i></p>"""

        content += """<hr>
<p><b>Commands:</b></p>
<p><code>ignore</code> - Keep all, don't notify me</p>
<p><code>mute [sender]</code> - Never show emails from this sender again</p>"""

        message_id = self.teams_client.send_notification(content)
        if message_id:
            self.log_action("newsletter_summary_sent", details={"count": len(newsletter_emails)})
        return message_id

    async def _send_action_email_notification(self, email: EmailRecord) -> Optional[str]:
        """Send ONE Teams message for an action-required email with detailed context."""
        token = email.approval_token or "N/A"
        vip_badge = "⭐ VIP " if email.is_vip else ""
        mailbox_tag = f" [{email.mailbox.split('@')[0]}]" if email.mailbox != settings.mailbox_email else ""
        priority_icon = "🔥 URGENT " if email.priority == 1 else ("⚡ " if email.priority == 2 else "")

        # Get summary - strip markdown and expand for better context
        summary_text = email.summary or email.body_preview[:500] or ""
        summary_text = re.sub(r'\*\*([^*]+)\*\*', r'\1', summary_text)
        summary_text = re.sub(r'\*([^*]+)\*', r'\1', summary_text)
        summary_text = re.sub(r'__([^_]+)__', r'\1', summary_text)
        summary_text = re.sub(r'_([^_]+)_', r'\1', summary_text)

        # Get more body preview for context (first 400 chars of actual content)
        body_context = ""
        if email.body_preview and len(email.body_preview) > 50:
            # Clean up the body preview
            body_text = email.body_preview[:400].replace('\n', ' ').replace('\r', ' ')
            body_text = re.sub(r'\s+', ' ', body_text).strip()
            if body_text and body_text != summary_text[:len(body_text)]:
                body_context = body_text

        # Get draft preview - show more content
        draft_preview = ""
        if email.current_draft:
            # Get substantive content from draft, skip greetings
            draft_lines = [l.strip() for l in email.current_draft.split('\n') if l.strip()]
            substantive_lines = []
            for line in draft_lines:
                if not line.lower().startswith(('hi ', 'hello', 'dear ', 'thanks', 'thank you', 'best', 'regards', 'cheers')):
                    substantive_lines.append(line)
            draft_preview = ' '.join(substantive_lines)[:300] if substantive_lines else draft_lines[0][:200] if draft_lines else ""

        # Format received time in local timezone
        received_time = format_local_time(email.received_at, '%I:%M %p')

        content = f"""<div style="border-left:4px solid {'#dc2626' if email.priority <= 2 else '#2563eb'}; padding-left:12px;">
<h3>{priority_icon}{vip_badge}{email.subject}{mailbox_tag}</h3>
<p><b>From:</b> {email.sender_name or 'Unknown'} &lt;{email.sender_email}&gt; • {received_time}</p>
<hr>
<p><b>What they need:</b> {summary_text[:400]}</p>
{f'<p style="background:#f8f8f8;padding:8px;border-radius:4px;font-size:0.9em;"><i>"{body_context[:300]}..."</i></p>' if body_context else ''}
{f'<hr><p><b>📝 Suggested Reply:</b></p><p style="background:#e8f5e9;padding:10px;border-radius:4px;">{draft_preview}</p>' if draft_preview else ''}
<hr>
<p><b>Quick Actions:</b></p>
<p>• <code>{token}</code> or <code>send</code> - Send reply</p>
<p>• <code>edit: [changes]</code> - Modify draft</p>
<p>• <code>ignore</code> - No reply needed</p>
<p>• <code>more</code> - Full email</p>
</div>"""

        message_id = self.teams_client.send_notification(content)
        if message_id:
            # Save the teams message ID for reply tracking
            email.teams_message_id = message_id
            self.db.save_email(email)
            self.log_action(
                "action_notification_sent",
                email_id=email.id,
                details={"subject": email.subject, "token": token}
            )
        return message_id

    async def _send_fyi_notification_deduped(self, email: EmailRecord) -> Optional[str]:
        """
        Send an FYI notification with deduplication.

        If a similar notification (same sender domain + normalized subject) already exists
        and the user hasn't responded, UPDATE that message instead of creating a new one.
        This prevents spam when the same alert fires repeatedly (e.g., VPN up/down).

        Args:
            email: The FYI email to notify about

        Returns:
            Teams message ID if sent/updated
        """
        outlook_link = get_outlook_deep_link(email.message_id)

        def content_builder(e: EmailRecord, count: int, email_ids: List[str], status_history: List[Dict[str, Any]] = None) -> str:
            """Build notification content with count and status history info."""
            sender = e.sender_name or e.sender_email
            preview = (e.body_preview or "")[:200].replace('\n', ' ').strip()
            status_history = status_history or []

            # Build status summary for alerts
            status_line = ""
            if status_history and count > 1:
                # Summarize status changes
                up_count = sum(1 for s in status_history if s.get("status") == "up")
                down_count = sum(1 for s in status_history if s.get("status") == "down")
                current = status_history[-1].get("status", "unknown") if status_history else "unknown"
                current_display = current.upper() if current in ["up", "down"] else current

                if up_count > 0 and down_count > 0:
                    status_line = f'<p style="background:#fff3cd;padding:6px;border-radius:4px;">⚡ <b>Flapped {count}x (↑{up_count} ↓{down_count}) → Currently {current_display}</b></p>'
                elif up_count > 0:
                    status_line = f'<p style="background:#d4edda;padding:6px;border-radius:4px;">✅ <b>{up_count} recoveries → Currently {current_display}</b></p>'
                elif down_count > 0:
                    status_line = f'<p style="background:#f8d7da;padding:6px;border-radius:4px;">⚠️ <b>{down_count} failures → Currently {current_display}</b></p>'
                else:
                    status_line = f'<p style="background:#fff3cd;padding:6px;border-radius:4px;">🔄 <b>Changed {count} times</b></p>'
            elif count > 1:
                status_line = f'<p style="background:#fff3cd;padding:6px;border-radius:4px;">🔄 <b>This alert has occurred {count} times</b></p>'

            if count > 1:
                # Multiple similar alerts - show count and status summary
                return f"""<div style="border-left: 3px solid #ffa500; padding-left: 10px;">
<p><b>ℹ️ FYI Alert ({count}x)</b> - Similar alerts grouped</p>
<p><b>From:</b> {sender}</p>
<p><b>Latest:</b> {e.subject}</p>
{status_line}
<hr>
<p style="color:#666;">{preview}...</p>
<hr>
<p><a href="{outlook_link}">📬 Open in Outlook</a></p>
<p><b>Commands:</b> <code>ignore</code> dismiss • <code>mute {e.sender_email}</code> stop these alerts</p>
</div>"""
            else:
                # First occurrence
                return f"""<div style="border-left: 3px solid #17a2b8; padding-left: 10px;">
<p><b>ℹ️ FYI</b></p>
<p><b>From:</b> {sender}</p>
<p><b>Subject:</b> {e.subject}</p>
<hr>
<p style="color:#666;">{preview}...</p>
<hr>
<p><a href="{outlook_link}">📬 Open in Outlook</a></p>
<p><b>Commands:</b> <code>ignore</code> dismiss • <code>mute {e.sender_email}</code> stop these</p>
</div>"""

        message_id, was_updated = self.teams_client.send_fyi_notification_deduped(
            email,
            content_builder=content_builder
        )

        if message_id:
            email.teams_message_id = message_id
            self.db.save_email(email)
            self.log_action(
                "fyi_notification_sent",
                email_id=email.id,
                details={
                    "subject": email.subject,
                    "was_update": was_updated
                }
            )

        return message_id

    async def _send_fyi_summary(
        self,
        fyi_emails: List[EmailRecord],
        auto_sent_emails: List[EmailRecord]
    ) -> Optional[str]:
        """Send grouped FYI and auto-sent summary."""
        if not fyi_emails and not auto_sent_emails:
            return None

        content = "<h3>📋 FYI Summary</h3><hr>"

        # Auto-sent section
        if auto_sent_emails:
            content += f"""<h4>✅ Auto-Sent ({len(auto_sent_emails)})</h4>"""
            for email in auto_sent_emails[:5]:
                mailbox_tag = f" [{email.mailbox.split('@')[0]}]" if email.mailbox != settings.mailbox_email else ""
                content += f"""<p>• {email.subject[:50]}{mailbox_tag} → {email.sender_email}</p>"""
            if len(auto_sent_emails) > 5:
                content += f"""<p><i>...and {len(auto_sent_emails) - 5} more</i></p>"""

        # FYI section
        if fyi_emails:
            content += f"""<h4>📬 No Action Needed ({len(fyi_emails)})</h4>"""
            for email in fyi_emails[:10]:
                sender_short = email.sender_name or email.sender_email.split('@')[0]
                subject_short = email.subject[:50] + ('...' if len(email.subject) > 50 else '')
                content += f"""<p style="margin:2px 0;">• <b>{sender_short}:</b> {subject_short}</p>"""
            if len(fyi_emails) > 10:
                content += f"""<p><i>...+{len(fyi_emails) - 10} more</i></p>"""

        message_id = self.teams_client.send_notification(content)
        if message_id:
            self.log_action(
                "fyi_summary_sent",
                details={"fyi": len(fyi_emails), "auto_sent": len(auto_sent_emails)}
            )
        return message_id

    def _group_by_sender(self, emails: List[EmailRecord]) -> Dict[str, List[EmailRecord]]:
        """Group emails by sender domain/email for cleaner display."""
        grouped = defaultdict(list)
        for email in emails:
            # Use domain as key for grouping
            sender = email.sender_email.lower()
            domain = sender.split('@')[-1] if '@' in sender else sender
            grouped[domain].append(email)
        # Sort by count (most emails first) then by priority
        return dict(sorted(grouped.items(), key=lambda x: (-len(x[1]), min(e.priority for e in x[1]))))

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

                # Debug logging to track what's being processed
                already_processed = 0
                new_count = 0

                for message in messages:
                    message_id = message.get("id", "")

                    # Skip if already processed
                    if self.db.is_message_processed(message_id, mailbox):
                        already_processed += 1
                        continue

                    new_count += 1
                    subject = message.get("subject", "No Subject")[:50]
                    logger.info(f"NEW email found: {subject} (id: {message_id[:30]}...)")

                    # Fetch full email details (list response has truncated body)
                    full_message = self.email_client.get_email_details(message_id, mailbox)
                    if full_message:
                        message = full_message

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

                # Log summary for this mailbox
                logger.info(f"Mailbox {mailbox}: {len(messages)} fetched, {already_processed} already processed, {new_count} new")

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

    async def check_followup_reminders(self) -> int:
        """
        Check for emails that need follow-up reminders and send notifications.

        Returns:
            Number of reminders sent
        """
        try:
            pending_followups = self.db.get_pending_followups()
            if not pending_followups:
                return 0

            reminders_sent = 0
            for email in pending_followups:
                # Build reminder notification
                outlook_link = get_outlook_deep_link(email.message_id)
                days_overdue = (datetime.utcnow() - email.follow_up_at).days if email.follow_up_at else 0

                urgency = "⏰"
                if days_overdue > 2:
                    urgency = "🚨"
                elif days_overdue > 0:
                    urgency = "⚠️"

                content = f"""<div style="border-left:4px solid #f59e0b; padding-left:12px;">
<h3>{urgency} Follow-up Reminder</h3>
<p><b>From:</b> {email.sender_name or email.sender_email}</p>
<p><b>Subject:</b> {email.subject}</p>
<p><b>Received:</b> {format_local_time(email.received_at, '%Y-%m-%d')}</p>
"""
                if email.follow_up_note:
                    content += f"<p><b>Note:</b> {email.follow_up_note}</p>"

                if days_overdue > 0:
                    content += f"<p style='color:#dc2626;'><b>Overdue by {days_overdue} day(s)</b></p>"

                content += f"""<hr>
<p><a href="{outlook_link}">📬 Open in Outlook</a></p>
<p><b>Commands:</b></p>
<p><code>done</code> complete • <code>followup 2d</code> snooze 2 days • <code>ignore</code> dismiss</p>
</div>"""

                self.teams_client.send_notification(content)

                # Increment reminder count
                email.follow_up_reminded_count += 1
                # Push reminder 1 day if overdue
                if days_overdue > 0:
                    from datetime import timedelta
                    email.follow_up_at = datetime.utcnow() + timedelta(days=1)
                self.db.save_email(email)

                reminders_sent += 1
                logger.info(f"Sent follow-up reminder for: {email.subject}")

            if reminders_sent > 0:
                self.log_action(
                    "followup_reminders_sent",
                    details={"count": reminders_sent}
                )

            return reminders_sent

        except Exception as e:
            logger.error(f"Error checking follow-up reminders: {e}")
            return 0

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

                # Handle batch commands (no email required)
                if command_type in ["dismiss_all", "review", "keep"]:
                    result = await self.handle_spam_batch_command(
                        command_type,
                        parameter
                    )
                    processed.append(result)
                    continue

                # Handle archive all - acknowledge all emails in morning summary
                if command_type == "archive_all":
                    result = await self.handle_archive_all_command()
                    processed.append(result)
                    continue

                # Handle numbered commands (more 3, spam 5, mute 2) from summary
                if command_type in ["more", "spam", "mute"] and parameter and parameter.isdigit():
                    num = int(parameter)
                    if num in self.summary_email_mapping:
                        email_id = self.summary_email_mapping[num]
                        email = self.db.get_email(email_id)
                        if email:
                            result = await self.handle_numbered_command(
                                email, command_type, num
                            )
                            processed.append(result)
                    else:
                        # Invalid number
                        self.teams_client.send_notification(
                            f"<p>❌ No email #{num} in the current summary. Valid numbers: {list(self.summary_email_mapping.keys())}</p>"
                        )
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
                # Teams approval ONLY works with explicit confirmation phrase
                # Regular "approve", "send", "yes" are NOT enough
                self.teams_client.send_notification(
                    f"<p>⚠️ To send emails via Teams, reply with: <b>CONFIRM SEND</b></p>"
                    f"<p>Or use the web dashboard: http://localhost:8080</p>"
                )
                result["success"] = False
                result["message"] = "Reply 'CONFIRM SEND' to approve"

            elif command_type == "confirm_send":
                # Explicit confirmation - actually send the email
                if email.current_draft:
                    sent = self.email_client.send_reply(email, email.current_draft)
                    if sent:
                        email.transition_to(EmailState.APPROVED)
                        email.transition_to(EmailState.SENT)
                        email.sent_at = datetime.utcnow()
                        email.handled_by = "user_teams"
                        result["success"] = True
                        self.teams_client.send_notification(
                            f"<p>✅ Email sent to {email.sender_email}</p>"
                        )

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
                # Clear any pending deduped notification for this pattern
                self.teams_client.clear_pending_for_email(email)
                result["success"] = True

            elif command_type == "spam":
                email.category = EmailCategory.SPAM_CANDIDATE
                email.transition_to(EmailState.SPAM_DETECTED)
                email.transition_to(EmailState.ARCHIVED)
                # Learn spam pattern from sender
                self._learn_spam_pattern(email)
                result["success"] = True

            elif command_type == "forward":
                if parameter:
                    forwarded = self.email_client.forward_email(email, parameter)
                    if forwarded:
                        email.transition_to(EmailState.FORWARDED)
                        result["success"] = True

            elif command_type == "mute":
                # Mute sender - never show emails from this sender again
                sender_to_mute = parameter or email.sender_email
                self.db.mute_sender(sender_to_mute, f"Muted via Teams command for email: {email.subject[:50]}")
                # Also archive the current email
                email.transition_to(EmailState.ARCHIVED)
                email.handled_by = "user"
                # Clear any pending deduped notification for this pattern
                self.teams_client.clear_pending_for_email(email)
                result["success"] = True
                # Send confirmation
                self.teams_client.send_notification(
                    f"<p>🔇 Muted <b>{sender_to_mute}</b> - You won't see emails from this sender again.</p>"
                )

            elif command_type == "followup":
                # Mark email for follow-up with optional time parameter
                follow_up_time = self._parse_followup_time(parameter)
                email.state = EmailState.FOLLOW_UP
                email.follow_up_at = follow_up_time
                email.follow_up_note = parameter
                email.handled_by = "user"
                result["success"] = True
                # Send confirmation with local time
                time_str = format_local_time(follow_up_time, '%Y-%m-%d %H:%M')
                self.teams_client.send_notification(
                    f"<p>⏰ Follow-up set for <b>{email.subject[:40]}</b> - Reminder: {time_str}</p>"
                )

            elif command_type == "more":
                # Fetch full email content from MS365 if not already populated
                if not email.body_full or len(email.body_full) < 100:
                    # The stored body might just be a preview - fetch full from MS365
                    try:
                        full_message = self.email_client.get_email_details(
                            email.message_id,
                            email.mailbox
                        )
                        if full_message:
                            body = full_message.get("body", {})
                            body_content = body.get("content", "")
                            if body_content:
                                email.body_full = body_content
                                self.db.save_email(email)
                                logger.info(f"Fetched full body for email {email.id}")
                    except Exception as e:
                        logger.warning(f"Could not fetch full email body: {e}")

                # Send full email content to Teams
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

    async def handle_numbered_command(
        self,
        email: EmailRecord,
        command_type: str,
        num: int
    ) -> Dict[str, Any]:
        """
        Handle commands for numbered emails from summary (more 3, spam 5).

        Args:
            email: The email record
            command_type: 'more' or 'spam'
            num: The number from the summary

        Returns:
            Result of command execution
        """
        result = {"success": False, "action": command_type, "email_num": num}

        try:
            if command_type == "more":
                # Send full email details
                outlook_link = get_outlook_deep_link(email.message_id)
                sender_name = email.sender_name or email.sender_email

                # Fetch full body if not available
                if not email.body_full:
                    try:
                        full_message = self.email_client.get_email_details(
                            email.message_id, email.mailbox
                        )
                        if full_message:
                            body = full_message.get("body", {})
                            email.body_full = body.get("content", "")
                            self.db.save_email(email)
                    except Exception as e:
                        logger.warning(f"Could not fetch full body: {e}")

                body_preview = (email.body_full or email.body_preview or "No content")[:800]
                # Clean up HTML tags for display
                import re as regex
                body_text = regex.sub(r'<[^>]+>', ' ', body_preview)
                body_text = regex.sub(r'\s+', ' ', body_text).strip()

                content = f"""<div style="border-left:4px solid #2563eb; padding-left:12px;">
<h3>📧 Email #{num} Details</h3>
<p><b>From:</b> {sender_name} &lt;{email.sender_email}&gt;</p>
<p><b>Subject:</b> {email.subject}</p>
<p><b>Received:</b> {format_local_time(email.received_at, '%Y-%m-%d %H:%M')}</p>
<hr>
<p><b>Content:</b></p>
<p style="background:#f5f5f5;padding:10px;border-radius:4px;">{body_text[:600]}...</p>
<hr>
<p><a href="{outlook_link}">📬 Open in Outlook</a></p>
<p><b>Commands:</b> <code>spam {num}</code> mark as spam • <code>mute {email.sender_email}</code> block sender</p>
</div>"""

                self.teams_client.send_notification(content)
                result["success"] = True

            elif command_type == "spam":
                # Mark as spam and delete
                self._learn_spam_pattern(email)
                try:
                    self.email_client.delete_email(email.message_id, email.mailbox)
                except Exception:
                    pass
                self.db.delete_email(email.id)

                # Remove from mapping and persist
                if num in self.summary_email_mapping:
                    del self.summary_email_mapping[num]
                    self.db.save_summary_mapping(self.summary_email_mapping)

                # Extract domain for the message
                sender_domain = email.sender_email.split('@')[1] if '@' in email.sender_email else email.sender_email
                self.teams_client.send_notification(
                    f"<p>🗑️ Email #{num} from <b>{email.sender_email}</b> marked as spam and deleted.</p>"
                    f"<p><i>📧 Promotional emails from {sender_domain} will be filtered. "
                    f"Important emails (password resets, orders) will still come through.</i></p>"
                )
                result["success"] = True

            elif command_type == "mute":
                # Mute sender - never notify about emails from this sender again
                sender_email = email.sender_email
                self.db.mute_sender(sender_email, f"Muted via 'mute {num}' command")

                # Mark email as acknowledged
                email.transition_to(EmailState.ACKNOWLEDGED)
                email.handled_by = "user"
                self.db.save_email(email)

                # Clear any pending deduped notification for this pattern
                self.teams_client.clear_pending_for_email(email)

                # Remove from mapping and persist
                if num in self.summary_email_mapping:
                    del self.summary_email_mapping[num]
                    self.db.save_summary_mapping(self.summary_email_mapping)

                self.teams_client.send_notification(
                    f"<p>🔇 Muted <b>{sender_email}</b> - You won't be notified about emails from this sender.</p>"
                )
                result["success"] = True

            self.log_action(
                f"numbered_{command_type}",
                email_id=email.id,
                details={"num": num, "sender": email.sender_email}
            )

        except Exception as e:
            logger.error(f"Error handling numbered command {command_type} {num}: {e}")
            result["error"] = str(e)

        return result

    async def handle_archive_all_command(self) -> Dict[str, Any]:
        """
        Handle 'archive all' command - acknowledge all emails in the morning summary.
        """
        result = {"success": False, "action": "archive_all"}

        try:
            if not self.summary_email_mapping:
                self.teams_client.send_notification(
                    "<p>❌ No emails in the current summary to archive.</p>"
                )
                return result

            archived_count = 0
            for num, email_id in list(self.summary_email_mapping.items()):
                email = self.db.get_email(email_id)
                if email:
                    try:
                        email.transition_to(EmailState.ACKNOWLEDGED)
                        email.handled_by = "user"
                        self.db.save_email(email)
                        archived_count += 1
                    except Exception as e:
                        logger.warning(f"Could not archive email {email_id}: {e}")

            # Clear the mapping since all are archived
            self.summary_email_mapping = {}
            self.db.save_summary_mapping(self.summary_email_mapping)

            self.teams_client.send_notification(
                f"<p>✅ Archived {archived_count} emails from the morning summary.</p>"
            )

            self.log_action(
                "archive_all",
                details={"count": archived_count}
            )

            result["success"] = True
            result["archived_count"] = archived_count

        except Exception as e:
            logger.error(f"Error archiving all: {e}")
            result["error"] = str(e)

        return result

    def _learn_spam_pattern(self, email: EmailRecord) -> None:
        """
        Learn a spam pattern from an email marked as spam by the user.

        Creates or updates spam rules based on sender domain.
        NOTE: Spam rules are advisory - transactional emails (password resets,
        order confirmations) from the same domain will NOT be blocked.
        """
        try:
            # Extract sender domain
            sender_lower = email.sender_email.lower()
            if '@' in sender_lower:
                sender_domain = sender_lower.split('@')[1]

                # Check if we already have a rule for this domain
                existing_rules = self.db.get_active_spam_rules()
                domain_rule = next(
                    (r for r in existing_rules if r.rule_type == 'domain' and r.pattern == sender_domain),
                    None
                )

                if domain_rule:
                    # Increase confidence of existing rule
                    self.db.increment_spam_rule_hit(domain_rule.id)
                    logger.info(f"Increased confidence for spam domain rule: {sender_domain} (transactional emails still allowed)")
                else:
                    # Create new rule for this domain
                    # Note: This only affects promotional/marketing emails
                    # Transactional emails (password resets, orders) bypass spam rules
                    new_rule = SpamRule(
                        rule_type='domain',
                        pattern=sender_domain,
                        action='archive',
                        confidence=60,  # Start with moderate confidence
                    )
                    self.db.save_spam_rule(new_rule)
                    logger.info(f"Created spam domain rule: {sender_domain} (transactional emails excluded)")

                self.log_action(
                    "spam_pattern_learned",
                    email_id=email.id,
                    details={
                        "domain": sender_domain,
                        "note": "Transactional emails from this domain will still be delivered"
                    }
                )

        except Exception as e:
            logger.warning(f"Failed to learn spam pattern: {e}")

    def _parse_followup_time(self, param: Optional[str]) -> datetime:
        """
        Parse follow-up time from user parameter.

        Supports:
        - "tomorrow" / "tmr"
        - "1d", "2d", "3d" (days)
        - "1h", "2h" (hours)
        - "monday", "tuesday", etc.
        - Number only (assumes email number, default to 1 day)
        - None/empty (default to 1 day)
        """
        from datetime import timedelta

        now = datetime.utcnow()
        default = now + timedelta(days=1)

        if not param:
            return default

        param_lower = param.strip().lower()

        # Check for number only (email reference from summary)
        if param_lower.isdigit():
            return default

        # Tomorrow
        if param_lower in ["tomorrow", "tmr", "tom"]:
            return now + timedelta(days=1)

        # Days: 1d, 2d, etc.
        if re.match(r"^\d+d$", param_lower):
            days = int(param_lower[:-1])
            return now + timedelta(days=days)

        # Hours: 1h, 2h, etc.
        if re.match(r"^\d+h$", param_lower):
            hours = int(param_lower[:-1])
            return now + timedelta(hours=hours)

        # Weekday names
        weekdays = {
            "monday": 0, "mon": 0,
            "tuesday": 1, "tue": 1,
            "wednesday": 2, "wed": 2,
            "thursday": 3, "thu": 3,
            "friday": 4, "fri": 4,
            "saturday": 5, "sat": 5,
            "sunday": 6, "sun": 6,
        }
        if param_lower in weekdays:
            target_day = weekdays[param_lower]
            current_day = now.weekday()
            days_ahead = target_day - current_day
            if days_ahead <= 0:  # Target day already happened this week
                days_ahead += 7
            return now + timedelta(days=days_ahead)

        # Default to 1 day if can't parse
        return default

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
