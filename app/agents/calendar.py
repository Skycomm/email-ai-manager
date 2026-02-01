"""
Calendar Agent - Handles meeting-related emails.

This agent:
- Detects meeting invitations
- Checks for calendar conflicts
- Suggests accept/decline/tentative responses
- Can auto-accept internal meetings (configurable)
"""

import logging
from datetime import datetime, timedelta
from typing import Dict, Any, Optional, List

from .base import BaseAgent
from ..db import Database
from ..models import EmailRecord, EmailCategory
from ..integrations import MCPClient
from ..config import settings

logger = logging.getLogger(__name__)


# Meeting-related keywords
MEETING_KEYWORDS = [
    "meeting", "invite", "invitation", "calendar", "schedule",
    "appointment", "call", "conference", "sync", "standup",
    "review", "discussion", "catch up", "1:1", "one-on-one"
]

MEETING_SUBJECT_PATTERNS = [
    "meeting request", "invitation:", "invite:", "calendar:",
    "accepted:", "declined:", "tentative:", "canceled:",
    "updated:", "rescheduled:"
]


def is_meeting_email(email: EmailRecord) -> bool:
    """Check if an email is meeting-related."""
    subject_lower = email.subject.lower()
    body_lower = (email.body_preview or "").lower()

    # Check subject patterns
    for pattern in MEETING_SUBJECT_PATTERNS:
        if pattern in subject_lower:
            return True

    # Check keywords in subject
    for keyword in MEETING_KEYWORDS:
        if keyword in subject_lower:
            return True

    # Check for .ics attachment (calendar invite)
    if email.has_attachments and ".ics" in body_lower:
        return True

    return False


class CalendarAgent(BaseAgent):
    """
    Agent for handling meeting-related emails.
    """

    def __init__(self, db: Database, mcp_client: Optional[MCPClient] = None):
        super().__init__(
            db=db,
            name="calendar",
            system_prompt="You help analyze meeting invitations and suggest appropriate responses."
        )
        self.mcp = mcp_client or MCPClient()

    async def process(self, email: EmailRecord) -> Dict[str, Any]:
        """
        Process a meeting-related email.

        Args:
            email: The meeting email to process

        Returns:
            Dict with meeting details and suggested action
        """
        result = {
            "is_meeting": False,
            "meeting_type": None,  # invite, response, update, cancellation
            "suggested_action": None,
            "has_conflict": False,
            "conflict_details": None,
            "auto_responded": False
        }

        if not settings.calendar_integration_enabled:
            return result

        # Check if this is a meeting email
        if not is_meeting_email(email):
            return result

        result["is_meeting"] = True
        email.category = EmailCategory.MEETING

        # Determine meeting type
        result["meeting_type"] = self._detect_meeting_type(email)

        # For invitations, check for conflicts
        if result["meeting_type"] == "invite":
            # Try to extract meeting time from email
            meeting_time = await self._extract_meeting_time(email)

            if meeting_time and settings.calendar_check_conflicts:
                conflict = await self._check_calendar_conflict(
                    email.mailbox,
                    meeting_time["start"],
                    meeting_time["end"]
                )
                result["has_conflict"] = conflict.get("has_conflict", False)
                result["conflict_details"] = conflict.get("conflicting_events", [])

            # Suggest action
            result["suggested_action"] = self._suggest_action(
                email,
                result["has_conflict"],
                result["meeting_type"]
            )

            # Auto-accept internal meetings if configured
            if (settings.calendar_auto_accept_internal and
                not result["has_conflict"] and
                self._is_internal_sender(email)):

                # Try to auto-accept
                accepted = await self._auto_accept_meeting(email)
                result["auto_responded"] = accepted
                if accepted:
                    logger.info(f"Auto-accepted meeting from {email.sender_email}: {email.subject}")

        self.log_action(
            "meeting_processed",
            email_id=email.id,
            details=result
        )

        return result

    def _detect_meeting_type(self, email: EmailRecord) -> str:
        """Detect the type of meeting email."""
        subject_lower = email.subject.lower()

        if "canceled" in subject_lower or "cancelled" in subject_lower:
            return "cancellation"
        elif "accepted:" in subject_lower:
            return "response_accepted"
        elif "declined:" in subject_lower:
            return "response_declined"
        elif "tentative:" in subject_lower:
            return "response_tentative"
        elif "updated:" in subject_lower or "rescheduled:" in subject_lower:
            return "update"
        elif any(p in subject_lower for p in ["invite", "invitation", "meeting request"]):
            return "invite"
        else:
            return "meeting_related"

    async def _extract_meeting_time(self, email: EmailRecord) -> Optional[Dict[str, str]]:
        """
        Try to extract meeting start/end time from email.

        In a real implementation, this would parse .ics attachments
        or use Claude to extract time from email body.
        """
        # For now, use Claude to try to extract meeting time
        messages = [{
            "role": "user",
            "content": f"""Extract the meeting time from this email if present:

Subject: {email.subject}
Body: {email.body_preview[:500]}

If you can identify a meeting time, respond with JSON:
{{"start": "YYYY-MM-DDTHH:MM:SS", "end": "YYYY-MM-DDTHH:MM:SS"}}

If no meeting time found, respond with: {{"start": null, "end": null}}

Only output JSON."""
        }]

        try:
            result = self.call_claude_structured(
                messages,
                schema={
                    "type": "object",
                    "properties": {
                        "start": {"type": ["string", "null"]},
                        "end": {"type": ["string", "null"]}
                    }
                }
            )

            if result.get("start"):
                return result
        except Exception as e:
            logger.debug(f"Could not extract meeting time: {e}")

        return None

    async def _check_calendar_conflict(
        self,
        mailbox: str,
        start_time: str,
        end_time: str
    ) -> Dict[str, Any]:
        """Check for calendar conflicts at the given time."""
        result = {"has_conflict": False, "conflicting_events": []}

        try:
            # Get events in the time window
            events = self.mcp.get_calendar_view(
                start_datetime=start_time,
                end_datetime=end_time,
                user_email=mailbox
            )

            if events:
                result["has_conflict"] = True
                result["conflicting_events"] = [
                    {
                        "subject": e.get("subject", "Untitled"),
                        "start": e.get("start", {}).get("dateTime"),
                        "end": e.get("end", {}).get("dateTime")
                    }
                    for e in events[:3]  # Limit to 3
                ]

        except Exception as e:
            logger.warning(f"Could not check calendar conflicts: {e}")

        return result

    def _suggest_action(
        self,
        email: EmailRecord,
        has_conflict: bool,
        meeting_type: str
    ) -> str:
        """Suggest appropriate action for the meeting email."""
        if meeting_type != "invite":
            return "acknowledge"  # Just FYI

        if has_conflict:
            return "decline_or_reschedule"

        if self._is_internal_sender(email):
            return "accept"
        else:
            return "review"  # External meeting needs review

    def _is_internal_sender(self, email: EmailRecord) -> bool:
        """Check if sender is from internal domain."""
        sender_domain = email.sender_email.split('@')[-1].lower()
        for domain in settings.internal_domains:
            if domain.lower() == sender_domain:
                return True
        return False

    async def _auto_accept_meeting(self, email: EmailRecord) -> bool:
        """
        Try to auto-accept a meeting invitation.

        Note: This requires the email to have an associated calendar event ID,
        which we may not always have from the email alone.
        """
        # In a real implementation, we'd need to:
        # 1. Find the calendar event ID from the meeting request
        # 2. Call accept-event-invite with that ID

        # For now, just log that we would auto-accept
        logger.info(f"Would auto-accept meeting: {email.subject}")
        return False  # Return False since we can't actually auto-accept without event ID

    def suggest_meeting_response(self, email: EmailRecord, conflict_info: Optional[Dict] = None) -> str:
        """Generate a suggested response for a meeting invitation."""
        if conflict_info and conflict_info.get("has_conflict"):
            conflicts = conflict_info.get("conflicting_events", [])
            conflict_names = [c.get("subject", "another meeting") for c in conflicts[:2]]

            return f"""Hi {email.sender_name or email.sender_email.split('@')[0]},

Thanks for the meeting invitation. Unfortunately, I have a conflict at that time ({', '.join(conflict_names)}).

Could we look at an alternative time? I'm happy to find a slot that works for both of us.

Best regards,
David"""

        return f"""Hi {email.sender_name or email.sender_email.split('@')[0]},

Thanks for the meeting invitation. I'll review my calendar and get back to you shortly.

Best regards,
David"""
