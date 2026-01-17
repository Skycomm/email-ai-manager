"""
Spam Filter Agent - Identifies and batches spam/unwanted emails.

This agent:
- Evaluates emails against spam patterns
- Learns from user corrections
- Batches spam notifications for bulk dismissal
"""

import logging
from datetime import datetime, timedelta
from typing import Dict, Any, List, Optional, Tuple

from .base import BaseAgent
from ..db import Database
from ..models import EmailRecord, EmailState, EmailCategory

logger = logging.getLogger(__name__)


# Common spam indicators
SPAM_KEYWORDS = [
    "unsubscribe", "newsletter", "promotional", "limited time",
    "act now", "click here", "free", "winner", "congratulations",
    "urgent action required", "verify your account", "confirm your",
    "marketing", "sale ends", "discount code", "special offer"
]

SPAM_SENDER_PATTERNS = [
    "noreply", "no-reply", "newsletter", "marketing",
    "promo", "info@", "notifications@", "alerts@"
]

# Sender domains that are always spam/FYI (automated system emails)
SPAM_SENDER_DOMAINS = [
    "semsportalmail1.com",  # Solar monitoring daily generation reports
]

# Subject patterns that indicate spam (exact substring matches)
SPAM_SUBJECT_PATTERNS = [
    "[openai/whisper]",  # GitHub discussions for openai/whisper repo
    "daily generation info",  # Solar monitoring automated reports
]


SPAM_FILTER_SYSTEM_PROMPT = """You are a Spam Filter Agent responsible for identifying unwanted emails.

Analyze emails to determine spam probability (0-100) based on:
1. Sender patterns (noreply, marketing addresses)
2. Subject line keywords (unsubscribe, promotional, etc.)
3. Content patterns (urgent requests, verification links)
4. Whether it requires genuine user action vs automated notification

Categories:
- SPAM: Marketing, newsletters, promotional (score >= 70)
- LIKELY_SPAM: Automated notifications with no action needed (score 40-69)
- NOT_SPAM: Legitimate emails needing attention (score < 40)

Be conservative - when in doubt, mark as NOT_SPAM to avoid missing important emails.
"""


class SpamFilterAgent(BaseAgent):
    """
    Agent for detecting and batching spam emails.
    """

    def __init__(self, db: Database):
        super().__init__(
            db=db,
            name="spam_filter",
            system_prompt=SPAM_FILTER_SYSTEM_PROMPT
        )
        self._spam_batch: List[EmailRecord] = []
        self._last_notification_time: Optional[datetime] = None
        self._notification_interval = timedelta(minutes=5)

    async def process(self, email: EmailRecord) -> Dict[str, Any]:
        """
        Analyze an email for spam characteristics.

        Args:
            email: The email to analyze

        Returns:
            Dict with spam_score, is_spam, and reasoning
        """
        # Quick heuristic check first
        heuristic_score = self._heuristic_spam_score(email)

        # If clearly spam by heuristics, skip AI call
        if heuristic_score >= 80:
            return {
                "spam_score": heuristic_score,
                "is_spam": True,
                "likely_spam": True,
                "reasoning": "High-confidence spam based on patterns"
            }

        # Use AI for borderline cases
        if heuristic_score >= 30:
            ai_result = await self._ai_spam_analysis(email)
            # Combine scores (weighted average)
            combined_score = int(heuristic_score * 0.4 + ai_result["spam_score"] * 0.6)
            return {
                "spam_score": combined_score,
                "is_spam": combined_score >= 70,
                "likely_spam": combined_score >= 40,
                "reasoning": ai_result.get("reasoning", "AI analysis")
            }

        return {
            "spam_score": heuristic_score,
            "is_spam": False,
            "likely_spam": False,
            "reasoning": "Low spam probability"
        }

    def _heuristic_spam_score(self, email: EmailRecord) -> int:
        """
        Calculate spam score using simple heuristics.

        Args:
            email: The email to analyze

        Returns:
            Spam score 0-100
        """
        score = 0
        reasons = []

        # Check sender patterns
        sender_lower = email.sender_email.lower()
        for pattern in SPAM_SENDER_PATTERNS:
            if pattern in sender_lower:
                score += 15
                reasons.append(f"Sender matches '{pattern}'")
                break

        # Check sender domains (high confidence spam)
        for domain in SPAM_SENDER_DOMAINS:
            if domain.lower() in sender_lower:
                score += 80  # High score - these are definite spam
                reasons.append(f"Sender domain matches spam domain '{domain}'")
                break

        # Check subject patterns (high confidence spam)
        subject_lower = email.subject.lower()
        for pattern in SPAM_SUBJECT_PATTERNS:
            if pattern.lower() in subject_lower:
                score += 80  # High score - these are definite spam
                reasons.append(f"Subject matches spam pattern '{pattern}'")
                break

        # Check subject keywords
        keyword_hits = 0
        for keyword in SPAM_KEYWORDS:
            if keyword in subject_lower:
                keyword_hits += 1
        if keyword_hits > 0:
            score += min(30, keyword_hits * 10)
            reasons.append(f"Subject contains {keyword_hits} spam keywords")

        # Check body preview
        body_lower = (email.body_preview or "").lower()
        body_hits = 0
        for keyword in SPAM_KEYWORDS:
            if keyword in body_lower:
                body_hits += 1
        if body_hits > 0:
            score += min(20, body_hits * 5)
            reasons.append(f"Body contains {body_hits} spam keywords")

        # Check for unsubscribe indicator
        if "unsubscribe" in body_lower or "unsubscribe" in subject_lower:
            score += 25
            reasons.append("Contains unsubscribe link")

        # Check for promotional domains
        promotional_domains = [
            "mailchimp.com", "sendgrid.net", "constantcontact.com",
            "campaign-archive.com", "e.email", "mail.beehiiv.com"
        ]
        for domain in promotional_domains:
            if domain in sender_lower:
                score += 20
                reasons.append(f"Sent via promotional service: {domain}")
                break

        # Log the analysis
        if reasons:
            logger.debug(f"Spam heuristics for '{email.subject}': score={score}, reasons={reasons}")

        return min(100, score)

    async def _ai_spam_analysis(self, email: EmailRecord) -> Dict[str, Any]:
        """
        Use Claude to analyze spam probability for borderline cases.

        Args:
            email: The email to analyze

        Returns:
            Dict with spam_score and reasoning
        """
        messages = [{
            "role": "user",
            "content": f"""Analyze this email for spam probability:

From: {email.sender_name or email.sender_email} <{email.sender_email}>
Subject: {email.subject}
Body Preview: {email.body_preview[:500] if email.body_preview else 'N/A'}

Rate the spam probability from 0-100 where:
- 0-39: Legitimate email requiring attention
- 40-69: Automated notification, likely doesn't need action
- 70-100: Spam/marketing/newsletter

Consider:
1. Is this a marketing email or newsletter?
2. Does this require any action from the recipient?
3. Is this a transactional notification (receipts, confirmations)?
4. Could missing this email cause problems?

Respond in JSON with: spam_score (number), reasoning (string)"""
        }]

        schema = {
            "type": "object",
            "properties": {
                "spam_score": {"type": "integer"},
                "reasoning": {"type": "string"}
            }
        }

        try:
            result = self.call_claude_structured(messages, schema)
            return {
                "spam_score": min(100, max(0, result.get("spam_score", 50))),
                "reasoning": result.get("reasoning", "AI analysis")
            }
        except Exception as e:
            logger.warning(f"AI spam analysis failed: {e}")
            return {
                "spam_score": 50,
                "reasoning": "Fallback score due to AI error"
            }

    def add_to_batch(self, email: EmailRecord) -> int:
        """
        Add a spam email to the batch for bulk notification.

        Args:
            email: The spam email to batch

        Returns:
            Current batch size
        """
        self._spam_batch.append(email)
        self.log_action(
            "spam_batched",
            email_id=email.id,
            details={
                "subject": email.subject,
                "batch_size": len(self._spam_batch)
            }
        )
        return len(self._spam_batch)

    def get_batch(self) -> List[EmailRecord]:
        """Get the current spam batch."""
        return self._spam_batch.copy()

    def clear_batch(self) -> List[EmailRecord]:
        """Clear and return the spam batch."""
        batch = self._spam_batch.copy()
        self._spam_batch = []
        return batch

    def should_send_notification(self) -> bool:
        """
        Check if it's time to send a batched spam notification.

        Returns:
            True if batch should be sent
        """
        if not self._spam_batch:
            return False

        # Send if batch is large enough
        if len(self._spam_batch) >= 5:
            return True

        # Send if interval has passed since last notification
        if self._last_notification_time:
            if datetime.utcnow() - self._last_notification_time >= self._notification_interval:
                return True
        else:
            # First notification - send after interval from first email
            return len(self._spam_batch) >= 1

        return False

    def mark_notification_sent(self):
        """Mark that a spam notification was sent."""
        self._last_notification_time = datetime.utcnow()

    async def archive_batch(self) -> int:
        """
        Archive all emails in the current spam batch.

        Returns:
            Number of emails archived
        """
        count = 0
        for email in self._spam_batch:
            try:
                email.transition_to(EmailState.ARCHIVED)
                self.db.save_email(email)
                count += 1
            except Exception as e:
                logger.error(f"Failed to archive spam email {email.id}: {e}")

        self.log_action(
            "spam_batch_archived",
            details={"count": count}
        )

        self._spam_batch = []
        return count
