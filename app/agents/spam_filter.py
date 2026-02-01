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
from ..config import settings
from ..db import Database
from ..models import EmailRecord, EmailState, EmailCategory

logger = logging.getLogger(__name__)


# Common spam indicators (built-in defaults)
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

# Transactional email indicators - these override spam rules for the same domain
# e.g., Netflix promotional emails might be spam, but password resets are not
TRANSACTIONAL_SUBJECT_PATTERNS = [
    # Account security
    "password reset", "reset your password", "password changed",
    "reset password", "forgot password", "new password",
    "verify your email", "confirm your email", "email verification",
    "two-factor", "2fa", "authentication code", "security code",
    "login attempt", "sign-in", "signin", "sign in",
    "suspicious activity", "security alert", "account locked",
    # Orders & purchases
    "order confirmation", "order confirmed", "your order",
    "receipt for", "payment received", "payment confirmation",
    "invoice", "your purchase", "shipping confirmation",
    "delivery", "tracking number", "shipped", "dispatched",
    "refund", "return confirmation",
    # Subscriptions & billing
    "subscription", "renewal", "billing", "payment due",
    "card expiring", "payment failed", "payment method",
    # Account changes
    "account created", "welcome to", "registration",
    "profile updated", "settings changed", "email changed",
]

# Body patterns that indicate transactional (supplement subject patterns)
TRANSACTIONAL_BODY_PATTERNS = [
    "click the link below to reset",
    "didn't request this", "ignore this email",
    "order total", "subtotal", "grand total",
    "tracking number", "track your order",
    "your verification code is",
    "expires in", "valid for",
]


SPAM_FILTER_SYSTEM_PROMPT = """You are a Spam Filter Agent responsible for identifying unwanted emails.

Classify emails into:
1. HARD_SPAM: Scams, phishing, unsolicited junk - should be deleted
2. NEWSLETTER: Legitimate newsletters/promotional user subscribed to - group and summarize
3. NOT_SPAM: Legitimate emails needing attention

Return spam_score (0-100), is_newsletter (bool), and reasoning.
Be conservative - when in doubt, mark as NOT_SPAM.
"""

# Known newsletter senders (legitimate promotional)
NEWSLETTER_DOMAINS = [
    "substack.com", "beehiiv.com", "mailchimp.com", "sendgrid.net",
    "constantcontact.com", "buttondown.email", "convertkit.com",
    "skool.com", "circle.so"
]


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
        self._notification_interval = timedelta(minutes=settings.spam_notification_interval_minutes)
        self._batch_size = settings.spam_batch_size

    def _is_transactional_email(self, email: EmailRecord) -> Tuple[bool, str]:
        """
        Check if an email is transactional (password reset, order confirmation, etc.)
        Transactional emails should NOT be marked as spam even if from a "spam" domain.

        Args:
            email: The email to check

        Returns:
            Tuple of (is_transactional, reason)
        """
        subject_lower = email.subject.lower()
        body_lower = (email.body_preview or "").lower()

        # Check subject patterns
        for pattern in TRANSACTIONAL_SUBJECT_PATTERNS:
            if pattern in subject_lower:
                return True, f"Subject matches transactional pattern: '{pattern}'"

        # Check body patterns
        for pattern in TRANSACTIONAL_BODY_PATTERNS:
            if pattern in body_lower:
                return True, f"Body matches transactional pattern: '{pattern}'"

        return False, ""

    async def process(self, email: EmailRecord) -> Dict[str, Any]:
        """
        Analyze an email for spam/newsletter characteristics.

        IMPORTANT: Transactional emails (password resets, order confirmations, etc.)
        are NEVER marked as spam, even if from a domain with a spam rule.
        This prevents blocking legitimate emails like Netflix password resets
        just because Netflix promotional emails were marked as spam.

        Args:
            email: The email to analyze

        Returns:
            Dict with spam_score, is_spam, is_newsletter, is_transactional, and reasoning
        """
        # FIRST: Check if this is a transactional email - these are NEVER spam
        is_transactional, transactional_reason = self._is_transactional_email(email)
        if is_transactional:
            logger.info(f"Transactional email detected: {email.subject} - {transactional_reason}")
            return {
                "spam_score": 0,
                "is_spam": False,
                "is_newsletter": False,
                "is_transactional": True,
                "likely_spam": False,
                "reasoning": transactional_reason
            }

        # Check for known newsletter domains
        sender_lower = email.sender_email.lower()
        is_newsletter = any(domain in sender_lower for domain in NEWSLETTER_DOMAINS)

        # Quick heuristic check
        heuristic_score, heuristic_is_newsletter = self._heuristic_spam_score(email)

        # Combine newsletter detection
        is_newsletter = is_newsletter or heuristic_is_newsletter

        # If clearly spam by heuristics (and not a newsletter), skip AI call
        if heuristic_score >= 80 and not is_newsletter:
            return {
                "spam_score": heuristic_score,
                "is_spam": True,
                "is_newsletter": False,
                "is_transactional": False,
                "likely_spam": True,
                "reasoning": "High-confidence spam based on patterns"
            }

        # If it's a newsletter, mark it as such
        if is_newsletter or (heuristic_score >= 50 and "unsubscribe" in (email.body_preview or "").lower()):
            return {
                "spam_score": heuristic_score,
                "is_spam": False,
                "is_newsletter": True,
                "is_transactional": False,
                "likely_spam": False,
                "reasoning": "Newsletter/promotional content"
            }

        # Use AI for borderline cases
        if heuristic_score >= 30:
            ai_result = await self._ai_spam_analysis(email)
            combined_score = int(heuristic_score * 0.4 + ai_result["spam_score"] * 0.6)
            return {
                "spam_score": combined_score,
                "is_spam": combined_score >= 70 and not ai_result.get("is_newsletter"),
                "is_newsletter": ai_result.get("is_newsletter", False),
                "is_transactional": False,
                "likely_spam": combined_score >= 40,
                "reasoning": ai_result.get("reasoning", "AI analysis")
            }

        return {
            "spam_score": heuristic_score,
            "is_spam": False,
            "is_newsletter": False,
            "is_transactional": False,
            "likely_spam": False,
            "reasoning": "Low spam probability"
        }

    def _heuristic_spam_score(self, email: EmailRecord) -> Tuple[int, bool]:
        """
        Calculate spam score using simple heuristics.

        Args:
            email: The email to analyze

        Returns:
            Tuple of (spam_score 0-100, is_newsletter bool)
        """
        score = 0
        reasons = []
        is_newsletter = False

        # Check sender patterns
        sender_lower = email.sender_email.lower()
        for pattern in SPAM_SENDER_PATTERNS:
            if pattern in sender_lower:
                score += 15
                reasons.append(f"Sender matches '{pattern}'")
                break

        # Check sender domains from config (high confidence spam)
        for domain in settings.spam_sender_domains:
            if domain.lower() in sender_lower:
                score += 80  # High score - these are definite spam
                reasons.append(f"Sender domain matches spam domain '{domain}'")
                break

        # Check subject patterns from config (high confidence spam)
        subject_lower = email.subject.lower()
        for pattern in settings.spam_subject_patterns:
            if pattern.lower() in subject_lower:
                score += 80  # High score - these are definite spam
                reasons.append(f"Subject matches spam pattern '{pattern}'")
                break

        # Check for newsletter indicators
        newsletter_indicators = ["digest", "weekly", "newsletter", "update from", "your daily"]
        for indicator in newsletter_indicators:
            if indicator in subject_lower:
                is_newsletter = True
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

        # Check for unsubscribe indicator - this suggests newsletter not hard spam
        if "unsubscribe" in body_lower or "unsubscribe" in subject_lower:
            score += 25
            is_newsletter = True  # Legit newsletters have unsubscribe
            reasons.append("Contains unsubscribe link (likely newsletter)")

        # Check for known newsletter domains
        for domain in NEWSLETTER_DOMAINS:
            if domain in sender_lower:
                is_newsletter = True
                score += 30  # Still spammy but it's a newsletter
                reasons.append(f"Sent via newsletter service: {domain}")
                break

        # Log the analysis
        if reasons:
            logger.debug(f"Spam heuristics for '{email.subject}': score={score}, newsletter={is_newsletter}, reasons={reasons}")

        return min(100, score), is_newsletter

    async def _ai_spam_analysis(self, email: EmailRecord) -> Dict[str, Any]:
        """
        Use Claude to analyze spam/newsletter probability.

        Args:
            email: The email to analyze

        Returns:
            Dict with spam_score, is_newsletter, and reasoning
        """
        messages = [{
            "role": "user",
            "content": f"""Analyze this email:

From: {email.sender_name or email.sender_email} <{email.sender_email}>
Subject: {email.subject}
Body Preview: {email.body_preview[:500] if email.body_preview else 'N/A'}

Classify:
1. spam_score (0-100): How spammy/unwanted is this?
2. is_newsletter (bool): Is this a legitimate newsletter/promotional the user subscribed to?
3. reasoning: Brief explanation

Guidelines:
- Hard spam (scams, phishing, unsolicited) = high score, is_newsletter=false
- Newsletters with unsubscribe = moderate score, is_newsletter=true
- Legitimate emails = low score, is_newsletter=false

Respond in JSON: spam_score, is_newsletter, reasoning"""
        }]

        schema = {
            "type": "object",
            "properties": {
                "spam_score": {"type": "integer"},
                "is_newsletter": {"type": "boolean"},
                "reasoning": {"type": "string"}
            }
        }

        try:
            result = self.call_claude_structured(messages, schema)
            return {
                "spam_score": min(100, max(0, result.get("spam_score", 50))),
                "is_newsletter": result.get("is_newsletter", False),
                "reasoning": result.get("reasoning", "AI analysis")
            }
        except Exception as e:
            logger.warning(f"AI spam analysis failed: {e}")
            return {
                "spam_score": 50,
                "is_newsletter": False,
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
        if len(self._spam_batch) >= self._batch_size:
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
