"""
Drafting Agent - Generates email summaries and reply drafts.

This agent:
- Summarizes incoming emails
- Generates appropriate reply drafts
- Adapts tone based on context
- Never makes commitments without approval
"""

import logging
from typing import Dict, Any, Optional

from .base import BaseAgent
from ..db import Database
from ..models import EmailRecord, DraftMode

logger = logging.getLogger(__name__)


DRAFTING_SYSTEM_PROMPT = """You are the Email Drafting Agent, responsible for creating professional email replies.

Your job is to:
1. Summarize incoming emails clearly and concisely (2-3 sentences)
2. Generate appropriate reply drafts that match the context

Guidelines for drafts:
- Be professional but warm
- Be concise - get to the point quickly
- Never make specific commitments (dates, amounts, promises) without explicit approval
- Never share confidential information
- Match the sender's tone (formal vs casual)
- Include appropriate greetings and sign-offs
- If the email requires information you don't have, acknowledge and say you'll follow up

Draft modes:
- PROFESSIONAL: Formal business tone, structured response
- FRIENDLY: Warm but professional, more conversational
- BRIEF: Minimal acknowledgment, short and direct
- DETAILED: Comprehensive response with full context

Always err on the side of being helpful while maintaining professionalism.
"""


class DraftingAgent(BaseAgent):
    """
    Agent for summarizing emails and generating reply drafts.
    """

    def __init__(self, db: Database):
        super().__init__(
            db=db,
            name="drafting",
            system_prompt=DRAFTING_SYSTEM_PROMPT
        )

    async def process(self, email: EmailRecord) -> Dict[str, Any]:
        """
        Generate summary and draft reply for an email.

        Args:
            email: The email to process

        Returns:
            Dict with summary and draft
        """
        # Generate summary
        summary = await self.summarize(email)

        # Generate draft based on mode
        draft = await self.generate_draft(email, summary)

        self.log_action(
            "draft_generated",
            email_id=email.id,
            details={
                "summary_length": len(summary),
                "draft_length": len(draft),
                "mode": email.draft_mode.value
            }
        )

        return {
            "summary": summary,
            "draft": draft
        }

    async def summarize(self, email: EmailRecord) -> str:
        """
        Generate a concise summary of an email.

        Args:
            email: The email to summarize

        Returns:
            Summary string (2-3 sentences)
        """
        messages = [{
            "role": "user",
            "content": f"""Summarize this email in 2-3 concise sentences:

From: {email.sender_name or email.sender_email} <{email.sender_email}>
Subject: {email.subject}

{email.body_full or email.body_preview}

Focus on:
- What the sender wants/needs
- Any action items or deadlines
- Key information

Be direct and factual."""
        }]

        try:
            summary = self.call_claude(messages, max_tokens=200, temperature=0.3)
            return summary.strip()
        except Exception as e:
            logger.error(f"Summarization failed: {e}")
            # Fallback to body preview
            return email.body_preview[:300]

    async def generate_draft(
        self,
        email: EmailRecord,
        summary: str,
        mode: Optional[DraftMode] = None
    ) -> str:
        """
        Generate a reply draft.

        Args:
            email: The email to reply to
            summary: The email summary
            mode: Draft mode (defaults to email's draft_mode)

        Returns:
            Draft reply string
        """
        mode = mode or email.draft_mode

        mode_instructions = {
            DraftMode.PROFESSIONAL: "Write a formal, professional response. Use proper business language.",
            DraftMode.FRIENDLY: "Write a warm, friendly but still professional response. Be conversational.",
            DraftMode.BRIEF: "Write a very brief acknowledgment. Keep it under 3 sentences.",
            DraftMode.DETAILED: "Write a comprehensive response addressing all points raised."
        }

        messages = [{
            "role": "user",
            "content": f"""Generate a reply to this email:

From: {email.sender_name or email.sender_email} <{email.sender_email}>
Subject: {email.subject}

Summary: {summary}

Full Content:
{email.body_full or email.body_preview}

---

Instructions:
{mode_instructions.get(mode, mode_instructions[DraftMode.PROFESSIONAL])}

Important rules:
- Start with an appropriate greeting
- Do NOT make specific commitments about dates, times, or amounts
- If you need more information, say you'll follow up
- If there's a question you can't answer, acknowledge it
- End with an appropriate sign-off
- Use "David" as the sender name

Only output the email reply, nothing else."""
        }]

        try:
            draft = self.call_claude(messages, max_tokens=1000, temperature=0.7)
            return draft.strip()
        except Exception as e:
            logger.error(f"Draft generation failed: {e}")
            return self._fallback_draft(email)

    async def edit_draft(
        self,
        email: EmailRecord,
        edit_instructions: str
    ) -> str:
        """
        Edit an existing draft based on user instructions.

        Args:
            email: The email with the current draft
            edit_instructions: User's editing instructions

        Returns:
            Updated draft string
        """
        if not email.current_draft:
            return await self.generate_draft(email, email.summary or "")

        messages = [{
            "role": "user",
            "content": f"""Edit this email draft based on the instructions:

ORIGINAL EMAIL:
From: {email.sender_name or email.sender_email}
Subject: {email.subject}

CURRENT DRAFT:
{email.current_draft}

EDIT INSTRUCTIONS:
{edit_instructions}

---

Apply the requested changes while maintaining professionalism.
Only output the revised email, nothing else."""
        }]

        try:
            edited = self.call_claude(messages, max_tokens=1000, temperature=0.5)

            self.log_action(
                "draft_edited",
                email_id=email.id,
                details={
                    "instructions": edit_instructions[:100],
                    "new_length": len(edited)
                }
            )

            return edited.strip()
        except Exception as e:
            logger.error(f"Draft editing failed: {e}")
            return email.current_draft

    async def rewrite_draft(
        self,
        email: EmailRecord,
        new_mode: Optional[DraftMode] = None
    ) -> str:
        """
        Generate a completely new draft.

        Args:
            email: The email to reply to
            new_mode: Optional new draft mode

        Returns:
            New draft string
        """
        # Cycle through modes if not specified
        if new_mode is None:
            modes = list(DraftMode)
            current_idx = modes.index(email.draft_mode)
            new_mode = modes[(current_idx + 1) % len(modes)]

        email.draft_mode = new_mode

        new_draft = await self.generate_draft(
            email,
            email.summary or "",
            mode=new_mode
        )

        self.log_action(
            "draft_rewritten",
            email_id=email.id,
            details={"new_mode": new_mode.value}
        )

        return new_draft

    def _fallback_draft(self, email: EmailRecord) -> str:
        """Generate a simple fallback draft when AI fails."""
        sender_name = email.sender_name or email.sender_email.split("@")[0]

        return f"""Hi {sender_name},

Thank you for your email regarding "{email.subject}".

I've received your message and will review it shortly. I'll get back to you with a detailed response as soon as possible.

Best regards,
David"""

    async def suggest_tone(self, email: EmailRecord) -> DraftMode:
        """
        Analyze email and suggest appropriate tone.

        Args:
            email: The email to analyze

        Returns:
            Suggested DraftMode
        """
        messages = [{
            "role": "user",
            "content": f"""Analyze this email and suggest the most appropriate reply tone:

From: {email.sender_email}
Subject: {email.subject}
Preview: {email.body_preview[:200]}

Options:
1. PROFESSIONAL - Formal business tone
2. FRIENDLY - Warm but professional
3. BRIEF - Short acknowledgment
4. DETAILED - Comprehensive response

Just respond with one word: PROFESSIONAL, FRIENDLY, BRIEF, or DETAILED"""
        }]

        try:
            response = self.call_claude(messages, max_tokens=20, temperature=0.3)
            mode_str = response.strip().upper()

            mode_map = {
                "PROFESSIONAL": DraftMode.PROFESSIONAL,
                "FRIENDLY": DraftMode.FRIENDLY,
                "BRIEF": DraftMode.BRIEF,
                "DETAILED": DraftMode.DETAILED
            }

            return mode_map.get(mode_str, DraftMode.PROFESSIONAL)
        except Exception:
            return DraftMode.PROFESSIONAL
