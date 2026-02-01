"""
FastAPI routes for the web dashboard.
"""

import logging
from datetime import datetime
from typing import Optional, List
from fastapi import APIRouter, HTTPException, Query, Depends
from pydantic import BaseModel

from ..db import Database
from ..models import EmailState, EmailCategory, SpamRule, EmailRecord, EmailRule, RuleAction
from ..config import settings
from ..integrations.mcp_email import EmailClient
from .schemas import (
    EmailSummary, EmailDetail, EmailListResponse,
    AuditEntry, AuditLogResponse,
    StatsResponse,
    ActionResponse,
    SpamRuleResponse, SpamRulesResponse, CreateSpamRuleRequest,
    MutedSenderResponse, MutedSendersResponse, MuteSenderRequest,
    SettingsResponse,
    EmailRuleResponse, EmailRulesResponse, CreateEmailRuleRequest, UpdateEmailRuleRequest, TestRuleRequest,
)

logger = logging.getLogger(__name__)

router = APIRouter()

# Database dependency
_db: Optional[Database] = None


def get_db() -> Database:
    """Get the database instance."""
    global _db
    if _db is None:
        _db = Database(settings.db_path)
    return _db


def set_db(db: Database):
    """Set the database instance (for dependency injection)."""
    global _db
    _db = db


# Helper functions
def email_to_summary(email: EmailRecord) -> EmailSummary:
    """Convert EmailRecord to EmailSummary."""
    return EmailSummary(
        id=email.id,
        subject=email.subject,
        sender_email=email.sender_email,
        sender_name=email.sender_name,
        state=email.state.value,
        category=email.category.value if email.category else None,
        priority=email.priority,
        spam_score=email.spam_score,
        received_at=email.received_at,
        has_draft=bool(email.current_draft),
        approval_token=email.approval_token,
    )


def email_to_detail(email: EmailRecord) -> EmailDetail:
    """Convert EmailRecord to EmailDetail."""
    return EmailDetail(
        id=email.id,
        message_id=email.message_id,
        mailbox=email.mailbox,
        thread_id=email.thread_id,
        sender_email=email.sender_email,
        sender_name=email.sender_name,
        to_recipients=email.to_recipients,
        cc_recipients=email.cc_recipients,
        subject=email.subject,
        body_preview=email.body_preview,
        body_full=email.body_full,
        received_at=email.received_at,
        has_attachments=email.has_attachments,
        importance=email.importance,
        state=email.state.value,
        category=email.category.value if email.category else None,
        priority=email.priority,
        spam_score=email.spam_score,
        summary=email.summary,
        current_draft=email.current_draft,
        draft_versions=email.draft_versions,
        approval_token=email.approval_token,
        created_at=email.created_at,
        updated_at=email.updated_at,
        sent_at=email.sent_at,
        handled_by=email.handled_by,
        error_message=email.error_message,
    )


# Email endpoints
@router.get("/emails", response_model=EmailListResponse)
async def list_emails(
    state: Optional[str] = Query(None, description="Filter by state"),
    category: Optional[str] = Query(None, description="Filter by category"),
    page: int = Query(1, ge=1),
    page_size: int = Query(20, ge=1, le=100),
    db: Database = Depends(get_db),
):
    """List emails with optional filtering."""
    # Get all pending emails for now (can be expanded later)
    if state:
        try:
            state_enum = EmailState(state)
            emails = db.get_emails_by_state(state_enum, limit=page_size * 10)
        except ValueError:
            raise HTTPException(status_code=400, detail=f"Invalid state: {state}")
    else:
        # Get recent emails + any pending/action emails regardless of age
        recent_emails = db.get_recent_emails(hours=168, limit=page_size * 10)  # Last week
        pending_emails = db.get_pending_emails()  # All pending regardless of age

        # Merge and deduplicate
        seen_ids = set()
        emails = []
        for e in pending_emails + recent_emails:
            if e.id not in seen_ids:
                seen_ids.add(e.id)
                emails.append(e)

        # Sort by received_at descending
        emails.sort(key=lambda x: x.received_at or "", reverse=True)

    # Filter by category if specified
    if category:
        try:
            cat_enum = EmailCategory(category)
            emails = [e for e in emails if e.category == cat_enum]
        except ValueError:
            raise HTTPException(status_code=400, detail=f"Invalid category: {category}")

    # Exclude spam/deleted emails from list (unless specifically filtering for them)
    if state not in ['spam_detected', 'ignored']:
        emails = [e for e in emails if e.state not in [EmailState.SPAM_DETECTED, EmailState.IGNORED]]

    # Paginate
    total = len(emails)
    start = (page - 1) * page_size
    end = start + page_size
    page_emails = emails[start:end]

    return EmailListResponse(
        emails=[email_to_summary(e) for e in page_emails],
        total=total,
        page=page,
        page_size=page_size,
        has_more=end < total,
    )


@router.get("/emails/pending", response_model=EmailListResponse)
async def list_pending_emails(
    page: int = Query(1, ge=1),
    page_size: int = Query(20, ge=1, le=100),
    db: Database = Depends(get_db),
):
    """List emails awaiting action."""
    emails = db.get_pending_emails()

    total = len(emails)
    start = (page - 1) * page_size
    end = start + page_size
    page_emails = emails[start:end]

    return EmailListResponse(
        emails=[email_to_summary(e) for e in page_emails],
        total=total,
        page=page,
        page_size=page_size,
        has_more=end < total,
    )


@router.get("/emails/followups", response_model=EmailListResponse)
async def list_followup_emails(
    page: int = Query(1, ge=1),
    page_size: int = Query(20, ge=1, le=100),
    db: Database = Depends(get_db),
):
    """List emails marked for follow-up."""
    emails = db.get_all_followups()

    total = len(emails)
    start = (page - 1) * page_size
    end = start + page_size
    page_emails = emails[start:end]

    return EmailListResponse(
        emails=[email_to_summary(e) for e in page_emails],
        total=total,
        page=page,
        page_size=page_size,
        has_more=end < total,
    )


@router.get("/emails/{email_id}", response_model=EmailDetail)
async def get_email(email_id: str, db: Database = Depends(get_db)):
    """Get full email details."""
    email = db.get_email(email_id)
    if not email:
        raise HTTPException(status_code=404, detail="Email not found")
    return email_to_detail(email)


# Email actions
@router.post("/emails/{email_id}/approve", response_model=ActionResponse)
async def approve_email(email_id: str, db: Database = Depends(get_db)):
    """Approve and send an email draft."""
    email = db.get_email(email_id)
    if not email:
        raise HTTPException(status_code=404, detail="Email not found")

    if email.state != EmailState.AWAITING_APPROVAL:
        raise HTTPException(
            status_code=400,
            detail=f"Email is not awaiting approval (state: {email.state.value})"
        )

    if not email.current_draft:
        raise HTTPException(status_code=400, detail="No draft to send")

    # Note: Actual sending happens through the coordinator agent
    # This just marks it for sending
    try:
        email.transition_to(EmailState.APPROVED)
        db.save_email(email)

        return ActionResponse(
            success=True,
            message="Email approved for sending",
            email_id=email_id,
            new_state=email.state.value,
        )
    except Exception as e:
        logger.error(f"Error approving email {email_id}: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/emails/{email_id}/ignore", response_model=ActionResponse)
async def ignore_email(email_id: str, db: Database = Depends(get_db)):
    """Mark an email as ignored/dismissed - hides from dashboard without deleting from MS365."""
    email = db.get_email(email_id)
    if not email:
        raise HTTPException(status_code=404, detail="Email not found")

    try:
        # Just mark as archived in the database - don't touch MS365
        email.state = EmailState.ARCHIVED
        email.handled_by = "user"
        db.save_email(email)

        return ActionResponse(
            success=True,
            message="Email dismissed from dashboard",
            email_id=email_id,
            new_state="archived",
        )
    except Exception as e:
        logger.error(f"Error ignoring email {email_id}: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/emails/{email_id}/fyi", response_model=ActionResponse)
async def mark_fyi(email_id: str, db: Database = Depends(get_db)):
    """Mark an email as FYI (informational, no action needed)."""
    email = db.get_email(email_id)
    if not email:
        raise HTTPException(status_code=404, detail="Email not found")

    try:
        email.category = EmailCategory.FYI
        email.transition_to(EmailState.FYI_NOTIFIED)
        email.transition_to(EmailState.ACKNOWLEDGED)
        email.handled_by = "user"
        db.save_email(email)

        return ActionResponse(
            success=True,
            message="Email marked as FYI",
            email_id=email_id,
            new_state=email.state.value,
        )
    except Exception as e:
        logger.error(f"Error marking email {email_id} as FYI: {e}")
        raise HTTPException(status_code=500, detail=str(e))


class RegenerateDraftRequest(BaseModel):
    instructions: str


@router.post("/emails/{email_id}/regenerate-draft", response_model=ActionResponse)
async def regenerate_draft(email_id: str, request: RegenerateDraftRequest, db: Database = Depends(get_db)):
    """Regenerate or create a draft reply based on user instructions."""
    email = db.get_email(email_id)
    if not email:
        raise HTTPException(status_code=404, detail="Email not found")

    try:
        import anthropic

        client = anthropic.Anthropic()

        # Build context for Claude
        prompt = f"""You are drafting an email reply for David at SkyComm (an IT services company).

Original email:
From: {email.sender_name or email.sender_email} <{email.sender_email}>
Subject: {email.subject}
Body:
{email.body_full or email.body_preview}

{f"Current draft reply:{chr(10)}{email.current_draft}" if email.current_draft else "No draft yet."}

User's instructions: {request.instructions}

Write a professional email reply based on the user's instructions. Keep it concise and appropriate for a business context.
Just output the email body text - no subject line, no "Subject:" prefix, no email headers. Start directly with the greeting."""

        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1024,
            messages=[{"role": "user", "content": prompt}]
        )

        new_draft = response.content[0].text

        # Update the email with new draft
        email.current_draft = new_draft
        if email.state == EmailState.NEW:
            email.transition_to(EmailState.PROCESSING)
        email.transition_to(EmailState.AWAITING_APPROVAL)
        db.save_email(email)

        return ActionResponse(
            success=True,
            message="Draft updated",
            email_id=email_id,
            new_state=email.state.value,
        )
    except Exception as e:
        logger.error(f"Error regenerating draft for {email_id}: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/emails/{email_id}/dismiss", response_model=ActionResponse)
async def dismiss_email(email_id: str, db: Database = Depends(get_db)):
    """Alias for ignore - hides from dashboard without deleting from MS365."""
    return await ignore_email(email_id, db)


@router.post("/emails/{email_id}/delete", response_model=ActionResponse)
async def delete_email(email_id: str, db: Database = Depends(get_db)):
    """Delete a single email without blocking sender."""
    email = db.get_email(email_id)
    if not email:
        raise HTTPException(status_code=404, detail="Email not found")

    try:
        email_client = EmailClient()

        # Delete from MS365 - this MUST succeed
        deleted = email_client.delete_email(email.message_id, email.mailbox)
        if not deleted:
            logger.error(f"Failed to delete email {email_id} from MS365")
            return ActionResponse(
                success=False,
                message="Failed to delete email from mailbox",
                email_id=email_id,
                new_state=email.state.value,
            )

        # Only mark as ignored AFTER successful MS365 delete
        email.state = EmailState.IGNORED
        db.save_email(email)

        return ActionResponse(
            success=True,
            message="Email deleted",
            email_id=email_id,
            new_state="deleted",
        )
    except Exception as e:
        logger.error(f"Error deleting email {email_id}: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/emails/{email_id}/spam", response_model=ActionResponse)
async def mark_spam(email_id: str, db: Database = Depends(get_db)):
    """Mark an email as spam, delete it from MS365, and block sender."""
    email = db.get_email(email_id)
    if not email:
        raise HTTPException(status_code=404, detail="Email not found")

    try:
        email_client = EmailClient()
        deleted_count = 0

        # Extract sender domain for finding similar emails
        sender_domain = None
        if '@' in email.sender_email.lower():
            sender_domain = email.sender_email.lower().split('@')[1]

        # Learn spam pattern - create/update domain rule
        if sender_domain:
            existing_rules = db.get_active_spam_rules()
            domain_rule = next(
                (r for r in existing_rules if r.rule_type == 'domain' and r.pattern == sender_domain),
                None
            )

            if domain_rule:
                db.increment_spam_rule_hit(domain_rule.id)
                logger.info(f"Increased confidence for spam domain rule: {sender_domain}")
            else:
                new_rule = SpamRule(
                    rule_type='domain',
                    pattern=sender_domain,
                    action='delete',  # Delete action for user-marked spam
                    confidence=80,  # Higher confidence since user marked it
                )
                db.save_spam_rule(new_rule)
                logger.info(f"Created new spam domain rule: {sender_domain}")

        # Find similar emails from the same domain in our database
        if sender_domain:
            all_emails = db.get_recent_emails(hours=168, limit=500)  # Get recent emails
            similar_emails = [
                e for e in all_emails
                if e.sender_email.lower().endswith(f"@{sender_domain}")
                and e.id != email_id
                and e.state not in [EmailState.ARCHIVED, EmailState.SPAM_DETECTED]
            ]

            # Delete similar emails from MS365 and from our database
            for similar in similar_emails:
                try:
                    # Delete from MS365
                    email_client.delete_email(similar.message_id, similar.mailbox)
                    # Delete from our database
                    db.delete_email(similar.id)
                    deleted_count += 1
                    logger.info(f"Deleted similar spam from {similar.sender_email}: {similar.subject[:50]}")
                except Exception as e:
                    logger.warning(f"Could not delete similar email {similar.id}: {e}")

        # Delete the original email from MS365
        try:
            email_client.delete_email(email.message_id, email.mailbox)
            deleted_count += 1
        except Exception as e:
            logger.warning(f"Could not delete original email from MS365: {e}")

        # Delete from our database too
        db.delete_email(email_id)

        message = f"Email marked as spam and deleted"
        if deleted_count > 1:
            message += f" ({deleted_count} similar emails also deleted)"

        return ActionResponse(
            success=True,
            message=message,
            email_id=email_id,
            new_state="deleted",
        )
    except Exception as e:
        logger.error(f"Error marking email {email_id} as spam: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/emails/{email_id}/followup", response_model=ActionResponse)
async def mark_followup(
    email_id: str,
    days: int = Query(1, ge=1, le=30, description="Days until follow-up reminder"),
    note: Optional[str] = Query(None, description="Optional note"),
    db: Database = Depends(get_db),
):
    """Mark an email for follow-up with a reminder."""
    email = db.get_email(email_id)
    if not email:
        raise HTTPException(status_code=404, detail="Email not found")

    try:
        from datetime import timedelta

        email.state = EmailState.FOLLOW_UP
        email.follow_up_at = datetime.utcnow() + timedelta(days=days)
        email.follow_up_note = note
        email.handled_by = "user"
        db.save_email(email)

        return ActionResponse(
            success=True,
            message=f"Follow-up reminder set for {days} day(s)",
            email_id=email_id,
            new_state="follow_up",
        )
    except Exception as e:
        logger.error(f"Error setting follow-up for email {email_id}: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/emails/{email_id}/clear-followup", response_model=ActionResponse)
async def clear_followup(email_id: str, db: Database = Depends(get_db)):
    """Clear follow-up status (mark as done)."""
    email = db.get_email(email_id)
    if not email:
        raise HTTPException(status_code=404, detail="Email not found")

    try:
        email.state = EmailState.ARCHIVED
        email.follow_up_at = None
        email.follow_up_note = None
        email.handled_by = "user"
        db.save_email(email)

        return ActionResponse(
            success=True,
            message="Follow-up completed",
            email_id=email_id,
            new_state="archived",
        )
    except Exception as e:
        logger.error(f"Error clearing follow-up for email {email_id}: {e}")
        raise HTTPException(status_code=500, detail=str(e))


# Audit log endpoints
@router.get("/audit", response_model=AuditLogResponse)
async def get_audit_log(
    email_id: Optional[str] = Query(None, description="Filter by email ID"),
    page: int = Query(1, ge=1),
    page_size: int = Query(50, ge=1, le=200),
    db: Database = Depends(get_db),
):
    """Get audit log entries."""
    offset = (page - 1) * page_size
    entries = db.get_audit_log(email_id=email_id, limit=page_size, offset=offset)
    total = db.get_audit_log_count(email_id=email_id)

    # Convert to response format
    audit_entries = [
        AuditEntry(
            id=e.id,
            email_id=e.email_id,
            timestamp=e.timestamp,
            agent=e.agent,
            action=e.action,
            details=e.details,
            user_command=e.user_command,
            success=e.success,
            error=e.error,
        )
        for e in entries
    ]

    return AuditLogResponse(
        entries=audit_entries,
        total=total,
        page=page,
        page_size=page_size,
    )


# Stats endpoints
@router.get("/stats", response_model=StatsResponse)
async def get_stats(
    hours: int = Query(24, ge=1, le=720, description="Time period in hours"),
    db: Database = Depends(get_db),
):
    """Get email processing statistics."""
    stats = db.get_stats(hours=hours)
    pending = db.get_pending_emails()

    return StatsResponse(
        total_emails=stats.get("total_emails", 0),
        by_state=stats.get("by_state", {}),
        emails_sent=stats.get("emails_sent", 0),
        spam_filtered=stats.get("spam_filtered", 0),
        pending_count=len(pending),
        period_hours=hours,
    )


@router.get("/stats/advanced")
async def get_advanced_stats(
    hours: int = Query(24, ge=1, le=720, description="Time period in hours"),
    db: Database = Depends(get_db),
):
    """Get detailed analytics for the dashboard."""
    stats = db.get_advanced_stats(hours=hours)
    pending = db.get_pending_emails()

    return {
        **stats,
        "pending_count": len(pending),
        "period_hours": hours,
    }


@router.get("/stats/sender/{sender_email}")
async def get_sender_stats(
    sender_email: str,
    hours: int = Query(168, ge=1, le=720, description="Time period in hours (default: 1 week)"),
    db: Database = Depends(get_db),
):
    """Get statistics for a specific sender."""
    stats = db.get_sender_stats(sender_email, hours=hours)
    return {
        "sender_email": sender_email,
        "period_hours": hours,
        **stats,
    }


# Spam rules endpoints
@router.get("/spam-rules", response_model=SpamRulesResponse)
async def list_spam_rules(db: Database = Depends(get_db)):
    """List all spam rules."""
    rules = db.get_active_spam_rules()

    return SpamRulesResponse(
        rules=[
            SpamRuleResponse(
                id=r.id,
                rule_type=r.rule_type,
                pattern=r.pattern,
                action=r.action,
                confidence=r.confidence,
                hit_count=r.hit_count,
                false_positives=r.false_positives,
                created_at=r.created_at,
                last_hit=r.last_hit,
                is_active=r.is_active,
            )
            for r in rules
        ]
    )


@router.post("/spam-rules", response_model=SpamRuleResponse)
async def create_spam_rule(
    request: CreateSpamRuleRequest,
    db: Database = Depends(get_db),
):
    """Create a new spam rule."""
    rule = SpamRule(
        rule_type=request.rule_type,
        pattern=request.pattern,
        action=request.action,
        confidence=request.confidence,
    )

    db.save_spam_rule(rule)

    return SpamRuleResponse(
        id=rule.id,
        rule_type=rule.rule_type,
        pattern=rule.pattern,
        action=rule.action,
        confidence=rule.confidence,
        hit_count=rule.hit_count,
        false_positives=rule.false_positives,
        created_at=rule.created_at,
        last_hit=rule.last_hit,
        is_active=rule.is_active,
    )


# Muted Senders endpoints
@router.get("/muted-senders", response_model=MutedSendersResponse)
async def list_muted_senders(db: Database = Depends(get_db)):
    """List all muted senders."""
    senders = db.get_muted_senders()

    return MutedSendersResponse(
        senders=[
            MutedSenderResponse(
                pattern=s["pattern"],
                reason=s.get("reason"),
                muted_at=s["muted_at"],
            )
            for s in senders
        ]
    )


@router.post("/muted-senders", response_model=MutedSenderResponse)
async def mute_sender(
    request: MuteSenderRequest,
    db: Database = Depends(get_db),
):
    """Mute a sender (email or domain)."""
    db.mute_sender(request.pattern, request.reason)

    return MutedSenderResponse(
        pattern=request.pattern,
        reason=request.reason,
        muted_at=datetime.utcnow(),
    )


@router.delete("/muted-senders/{email_pattern:path}")
async def unmute_sender(email_pattern: str, db: Database = Depends(get_db)):
    """Unmute a sender."""
    db.unmute_sender(email_pattern)
    return {"success": True, "message": f"Unmuted {email_pattern}"}


# Delete spam rule endpoint
@router.delete("/spam-rules/{rule_id}")
async def delete_spam_rule(rule_id: str, db: Database = Depends(get_db)):
    """Delete a spam rule."""
    try:
        db.delete_spam_rule(rule_id)
        return {"success": True, "message": "Spam rule deleted"}
    except Exception as e:
        logger.error(f"Error deleting spam rule {rule_id}: {e}")
        raise HTTPException(status_code=500, detail=str(e))


# Settings endpoint
@router.get("/settings", response_model=SettingsResponse)
async def get_settings():
    """Get current settings (read-only view)."""
    return SettingsResponse(
        poll_interval_seconds=settings.poll_interval_seconds,
        mailbox_email=settings.mailbox_email,
        shared_mailbox_emails=settings.shared_mailbox_emails,
        teams_notify_urgent=settings.teams_notify_urgent,
        teams_morning_summary_hour=settings.teams_morning_summary_hour,
        auto_send_enabled=settings.auto_send_enabled,
        max_emails_per_hour=settings.max_emails_per_hour,
        agent_model=settings.agent_model,
        db_path=settings.db_path,
        calendar_integration_enabled=settings.calendar_integration_enabled,
    )


# Email Rules endpoints (LLM-based routing)
@router.get("/email-rules", response_model=EmailRulesResponse)
async def list_email_rules(
    include_inactive: bool = Query(False, description="Include inactive rules"),
    db: Database = Depends(get_db),
):
    """List all email routing rules."""
    if include_inactive:
        rules = db.get_all_email_rules()
    else:
        rules = db.get_active_email_rules()

    return EmailRulesResponse(
        rules=[
            EmailRuleResponse(
                id=r.id,
                name=r.name,
                description=r.description,
                match_prompt=r.match_prompt,
                action=r.action.value,
                action_value=r.action_value,
                priority=r.priority,
                is_active=r.is_active,
                stop_processing=r.stop_processing,
                hit_count=r.hit_count,
                last_hit=r.last_hit,
                false_positives=r.false_positives,
                created_at=r.created_at,
                updated_at=r.updated_at,
            )
            for r in rules
        ]
    )


@router.get("/email-rules/{rule_id}", response_model=EmailRuleResponse)
async def get_email_rule(rule_id: str, db: Database = Depends(get_db)):
    """Get a specific email rule."""
    rule = db.get_email_rule(rule_id)
    if not rule:
        raise HTTPException(status_code=404, detail="Rule not found")

    return EmailRuleResponse(
        id=rule.id,
        name=rule.name,
        description=rule.description,
        match_prompt=rule.match_prompt,
        action=rule.action.value,
        action_value=rule.action_value,
        priority=rule.priority,
        is_active=rule.is_active,
        stop_processing=rule.stop_processing,
        hit_count=rule.hit_count,
        last_hit=rule.last_hit,
        false_positives=rule.false_positives,
        created_at=rule.created_at,
        updated_at=rule.updated_at,
    )


@router.post("/email-rules", response_model=EmailRuleResponse)
async def create_email_rule(
    request: CreateEmailRuleRequest,
    db: Database = Depends(get_db),
):
    """Create a new email routing rule."""
    try:
        action = RuleAction(request.action)
    except ValueError:
        valid_actions = [a.value for a in RuleAction]
        raise HTTPException(
            status_code=400,
            detail=f"Invalid action. Valid actions: {valid_actions}"
        )

    rule = EmailRule(
        name=request.name,
        description=request.description,
        match_prompt=request.match_prompt,
        action=action,
        action_value=request.action_value,
        priority=request.priority,
        stop_processing=request.stop_processing,
    )

    db.save_email_rule(rule)
    logger.info(f"Created email rule: {rule.name} ({rule.id})")

    return EmailRuleResponse(
        id=rule.id,
        name=rule.name,
        description=rule.description,
        match_prompt=rule.match_prompt,
        action=rule.action.value,
        action_value=rule.action_value,
        priority=rule.priority,
        is_active=rule.is_active,
        stop_processing=rule.stop_processing,
        hit_count=rule.hit_count,
        last_hit=rule.last_hit,
        false_positives=rule.false_positives,
        created_at=rule.created_at,
        updated_at=rule.updated_at,
    )


@router.put("/email-rules/{rule_id}", response_model=EmailRuleResponse)
async def update_email_rule(
    rule_id: str,
    request: UpdateEmailRuleRequest,
    db: Database = Depends(get_db),
):
    """Update an existing email rule."""
    rule = db.get_email_rule(rule_id)
    if not rule:
        raise HTTPException(status_code=404, detail="Rule not found")

    # Update fields if provided
    if request.name is not None:
        rule.name = request.name
    if request.description is not None:
        rule.description = request.description
    if request.match_prompt is not None:
        rule.match_prompt = request.match_prompt
    if request.action is not None:
        try:
            rule.action = RuleAction(request.action)
        except ValueError:
            valid_actions = [a.value for a in RuleAction]
            raise HTTPException(
                status_code=400,
                detail=f"Invalid action. Valid actions: {valid_actions}"
            )
    if request.action_value is not None:
        rule.action_value = request.action_value
    if request.priority is not None:
        rule.priority = request.priority
    if request.is_active is not None:
        rule.is_active = request.is_active
    if request.stop_processing is not None:
        rule.stop_processing = request.stop_processing

    rule.updated_at = datetime.utcnow()
    db.save_email_rule(rule)
    logger.info(f"Updated email rule: {rule.name} ({rule.id})")

    return EmailRuleResponse(
        id=rule.id,
        name=rule.name,
        description=rule.description,
        match_prompt=rule.match_prompt,
        action=rule.action.value,
        action_value=rule.action_value,
        priority=rule.priority,
        is_active=rule.is_active,
        stop_processing=rule.stop_processing,
        hit_count=rule.hit_count,
        last_hit=rule.last_hit,
        false_positives=rule.false_positives,
        created_at=rule.created_at,
        updated_at=rule.updated_at,
    )


@router.delete("/email-rules/{rule_id}")
async def delete_email_rule(rule_id: str, db: Database = Depends(get_db)):
    """Delete an email rule."""
    rule = db.get_email_rule(rule_id)
    if not rule:
        raise HTTPException(status_code=404, detail="Rule not found")

    deleted = db.delete_email_rule(rule_id)
    if deleted:
        logger.info(f"Deleted email rule: {rule.name} ({rule_id})")
        return {"success": True, "message": f"Rule '{rule.name}' deleted"}
    else:
        raise HTTPException(status_code=500, detail="Failed to delete rule")


@router.post("/email-rules/test")
async def test_email_rule(
    request: TestRuleRequest,
    db: Database = Depends(get_db),
):
    """
    Test a rule condition against recent emails to see what it would match.
    This helps verify rules before creating them.
    """
    from ..agents.rules import RulesAgent

    try:
        # Create a temporary rule for testing
        test_rule = EmailRule(
            name="Test Rule",
            match_prompt=request.match_prompt,
            action=RuleAction.MOVE_TO_FOLDER,  # Doesn't matter for testing
        )

        # Get recent emails
        recent_emails = db.get_recent_emails(hours=168, limit=request.limit)

        if not recent_emails:
            return {
                "match_prompt": request.match_prompt,
                "total_tested": 0,
                "matches": [],
                "message": "No recent emails to test against"
            }

        # Initialize rules agent for testing
        rules_agent = RulesAgent(db)

        matches = []
        for email in recent_emails:
            result = await rules_agent.evaluate_rule(email, test_rule)

            if result["matches"] and result["confidence"] >= 50:
                matches.append({
                    "email_id": email.id,
                    "subject": email.subject[:60],
                    "sender": email.sender_email,
                    "confidence": result["confidence"],
                    "reason": result["reason"],
                })

        return {
            "match_prompt": request.match_prompt,
            "total_tested": len(recent_emails),
            "matches_found": len(matches),
            "matches": matches,
            "match_rate": f"{len(matches)}/{len(recent_emails)}",
        }

    except Exception as e:
        logger.error(f"Error testing rule: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/email-rules/{rule_id}/false-positive")
async def report_false_positive(rule_id: str, db: Database = Depends(get_db)):
    """Report a false positive for a rule (decreases confidence)."""
    rule = db.get_email_rule(rule_id)
    if not rule:
        raise HTTPException(status_code=404, detail="Rule not found")

    db.record_email_rule_false_positive(rule_id)

    return {
        "success": True,
        "message": f"False positive recorded for rule '{rule.name}'",
        "new_false_positive_count": rule.false_positives + 1,
    }


class RunRuleRequest(BaseModel):
    dry_run: bool = True
    limit: int = 50


from fastapi.responses import StreamingResponse
import json
import asyncio


@router.get("/email-rules/{rule_id}/run-stream")
async def run_email_rule_stream(
    rule_id: str,
    dry_run: bool = Query(True),
    limit: int = Query(50),
):
    """
    Run a rule with streaming progress updates via Server-Sent Events.
    """
    from ..agents.rules import RulesAgent

    db = get_db()
    rule = db.get_email_rule(rule_id)
    if not rule:
        raise HTTPException(status_code=404, detail="Rule not found")

    if not rule.is_active:
        raise HTTPException(status_code=400, detail="Rule is not active")

    async def generate():
        try:
            # Get recent emails
            recent_emails = db.get_recent_emails(hours=168, limit=limit)
            total = len(recent_emails)

            # Send initial info
            yield f"data: {json.dumps({'type': 'start', 'total': total, 'rule_name': rule.name})}\n\n"

            if not recent_emails:
                yield f"data: {json.dumps({'type': 'complete', 'total_evaluated': 0, 'matched': 0, 'processed': 0, 'matches': []})}\n\n"
                return

            # Initialize rules agent
            rules_agent = RulesAgent(db)

            matches = []
            errors = []
            processed = 0

            for i, email in enumerate(recent_emails):
                # Send progress update
                yield f"data: {json.dumps({'type': 'progress', 'current': i + 1, 'total': total, 'subject': email.subject[:50]})}\n\n"

                try:
                    # Evaluate single email
                    result = await rules_agent.evaluate_single_email(rule, email)

                    if result.get('matches'):
                        match_info = {
                            'email_id': email.id,
                            'subject': email.subject,
                            'sender': email.sender_email,
                            'confidence': result.get('confidence', 0),
                            'reason': result.get('reason', ''),
                        }
                        matches.append(match_info)

                        # Apply action if not dry run
                        if not dry_run:
                            try:
                                await rules_agent._execute_rule_action(email, rule)
                                processed += 1
                                yield f"data: {json.dumps({'type': 'action', 'email_id': email.id, 'subject': email.subject[:50], 'action': rule.action.value})}\n\n"
                            except Exception as e:
                                errors.append({'email_id': email.id, 'error': str(e)})

                except Exception as e:
                    errors.append({'email_id': email.id, 'error': str(e)})

            # Send final results
            yield f"data: {json.dumps({'type': 'complete', 'rule_id': rule_id, 'rule_name': rule.name, 'dry_run': dry_run, 'total_evaluated': total, 'matched': len(matches), 'processed': processed, 'errors': len(errors), 'matches': matches, 'error_details': errors})}\n\n"

        except Exception as e:
            logger.error(f"Error in rule stream: {e}")
            yield f"data: {json.dumps({'type': 'error', 'message': str(e)})}\n\n"

    return StreamingResponse(
        generate(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
        }
    )


@router.post("/email-rules/{rule_id}/run")
async def run_email_rule(
    rule_id: str,
    request: RunRuleRequest,
    db: Database = Depends(get_db),
):
    """
    Run a rule against recent emails.

    - dry_run=True (default): Only evaluate, don't apply actions
    - dry_run=False: Actually apply the rule actions (move emails, etc.)
    """
    from ..agents.rules import RulesAgent

    rule = db.get_email_rule(rule_id)
    if not rule:
        raise HTTPException(status_code=404, detail="Rule not found")

    if not rule.is_active:
        raise HTTPException(status_code=400, detail="Rule is not active")

    try:
        # Get recent emails to evaluate
        recent_emails = db.get_recent_emails(hours=168, limit=request.limit)

        if not recent_emails:
            return {
                "success": True,
                "rule_id": rule_id,
                "rule_name": rule.name,
                "message": "No recent emails to evaluate",
                "total_evaluated": 0,
                "matched": 0,
                "processed": 0,
            }

        # Initialize rules agent
        rules_agent = RulesAgent(db)

        # Run the rule
        result = await rules_agent.run_rule_on_emails(
            rule=rule,
            emails=recent_emails,
            dry_run=request.dry_run
        )

        return {
            "success": True,
            **result
        }

    except Exception as e:
        logger.error(f"Error running rule {rule_id}: {e}")
        raise HTTPException(status_code=500, detail=str(e))


# Folder contents endpoint - view emails in a specific folder
@router.get("/folder-emails/{folder_name:path}")
async def get_folder_emails(
    folder_name: str,
    mailbox: Optional[str] = Query(None, description="Mailbox email (uses primary if not specified)"),
    limit: int = Query(50, ge=1, le=200, description="Maximum emails to fetch"),
):
    """
    Get emails from a specific folder in the mailbox.
    Useful for viewing emails that rules have moved to folders like Clutter, Billing, etc.
    """
    from ..integrations.mcp_client import MCPClient
    from ..integrations.mcp_email import EmailClient

    try:
        mcp = MCPClient()
        email_client = EmailClient(mcp)
        mailbox_email = mailbox or settings.mailbox_email

        # Resolve folder name to ID
        folder_id = email_client._resolve_folder_id(folder_name, mailbox_email)

        if not folder_id:
            raise HTTPException(status_code=404, detail=f"Folder '{folder_name}' not found")

        # Fetch emails from the folder
        emails = mcp.list_mail_messages(
            mailbox=mailbox_email,
            folder=folder_id,
            top=limit
        )

        # Format for response
        formatted_emails = []
        for email in emails:
            sender = email.get("from", {}).get("emailAddress", {})
            formatted_emails.append({
                "id": email.get("id"),
                "subject": email.get("subject", "(No Subject)"),
                "sender_email": sender.get("address", ""),
                "sender_name": sender.get("name", ""),
                "received_at": email.get("receivedDateTime"),
                "body_preview": email.get("bodyPreview", "")[:200],
                "has_attachments": email.get("hasAttachments", False),
                "importance": email.get("importance", "normal"),
                "is_read": email.get("isRead", False),
            })

        return {
            "folder": folder_name,
            "mailbox": mailbox_email,
            "count": len(formatted_emails),
            "emails": formatted_emails,
        }

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error fetching emails from folder {folder_name}: {e}")
        raise HTTPException(status_code=500, detail=str(e))


class DeleteFolderEmailRequest(BaseModel):
    message_id: str
    mailbox: Optional[str] = None


@router.post("/folder-emails/delete")
async def delete_folder_email(request: DeleteFolderEmailRequest):
    """
    Delete an email by its MS365 message ID (moves to Deleted Items).
    Used for deleting emails from the folder browser.
    """
    from ..integrations.mcp_client import MCPClient

    try:
        mcp = MCPClient()
        mailbox_email = request.mailbox or settings.mailbox_email

        # Delete the email (moves to Deleted Items)
        mcp.delete_mail_message(
            message_id=request.message_id,
            sender_email=mailbox_email
        )

        return {"success": True, "message": "Email deleted"}

    except Exception as e:
        logger.error(f"Error deleting email {request.message_id}: {e}")
        raise HTTPException(status_code=500, detail=str(e))


# Email folders endpoint for folder browser
@router.get("/email-folders")
async def list_email_folders(
    mailbox: Optional[str] = Query(None, description="Mailbox email (uses primary if not specified)"),
    recursive: bool = Query(True, description="Fetch folders recursively including subfolders"),
):
    """
    List available email folders for the mailbox, including nested subfolders.
    """
    from ..integrations.mcp_client import MCPClient

    try:
        mcp = MCPClient()
        mailbox_email = mailbox or settings.mailbox_email

        if recursive:
            # Fetch all folders recursively including subfolders
            folders = mcp.list_all_mail_folders_recursive(mailbox=mailbox_email, max_depth=3)

            def sort_folders(folder_list):
                """Recursively sort folders alphabetically."""
                folder_list.sort(key=lambda f: f["name"].lower())
                for folder in folder_list:
                    if folder.get("children"):
                        sort_folders(folder["children"])
                return folder_list

            sort_folders(folders)

            return {
                "mailbox": mailbox_email,
                "folders": folders,
            }
        else:
            # Fetch only top-level folders
            raw_folders = mcp.list_mail_folders(mailbox=mailbox_email)

            # Format folders for the UI with child folder count
            formatted_folders = []
            for folder in raw_folders:
                formatted_folders.append({
                    "id": folder.get("id"),
                    "name": folder.get("displayName"),
                    "total_count": folder.get("totalItemCount", 0),
                    "unread_count": folder.get("unreadItemCount", 0),
                    "child_folder_count": folder.get("childFolderCount", 0),
                })

            # Sort alphabetically by name
            formatted_folders.sort(key=lambda f: f["name"].lower())

            return {
                "mailbox": mailbox_email,
                "folders": formatted_folders,
            }

    except Exception as e:
        logger.error(f"Error listing email folders: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to list folders: {str(e)}")


# Health check
@router.get("/health")
async def health_check(db: Database = Depends(get_db)):
    """Health check endpoint with extended status."""
    # Check MCP connectivity by testing the database
    mcp_connected = True
    try:
        # Simple DB check
        db.get_stats(hours=1)
    except Exception:
        mcp_connected = False

    # Check if Teams is configured
    teams_configured = bool(settings.teams_channel_id or settings.teams_chat_id)

    return {
        "status": "healthy",
        "service": "Email AI Manager Dashboard API",
        "timestamp": datetime.utcnow().isoformat(),
        "mcp_connected": mcp_connected,
        "teams_configured": teams_configured,
    }
