"""
API request/response schemas.
"""

from datetime import datetime
from typing import Optional, List, Dict, Any
from pydantic import BaseModel, Field


# Email schemas
class EmailSummary(BaseModel):
    """Summary view of an email."""
    id: str
    subject: str
    sender_email: str
    sender_name: Optional[str] = None
    state: str
    category: Optional[str] = None
    priority: int = 3
    spam_score: int = 0
    received_at: datetime
    has_draft: bool = False
    approval_token: Optional[str] = None


class EmailDetail(BaseModel):
    """Full email details."""
    id: str
    message_id: str
    mailbox: str
    thread_id: Optional[str] = None
    sender_email: str
    sender_name: Optional[str] = None
    to_recipients: List[str] = []
    cc_recipients: List[str] = []
    subject: str
    body_preview: str = ""
    body_full: Optional[str] = None
    received_at: datetime
    has_attachments: bool = False
    importance: str = "normal"
    state: str
    category: Optional[str] = None
    priority: int = 3
    spam_score: int = 0
    summary: Optional[str] = None
    current_draft: Optional[str] = None
    draft_versions: List[str] = []
    approval_token: Optional[str] = None
    created_at: datetime
    updated_at: datetime
    sent_at: Optional[datetime] = None
    handled_by: str = "pending"
    error_message: Optional[str] = None


class EmailListResponse(BaseModel):
    """Paginated list of emails."""
    emails: List[EmailSummary]
    total: int
    page: int
    page_size: int
    has_more: bool


# Audit log schemas
class AuditEntry(BaseModel):
    """Audit log entry."""
    id: str
    email_id: Optional[str] = None
    timestamp: datetime
    agent: str
    action: str
    details: Dict[str, Any] = {}
    user_command: Optional[str] = None
    success: bool = True
    error: Optional[str] = None


class AuditLogResponse(BaseModel):
    """Paginated audit log."""
    entries: List[AuditEntry]
    total: int
    page: int
    page_size: int


# Stats schemas
class StatsResponse(BaseModel):
    """Email processing statistics."""
    total_emails: int = 0
    by_state: Dict[str, int] = {}
    emails_sent: int = 0
    spam_filtered: int = 0
    pending_count: int = 0
    period_hours: int = 24


# Action schemas
class ApproveRequest(BaseModel):
    """Request to approve an email draft."""
    email_id: str


class EditDraftRequest(BaseModel):
    """Request to edit an email draft."""
    email_id: str
    changes: str = Field(..., description="Instructions for how to modify the draft")


class MarkSpamRequest(BaseModel):
    """Request to mark email as spam."""
    email_id: str
    learn_pattern: bool = True


class IgnoreRequest(BaseModel):
    """Request to ignore an email."""
    email_id: str


class ActionResponse(BaseModel):
    """Response for email actions."""
    success: bool
    message: str
    email_id: str
    new_state: Optional[str] = None


# Spam rule schemas
class SpamRuleResponse(BaseModel):
    """Spam rule details."""
    id: str
    rule_type: str
    pattern: str
    action: str
    confidence: int
    hit_count: int
    false_positives: int
    created_at: datetime
    last_hit: Optional[datetime] = None
    is_active: bool


class SpamRulesResponse(BaseModel):
    """List of spam rules."""
    rules: List[SpamRuleResponse]


class CreateSpamRuleRequest(BaseModel):
    """Request to create a spam rule."""
    rule_type: str = Field(..., description="Type: sender, domain, keyword, pattern")
    pattern: str = Field(..., description="The pattern to match")
    action: str = Field(default="archive", description="Action: archive, delete, digest")
    confidence: int = Field(default=80, ge=0, le=100)


# Muted sender schemas
class MutedSenderResponse(BaseModel):
    """Muted sender details."""
    pattern: str
    reason: Optional[str] = None
    muted_at: datetime


class MutedSendersResponse(BaseModel):
    """List of muted senders."""
    senders: List[MutedSenderResponse]


class MuteSenderRequest(BaseModel):
    """Request to mute a sender."""
    pattern: str = Field(..., description="Email address or domain to mute")
    reason: Optional[str] = Field(None, description="Reason for muting")


# Settings schema
class SettingsResponse(BaseModel):
    """Current application settings (read-only view)."""
    poll_interval_seconds: int
    mailbox_email: str
    shared_mailbox_emails: List[str] = []
    teams_notify_urgent: bool
    teams_morning_summary_hour: int
    auto_send_enabled: bool
    max_emails_per_hour: int
    agent_model: str
    db_path: str
    calendar_integration_enabled: bool


# Email Rule schemas (LLM-based routing)
class EmailRuleResponse(BaseModel):
    """Email rule details."""
    id: str
    name: str
    description: str = ""
    match_prompt: str
    action: str
    action_value: str = ""
    priority: int = 50
    is_active: bool = True
    stop_processing: bool = True
    hit_count: int = 0
    last_hit: Optional[datetime] = None
    false_positives: int = 0
    created_at: datetime
    updated_at: datetime


class EmailRulesResponse(BaseModel):
    """List of email rules."""
    rules: List[EmailRuleResponse]


class CreateEmailRuleRequest(BaseModel):
    """Request to create an email rule."""
    name: str = Field(..., description="Human-readable name for the rule")
    description: str = Field(default="", description="What this rule does")
    match_prompt: str = Field(..., description="Natural language condition to match emails (e.g., 'Invoices from subscription services like iTunes, Netflix')")
    action: str = Field(default="move_to_folder", description="Action: move_to_folder, archive, forward, set_priority, add_label, notify")
    action_value: str = Field(default="", description="Folder name, email address, priority value, etc.")
    priority: int = Field(default=50, ge=1, le=100, description="Rule priority (lower = evaluated first)")
    stop_processing: bool = Field(default=True, description="Stop processing other rules if this matches")


class UpdateEmailRuleRequest(BaseModel):
    """Request to update an email rule."""
    name: Optional[str] = None
    description: Optional[str] = None
    match_prompt: Optional[str] = None
    action: Optional[str] = None
    action_value: Optional[str] = None
    priority: Optional[int] = Field(default=None, ge=1, le=100)
    is_active: Optional[bool] = None
    stop_processing: Optional[bool] = None


class TestRuleRequest(BaseModel):
    """Request to test a rule against recent emails."""
    match_prompt: str = Field(..., description="The match condition to test")
    limit: int = Field(default=10, ge=1, le=50, description="Number of emails to test against")
