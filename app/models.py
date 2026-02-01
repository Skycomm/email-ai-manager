"""
Data models and state machine definitions.
"""

from enum import Enum
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional, List, Dict, Any
import uuid
import secrets
import json


class EmailState(Enum):
    """State machine states for email processing."""
    NEW = "new"
    PROCESSING = "processing"
    SPAM_DETECTED = "spam_detected"
    FYI_NOTIFIED = "fyi_notified"
    ACTION_REQUIRED = "action_required"
    DRAFT_GENERATED = "draft_generated"
    AWAITING_APPROVAL = "awaiting_approval"
    APPROVED = "approved"
    SENT = "sent"
    IGNORED = "ignored"
    FORWARD_SUGGESTED = "forward_suggested"
    FORWARDED = "forwarded"
    ARCHIVED = "archived"
    FOLLOW_UP = "follow_up"  # Needs follow-up reminder
    ACKNOWLEDGED = "acknowledged"  # Seen but no action needed
    ERROR = "error"


class EmailCategory(Enum):
    """Email categorization."""
    URGENT = "urgent"
    ACTION_REQUIRED = "action_required"
    FYI = "fyi"
    MEETING = "meeting"
    SPAM_CANDIDATE = "spam_candidate"  # Hard spam - auto-delete
    NEWSLETTER = "newsletter"  # Newsletters/promotional - group and show
    FORWARD_CANDIDATE = "forward_candidate"


class DraftMode(Enum):
    """Tone/style for draft replies."""
    PROFESSIONAL = "professional"
    FRIENDLY = "friendly"
    BRIEF = "brief"
    DETAILED = "detailed"


class CommandType(Enum):
    """User commands from Teams."""
    APPROVE = "approve"
    SEND = "send"
    YES = "yes"
    CONFIRM_SEND = "confirm_send"  # Explicit confirmation required to actually send
    EDIT = "edit"
    REWRITE = "rewrite"
    MORE = "more"
    IGNORE = "ignore"
    SKIP = "skip"
    DONE = "done"
    FORWARD = "forward"
    DELETE = "delete"
    SPAM = "spam"
    MUTE = "mute"  # Never show this sender again
    FOLLOWUP = "followup"  # Mark for follow-up reminder
    # Batch commands for morning summary
    ARCHIVE_ALL = "archive_all"  # Archive/acknowledge all emails in summary
    DISMISS_ALL = "dismiss_all"
    REVIEW = "review"
    KEEP = "keep"
    UNKNOWN = "unknown"


@dataclass
class EmailRecord:
    """Represents a tracked email."""
    id: str
    message_id: str  # MS365 message ID
    mailbox: str  # Which mailbox this came from
    thread_id: Optional[str] = None

    # Email metadata
    sender_email: str = ""
    sender_name: Optional[str] = None
    to_recipients: List[str] = field(default_factory=list)
    cc_recipients: List[str] = field(default_factory=list)
    subject: str = ""
    body_preview: str = ""  # First 500 chars
    body_full: Optional[str] = None  # Full body (may be encrypted)
    received_at: datetime = field(default_factory=datetime.utcnow)
    has_attachments: bool = False
    importance: str = "normal"  # low/normal/high

    # AI Processing
    state: EmailState = EmailState.NEW
    category: Optional[EmailCategory] = None
    priority: int = 3  # 1-5, 1 being highest
    spam_score: int = 0  # 0-100
    summary: Optional[str] = None
    is_vip: bool = False  # Sender is on VIP list
    thread_context: Optional[str] = None  # Summary of previous messages in thread
    auto_send_eligible: bool = False  # Can be auto-sent without approval

    # Draft handling
    current_draft: Optional[str] = None
    draft_versions: List[str] = field(default_factory=list)
    draft_mode: DraftMode = DraftMode.PROFESSIONAL
    approval_token: Optional[str] = None

    # Teams tracking
    teams_message_id: Optional[str] = None
    teams_thread_id: Optional[str] = None

    # Audit
    created_at: datetime = field(default_factory=datetime.utcnow)
    updated_at: datetime = field(default_factory=datetime.utcnow)
    sent_at: Optional[datetime] = None
    handled_by: str = "pending"  # "ai", "user", "pending"

    # Follow-up tracking
    follow_up_at: Optional[datetime] = None  # When to remind
    follow_up_note: Optional[str] = None  # User's note about why
    follow_up_reminded_count: int = 0  # How many times reminded

    # Error handling
    error_message: Optional[str] = None
    retry_count: int = 0

    @classmethod
    def create(cls, message_id: str, mailbox: str, **kwargs) -> "EmailRecord":
        """Factory method to create a new EmailRecord."""
        return cls(
            id=str(uuid.uuid4()),
            message_id=message_id,
            mailbox=mailbox,
            **kwargs
        )

    def generate_approval_token(self) -> str:
        """Generate a short approval token."""
        self.approval_token = secrets.token_hex(3)  # 6 character hex
        return self.approval_token

    def add_draft_version(self, draft: str) -> None:
        """Add a draft version to history."""
        if self.current_draft:
            self.draft_versions.append(self.current_draft)
        self.current_draft = draft
        self.updated_at = datetime.utcnow()

    def transition_to(self, new_state: EmailState) -> None:
        """Transition to a new state with validation."""
        # Define valid transitions
        valid_transitions = {
            EmailState.NEW: [EmailState.PROCESSING],
            EmailState.PROCESSING: [
                EmailState.SPAM_DETECTED,
                EmailState.FYI_NOTIFIED,
                EmailState.ACTION_REQUIRED,
                EmailState.ERROR
            ],
            EmailState.SPAM_DETECTED: [EmailState.ARCHIVED, EmailState.ACTION_REQUIRED],
            EmailState.FYI_NOTIFIED: [EmailState.ARCHIVED, EmailState.ACTION_REQUIRED],
            EmailState.ACTION_REQUIRED: [
                EmailState.DRAFT_GENERATED,
                EmailState.FORWARD_SUGGESTED,
                EmailState.IGNORED
            ],
            EmailState.DRAFT_GENERATED: [EmailState.AWAITING_APPROVAL],
            EmailState.AWAITING_APPROVAL: [
                EmailState.APPROVED,
                EmailState.DRAFT_GENERATED,  # Re-edit
                EmailState.IGNORED,
                EmailState.SPAM_DETECTED,  # User marks as spam
                EmailState.ARCHIVED  # User dismisses
            ],
            EmailState.APPROVED: [EmailState.SENT, EmailState.ERROR],
            EmailState.FORWARD_SUGGESTED: [EmailState.FORWARDED, EmailState.IGNORED],
        }

        allowed = valid_transitions.get(self.state, [])
        if new_state not in allowed and new_state != EmailState.ERROR:
            raise ValueError(
                f"Invalid state transition: {self.state.value} -> {new_state.value}"
            )

        self.state = new_state
        self.updated_at = datetime.utcnow()

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for storage."""
        return {
            "id": self.id,
            "message_id": self.message_id,
            "mailbox": self.mailbox,
            "thread_id": self.thread_id,
            "sender_email": self.sender_email,
            "sender_name": self.sender_name,
            "to_recipients": json.dumps(self.to_recipients),
            "cc_recipients": json.dumps(self.cc_recipients),
            "subject": self.subject,
            "body_preview": self.body_preview,
            "body_full": self.body_full,
            "received_at": self.received_at.isoformat(),
            "has_attachments": self.has_attachments,
            "importance": self.importance,
            "state": self.state.value,
            "category": self.category.value if self.category else None,
            "priority": self.priority,
            "spam_score": self.spam_score,
            "summary": self.summary,
            "is_vip": self.is_vip,
            "thread_context": self.thread_context,
            "auto_send_eligible": self.auto_send_eligible,
            "current_draft": self.current_draft,
            "draft_versions": json.dumps(self.draft_versions),
            "draft_mode": self.draft_mode.value,
            "approval_token": self.approval_token,
            "teams_message_id": self.teams_message_id,
            "teams_thread_id": self.teams_thread_id,
            "created_at": self.created_at.isoformat(),
            "updated_at": self.updated_at.isoformat(),
            "sent_at": self.sent_at.isoformat() if self.sent_at else None,
            "handled_by": self.handled_by,
            "follow_up_at": self.follow_up_at.isoformat() if self.follow_up_at else None,
            "follow_up_note": self.follow_up_note,
            "follow_up_reminded_count": self.follow_up_reminded_count,
            "error_message": self.error_message,
            "retry_count": self.retry_count,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "EmailRecord":
        """Create from dictionary."""
        return cls(
            id=data["id"],
            message_id=data["message_id"],
            mailbox=data["mailbox"],
            thread_id=data.get("thread_id"),
            sender_email=data["sender_email"],
            sender_name=data.get("sender_name"),
            to_recipients=json.loads(data.get("to_recipients", "[]")),
            cc_recipients=json.loads(data.get("cc_recipients", "[]")),
            subject=data["subject"],
            body_preview=data.get("body_preview", ""),
            body_full=data.get("body_full"),
            received_at=datetime.fromisoformat(data["received_at"]),
            has_attachments=data.get("has_attachments", False),
            importance=data.get("importance", "normal"),
            state=EmailState(data["state"]),
            category=EmailCategory(data["category"]) if data.get("category") else None,
            priority=data.get("priority", 3),
            spam_score=data.get("spam_score", 0),
            summary=data.get("summary"),
            is_vip=data.get("is_vip", False),
            thread_context=data.get("thread_context"),
            auto_send_eligible=data.get("auto_send_eligible", False),
            current_draft=data.get("current_draft"),
            draft_versions=json.loads(data.get("draft_versions", "[]")),
            draft_mode=DraftMode(data.get("draft_mode", "professional")),
            approval_token=data.get("approval_token"),
            teams_message_id=data.get("teams_message_id"),
            teams_thread_id=data.get("teams_thread_id"),
            created_at=datetime.fromisoformat(data["created_at"]),
            updated_at=datetime.fromisoformat(data["updated_at"]),
            sent_at=datetime.fromisoformat(data["sent_at"]) if data.get("sent_at") else None,
            handled_by=data.get("handled_by", "pending"),
            follow_up_at=datetime.fromisoformat(data["follow_up_at"]) if data.get("follow_up_at") else None,
            follow_up_note=data.get("follow_up_note"),
            follow_up_reminded_count=data.get("follow_up_reminded_count", 0),
            error_message=data.get("error_message"),
            retry_count=data.get("retry_count", 0),
        )


@dataclass
class AuditLogEntry:
    """Audit log for tracking all actions."""
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    email_id: Optional[str] = None
    timestamp: datetime = field(default_factory=datetime.utcnow)
    agent: str = ""  # Which agent took action
    action: str = ""  # What was done
    details: Dict[str, Any] = field(default_factory=dict)
    user_command: Optional[str] = None
    success: bool = True
    error: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for storage."""
        return {
            "id": self.id,
            "email_id": self.email_id,
            "timestamp": self.timestamp.isoformat(),
            "agent": self.agent,
            "action": self.action,
            "details": json.dumps(self.details),
            "user_command": self.user_command,
            "success": self.success,
            "error": self.error,
        }


@dataclass
class SpamRule:
    """Rule for spam filtering (learned from user behavior)."""
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    rule_type: str = ""  # "sender", "domain", "subject_keyword", "pattern"
    pattern: str = ""
    action: str = "archive"  # "archive", "delete", "digest"
    confidence: int = 50  # 0-100
    hit_count: int = 0
    false_positives: int = 0
    created_at: datetime = field(default_factory=datetime.utcnow)
    last_hit: Optional[datetime] = None
    is_active: bool = True

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for storage."""
        return {
            "id": self.id,
            "rule_type": self.rule_type,
            "pattern": self.pattern,
            "action": self.action,
            "confidence": self.confidence,
            "hit_count": self.hit_count,
            "false_positives": self.false_positives,
            "created_at": self.created_at.isoformat(),
            "last_hit": self.last_hit.isoformat() if self.last_hit else None,
            "is_active": self.is_active,
        }


@dataclass
class ProcessedMessage:
    """Tracks which MS365 message IDs have been processed."""
    message_id: str
    mailbox: str
    processed_at: datetime = field(default_factory=datetime.utcnow)


class RuleAction(Enum):
    """Actions that can be taken by email rules."""
    MOVE_TO_FOLDER = "move_to_folder"  # Move email to a specific MS365 folder
    ADD_LABEL = "add_label"  # Add a category/label
    SET_PRIORITY = "set_priority"  # Override priority
    FORWARD = "forward"  # Forward to another email
    ARCHIVE = "archive"  # Archive immediately
    NOTIFY = "notify"  # Send Teams notification with custom message


@dataclass
class EmailRule:
    """
    LLM-based email rule for smart routing.

    Uses natural language prompts to match emails and take actions.
    Examples:
    - "Invoices from subscription services like iTunes, Netflix, Spotify"
    - "Meeting confirmations from Calendly or booking systems"
    - "Shipping notifications and order confirmations"
    """
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    name: str = ""  # Human-readable name
    description: str = ""  # What this rule does

    # LLM matching criteria (natural language prompt)
    match_prompt: str = ""  # e.g., "Invoices from subscription services"

    # Action configuration
    action: RuleAction = RuleAction.MOVE_TO_FOLDER
    action_value: str = ""  # Folder name, email address, priority value, etc.

    # Rule settings
    priority: int = 50  # Lower = higher priority (evaluated first)
    is_active: bool = True
    stop_processing: bool = True  # If true, don't apply more rules after this one matches

    # Stats
    hit_count: int = 0
    last_hit: Optional[datetime] = None
    false_positives: int = 0  # User corrections

    # Audit
    created_at: datetime = field(default_factory=datetime.utcnow)
    updated_at: datetime = field(default_factory=datetime.utcnow)

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for storage."""
        return {
            "id": self.id,
            "name": self.name,
            "description": self.description,
            "match_prompt": self.match_prompt,
            "action": self.action.value,
            "action_value": self.action_value,
            "priority": self.priority,
            "is_active": self.is_active,
            "stop_processing": self.stop_processing,
            "hit_count": self.hit_count,
            "last_hit": self.last_hit.isoformat() if self.last_hit else None,
            "false_positives": self.false_positives,
            "created_at": self.created_at.isoformat(),
            "updated_at": self.updated_at.isoformat(),
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "EmailRule":
        """Create from dictionary."""
        return cls(
            id=data["id"],
            name=data["name"],
            description=data.get("description", ""),
            match_prompt=data["match_prompt"],
            action=RuleAction(data["action"]),
            action_value=data.get("action_value", ""),
            priority=data.get("priority", 50),
            is_active=data.get("is_active", True),
            stop_processing=data.get("stop_processing", True),
            hit_count=data.get("hit_count", 0),
            last_hit=datetime.fromisoformat(data["last_hit"]) if data.get("last_hit") else None,
            false_positives=data.get("false_positives", 0),
            created_at=datetime.fromisoformat(data["created_at"]),
            updated_at=datetime.fromisoformat(data["updated_at"]),
        )
