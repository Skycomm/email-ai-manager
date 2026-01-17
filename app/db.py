"""
SQLite database operations.
"""

import sqlite3
import json
import logging
from pathlib import Path
from datetime import datetime
from typing import Optional, List, Dict, Any
from contextlib import contextmanager

from .models import EmailRecord, AuditLogEntry, SpamRule, EmailState

logger = logging.getLogger(__name__)


class Database:
    """SQLite database handler."""

    def __init__(self, db_path: str):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    @contextmanager
    def _get_connection(self):
        """Get a database connection with proper handling."""
        conn = sqlite3.connect(str(self.db_path))
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        except Exception as e:
            conn.rollback()
            raise e
        finally:
            conn.close()

    def _init_db(self):
        """Initialize database schema."""
        schema_path = Path(__file__).parent.parent / "migrations" / "001_initial.sql"

        with self._get_connection() as conn:
            # Check if tables exist
            cursor = conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name='emails'"
            )
            if cursor.fetchone() is None:
                if schema_path.exists():
                    with open(schema_path) as f:
                        conn.executescript(f.read())
                    logger.info("Database schema initialized from migration file")
                else:
                    # Inline schema if migration file not found
                    self._create_schema(conn)
                    logger.info("Database schema initialized inline")

    def _create_schema(self, conn: sqlite3.Connection):
        """Create database schema inline."""
        conn.executescript("""
            -- Emails table
            CREATE TABLE IF NOT EXISTS emails (
                id TEXT PRIMARY KEY,
                message_id TEXT NOT NULL,
                mailbox TEXT NOT NULL,
                thread_id TEXT,
                sender_email TEXT NOT NULL,
                sender_name TEXT,
                to_recipients TEXT,
                cc_recipients TEXT,
                subject TEXT NOT NULL,
                body_preview TEXT,
                body_full TEXT,
                received_at TEXT NOT NULL,
                has_attachments INTEGER DEFAULT 0,
                importance TEXT DEFAULT 'normal',
                state TEXT NOT NULL DEFAULT 'new',
                category TEXT,
                priority INTEGER DEFAULT 3,
                spam_score INTEGER DEFAULT 0,
                summary TEXT,
                current_draft TEXT,
                draft_versions TEXT,
                draft_mode TEXT DEFAULT 'professional',
                approval_token TEXT,
                teams_message_id TEXT,
                teams_thread_id TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                sent_at TEXT,
                handled_by TEXT DEFAULT 'pending',
                error_message TEXT,
                retry_count INTEGER DEFAULT 0,
                UNIQUE(message_id, mailbox)
            );

            -- Audit log table
            CREATE TABLE IF NOT EXISTS audit_log (
                id TEXT PRIMARY KEY,
                email_id TEXT,
                timestamp TEXT NOT NULL,
                agent TEXT NOT NULL,
                action TEXT NOT NULL,
                details TEXT,
                user_command TEXT,
                success INTEGER DEFAULT 1,
                error TEXT,
                FOREIGN KEY (email_id) REFERENCES emails(id)
            );

            -- Spam rules table
            CREATE TABLE IF NOT EXISTS spam_rules (
                id TEXT PRIMARY KEY,
                rule_type TEXT NOT NULL,
                pattern TEXT NOT NULL,
                action TEXT DEFAULT 'archive',
                confidence INTEGER DEFAULT 50,
                hit_count INTEGER DEFAULT 0,
                false_positives INTEGER DEFAULT 0,
                created_at TEXT NOT NULL,
                last_hit TEXT,
                is_active INTEGER DEFAULT 1
            );

            -- Processed messages tracking
            CREATE TABLE IF NOT EXISTS processed_messages (
                message_id TEXT NOT NULL,
                mailbox TEXT NOT NULL,
                processed_at TEXT NOT NULL,
                PRIMARY KEY (message_id, mailbox)
            );

            -- Indexes for common queries
            CREATE INDEX IF NOT EXISTS idx_emails_state ON emails(state);
            CREATE INDEX IF NOT EXISTS idx_emails_mailbox ON emails(mailbox);
            CREATE INDEX IF NOT EXISTS idx_emails_received ON emails(received_at);
            CREATE INDEX IF NOT EXISTS idx_emails_approval_token ON emails(approval_token);
            CREATE INDEX IF NOT EXISTS idx_audit_email_id ON audit_log(email_id);
            CREATE INDEX IF NOT EXISTS idx_audit_timestamp ON audit_log(timestamp);
        """)

    # Email operations
    def save_email(self, email: EmailRecord) -> None:
        """Save or update an email record."""
        with self._get_connection() as conn:
            data = email.to_dict()
            placeholders = ", ".join(["?" for _ in data])
            columns = ", ".join(data.keys())
            updates = ", ".join([f"{k}=excluded.{k}" for k in data.keys()])

            conn.execute(f"""
                INSERT INTO emails ({columns}) VALUES ({placeholders})
                ON CONFLICT(message_id, mailbox) DO UPDATE SET {updates}
            """, list(data.values()))

    def get_email(self, email_id: str) -> Optional[EmailRecord]:
        """Get an email by internal ID."""
        with self._get_connection() as conn:
            cursor = conn.execute(
                "SELECT * FROM emails WHERE id = ?",
                (email_id,)
            )
            row = cursor.fetchone()
            if row:
                return EmailRecord.from_dict(dict(row))
        return None

    def get_email_by_message_id(self, message_id: str, mailbox: str) -> Optional[EmailRecord]:
        """Get an email by MS365 message ID."""
        with self._get_connection() as conn:
            cursor = conn.execute(
                "SELECT * FROM emails WHERE message_id = ? AND mailbox = ?",
                (message_id, mailbox)
            )
            row = cursor.fetchone()
            if row:
                return EmailRecord.from_dict(dict(row))
        return None

    def get_email_by_approval_token(self, token: str) -> Optional[EmailRecord]:
        """Get an email by approval token."""
        with self._get_connection() as conn:
            cursor = conn.execute(
                "SELECT * FROM emails WHERE approval_token = ? AND state = ?",
                (token, EmailState.AWAITING_APPROVAL.value)
            )
            row = cursor.fetchone()
            if row:
                return EmailRecord.from_dict(dict(row))
        return None

    def get_emails_by_state(self, state: EmailState, limit: int = 100) -> List[EmailRecord]:
        """Get emails in a specific state."""
        with self._get_connection() as conn:
            cursor = conn.execute(
                "SELECT * FROM emails WHERE state = ? ORDER BY received_at DESC LIMIT ?",
                (state.value, limit)
            )
            return [EmailRecord.from_dict(dict(row)) for row in cursor.fetchall()]

    def get_pending_emails(self) -> List[EmailRecord]:
        """Get all emails awaiting action."""
        pending_states = [
            EmailState.NEW.value,
            EmailState.PROCESSING.value,
            EmailState.ACTION_REQUIRED.value,
            EmailState.AWAITING_APPROVAL.value,
        ]
        with self._get_connection() as conn:
            placeholders = ", ".join(["?" for _ in pending_states])
            cursor = conn.execute(
                f"SELECT * FROM emails WHERE state IN ({placeholders}) ORDER BY priority, received_at",
                pending_states
            )
            return [EmailRecord.from_dict(dict(row)) for row in cursor.fetchall()]

    def get_recent_emails(self, hours: int = 24, limit: int = 100) -> List[EmailRecord]:
        """Get emails from the last N hours."""
        with self._get_connection() as conn:
            cursor = conn.execute("""
                SELECT * FROM emails
                WHERE datetime(received_at) > datetime('now', ? || ' hours')
                ORDER BY received_at DESC
                LIMIT ?
            """, (f"-{hours}", limit))
            return [EmailRecord.from_dict(dict(row)) for row in cursor.fetchall()]

    # Processed message tracking
    def is_message_processed(self, message_id: str, mailbox: str) -> bool:
        """Check if a message has already been processed."""
        with self._get_connection() as conn:
            cursor = conn.execute(
                "SELECT 1 FROM processed_messages WHERE message_id = ? AND mailbox = ?",
                (message_id, mailbox)
            )
            return cursor.fetchone() is not None

    def mark_message_processed(self, message_id: str, mailbox: str) -> None:
        """Mark a message as processed."""
        with self._get_connection() as conn:
            conn.execute(
                "INSERT OR IGNORE INTO processed_messages (message_id, mailbox, processed_at) VALUES (?, ?, ?)",
                (message_id, mailbox, datetime.utcnow().isoformat())
            )

    # Audit log operations
    def log_audit(self, entry: AuditLogEntry) -> None:
        """Add an audit log entry."""
        with self._get_connection() as conn:
            data = entry.to_dict()
            placeholders = ", ".join(["?" for _ in data])
            columns = ", ".join(data.keys())
            conn.execute(
                f"INSERT INTO audit_log ({columns}) VALUES ({placeholders})",
                list(data.values())
            )

    def get_audit_log(
        self,
        email_id: Optional[str] = None,
        limit: int = 100,
        offset: int = 0
    ) -> List[AuditLogEntry]:
        """Get audit log entries."""
        with self._get_connection() as conn:
            if email_id:
                cursor = conn.execute(
                    "SELECT * FROM audit_log WHERE email_id = ? ORDER BY timestamp DESC LIMIT ? OFFSET ?",
                    (email_id, limit, offset)
                )
            else:
                cursor = conn.execute(
                    "SELECT * FROM audit_log ORDER BY timestamp DESC LIMIT ? OFFSET ?",
                    (limit, offset)
                )

            entries = []
            for row in cursor.fetchall():
                entry = AuditLogEntry(
                    id=row["id"],
                    email_id=row["email_id"],
                    timestamp=datetime.fromisoformat(row["timestamp"]),
                    agent=row["agent"],
                    action=row["action"],
                    details=json.loads(row["details"]) if row["details"] else {},
                    user_command=row["user_command"],
                    success=bool(row["success"]),
                    error=row["error"],
                )
                entries.append(entry)
            return entries

    # Spam rules operations
    def save_spam_rule(self, rule: SpamRule) -> None:
        """Save or update a spam rule."""
        with self._get_connection() as conn:
            data = rule.to_dict()
            placeholders = ", ".join(["?" for _ in data])
            columns = ", ".join(data.keys())
            updates = ", ".join([f"{k}=excluded.{k}" for k in data.keys() if k != "id"])

            conn.execute(f"""
                INSERT INTO spam_rules ({columns}) VALUES ({placeholders})
                ON CONFLICT(id) DO UPDATE SET {updates}
            """, list(data.values()))

    def get_active_spam_rules(self) -> List[SpamRule]:
        """Get all active spam rules."""
        with self._get_connection() as conn:
            cursor = conn.execute(
                "SELECT * FROM spam_rules WHERE is_active = 1 ORDER BY confidence DESC"
            )
            rules = []
            for row in cursor.fetchall():
                rule = SpamRule(
                    id=row["id"],
                    rule_type=row["rule_type"],
                    pattern=row["pattern"],
                    action=row["action"],
                    confidence=row["confidence"],
                    hit_count=row["hit_count"],
                    false_positives=row["false_positives"],
                    created_at=datetime.fromisoformat(row["created_at"]),
                    last_hit=datetime.fromisoformat(row["last_hit"]) if row["last_hit"] else None,
                    is_active=bool(row["is_active"]),
                )
                rules.append(rule)
            return rules

    def increment_spam_rule_hit(self, rule_id: str) -> None:
        """Increment hit count for a spam rule."""
        with self._get_connection() as conn:
            conn.execute("""
                UPDATE spam_rules
                SET hit_count = hit_count + 1, last_hit = ?
                WHERE id = ?
            """, (datetime.utcnow().isoformat(), rule_id))

    def record_spam_false_positive(self, rule_id: str) -> None:
        """Record a false positive for a spam rule."""
        with self._get_connection() as conn:
            conn.execute("""
                UPDATE spam_rules
                SET false_positives = false_positives + 1,
                    confidence = MAX(0, confidence - 10)
                WHERE id = ?
            """, (rule_id,))

    # Statistics
    def get_stats(self, hours: int = 24) -> Dict[str, Any]:
        """Get email processing statistics."""
        with self._get_connection() as conn:
            stats = {}

            # Total emails in period
            cursor = conn.execute("""
                SELECT COUNT(*) as total FROM emails
                WHERE datetime(received_at) > datetime('now', ? || ' hours')
            """, (f"-{hours}",))
            stats["total_emails"] = cursor.fetchone()["total"]

            # Emails by state
            cursor = conn.execute("""
                SELECT state, COUNT(*) as count FROM emails
                WHERE datetime(received_at) > datetime('now', ? || ' hours')
                GROUP BY state
            """, (f"-{hours}",))
            stats["by_state"] = {row["state"]: row["count"] for row in cursor.fetchall()}

            # Emails sent
            cursor = conn.execute("""
                SELECT COUNT(*) as sent FROM emails
                WHERE state = 'sent'
                AND datetime(sent_at) > datetime('now', ? || ' hours')
            """, (f"-{hours}",))
            stats["emails_sent"] = cursor.fetchone()["sent"]

            # Spam filtered
            cursor = conn.execute("""
                SELECT COUNT(*) as spam FROM emails
                WHERE state IN ('spam_detected', 'archived')
                AND category = 'spam_candidate'
                AND datetime(received_at) > datetime('now', ? || ' hours')
            """, (f"-{hours}",))
            stats["spam_filtered"] = cursor.fetchone()["spam"]

            return stats
