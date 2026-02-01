"""
SQLite database operations.
"""

import json
import logging
import sqlite3
import uuid
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from .models import AuditLogEntry, EmailRecord, EmailState, SpamRule, EmailRule, RuleAction

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
            else:
                # Apply Phase 4 migrations to existing database
                self._apply_migrations(conn)

    def _apply_migrations(self, conn: sqlite3.Connection):
        """Apply migrations for Phase 4 columns to existing database."""
        # Check if is_vip column exists
        cursor = conn.execute("PRAGMA table_info(emails)")
        columns = [row[1] for row in cursor.fetchall()]

        if 'is_vip' not in columns:
            logger.info("Applying Phase 4 migration: adding is_vip column")
            conn.execute("ALTER TABLE emails ADD COLUMN is_vip INTEGER DEFAULT 0")

        if 'is_auto_sent' not in columns:
            logger.info("Applying Phase 4 migration: adding is_auto_sent column")
            conn.execute("ALTER TABLE emails ADD COLUMN is_auto_sent INTEGER DEFAULT 0")

        if 'response_time_minutes' not in columns:
            logger.info("Applying Phase 4 migration: adding response_time_minutes column")
            conn.execute("ALTER TABLE emails ADD COLUMN response_time_minutes INTEGER")

        if 'thread_context' not in columns:
            logger.info("Applying Phase 4 migration: adding thread_context column")
            conn.execute("ALTER TABLE emails ADD COLUMN thread_context TEXT")

        if 'auto_send_eligible' not in columns:
            logger.info("Applying Phase 4 migration: adding auto_send_eligible column")
            conn.execute("ALTER TABLE emails ADD COLUMN auto_send_eligible INTEGER DEFAULT 0")

        # Check if muted_senders table exists
        cursor = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='muted_senders'"
        )
        if cursor.fetchone() is None:
            logger.info("Creating muted_senders table")
            conn.execute("""
                CREATE TABLE IF NOT EXISTS muted_senders (
                    id TEXT PRIMARY KEY,
                    email_pattern TEXT UNIQUE NOT NULL,
                    muted_at TEXT NOT NULL,
                    reason TEXT
                )
            """)

        # Check if settings table exists
        cursor = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='settings'"
        )
        if cursor.fetchone() is None:
            logger.info("Creating settings table")
            conn.execute("""
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT,
                    updated_at TEXT NOT NULL
                )
            """)

        # Follow-up tracking columns
        if 'follow_up_at' not in columns:
            logger.info("Applying migration: adding follow_up_at column")
            conn.execute("ALTER TABLE emails ADD COLUMN follow_up_at TEXT")

        if 'follow_up_note' not in columns:
            logger.info("Applying migration: adding follow_up_note column")
            conn.execute("ALTER TABLE emails ADD COLUMN follow_up_note TEXT")

        if 'follow_up_reminded_count' not in columns:
            logger.info("Applying migration: adding follow_up_reminded_count column")
            conn.execute("ALTER TABLE emails ADD COLUMN follow_up_reminded_count INTEGER DEFAULT 0")

        # Check if email_rules table exists
        cursor = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='email_rules'"
        )
        if cursor.fetchone() is None:
            logger.info("Creating email_rules table for LLM-based routing")
            conn.execute("""
                CREATE TABLE IF NOT EXISTS email_rules (
                    id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    description TEXT,
                    match_prompt TEXT NOT NULL,
                    action TEXT NOT NULL,
                    action_value TEXT,
                    priority INTEGER DEFAULT 50,
                    is_active INTEGER DEFAULT 1,
                    stop_processing INTEGER DEFAULT 1,
                    hit_count INTEGER DEFAULT 0,
                    last_hit TEXT,
                    false_positives INTEGER DEFAULT 0,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                )
            """)

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
                is_vip INTEGER DEFAULT 0,
                thread_context TEXT,
                auto_send_eligible INTEGER DEFAULT 0,
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

            -- Muted senders - never show these in Teams notifications
            CREATE TABLE IF NOT EXISTS muted_senders (
                id TEXT PRIMARY KEY,
                email_pattern TEXT NOT NULL UNIQUE,  -- email or domain pattern
                muted_at TEXT NOT NULL,
                reason TEXT  -- optional reason
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

    def delete_email(self, email_id: str) -> bool:
        """Delete an email record from the database."""
        with self._get_connection() as conn:
            cursor = conn.execute(
                "DELETE FROM emails WHERE id = ?",
                (email_id,)
            )
            return cursor.rowcount > 0

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

    def get_pending_followups(self) -> List[EmailRecord]:
        """Get emails that need follow-up reminders (due or overdue)."""
        with self._get_connection() as conn:
            cursor = conn.execute("""
                SELECT * FROM emails
                WHERE state = 'follow_up'
                AND follow_up_at IS NOT NULL
                AND datetime(follow_up_at) <= datetime('now')
                ORDER BY follow_up_at ASC
            """)
            return [EmailRecord.from_dict(dict(row)) for row in cursor.fetchall()]

    def get_all_followups(self) -> List[EmailRecord]:
        """Get all emails marked for follow-up (for dashboard)."""
        with self._get_connection() as conn:
            cursor = conn.execute("""
                SELECT * FROM emails
                WHERE state = 'follow_up'
                ORDER BY follow_up_at ASC NULLS LAST
            """)
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

    def get_audit_log_count(self, email_id: Optional[str] = None) -> int:
        """Get total count of audit log entries."""
        with self._get_connection() as conn:
            if email_id:
                cursor = conn.execute(
                    "SELECT COUNT(*) as count FROM audit_log WHERE email_id = ?",
                    (email_id,)
                )
            else:
                cursor = conn.execute("SELECT COUNT(*) as count FROM audit_log")
            return cursor.fetchone()["count"]

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

    def delete_spam_rule(self, rule_id: str) -> None:
        """Delete a spam rule."""
        with self._get_connection() as conn:
            conn.execute("DELETE FROM spam_rules WHERE id = ?", (rule_id,))

    # Email rules operations (LLM-based routing)
    def save_email_rule(self, rule: EmailRule) -> None:
        """Save or update an email rule."""
        with self._get_connection() as conn:
            data = rule.to_dict()
            placeholders = ", ".join(["?" for _ in data])
            columns = ", ".join(data.keys())
            updates = ", ".join([f"{k}=excluded.{k}" for k in data.keys() if k != "id"])

            conn.execute(f"""
                INSERT INTO email_rules ({columns}) VALUES ({placeholders})
                ON CONFLICT(id) DO UPDATE SET {updates}
            """, list(data.values()))

    def get_email_rule(self, rule_id: str) -> Optional[EmailRule]:
        """Get a single email rule by ID."""
        with self._get_connection() as conn:
            cursor = conn.execute(
                "SELECT * FROM email_rules WHERE id = ?",
                (rule_id,)
            )
            row = cursor.fetchone()
            if row:
                return EmailRule.from_dict(dict(row))
        return None

    def get_active_email_rules(self) -> List[EmailRule]:
        """Get all active email rules, ordered by priority."""
        with self._get_connection() as conn:
            cursor = conn.execute(
                "SELECT * FROM email_rules WHERE is_active = 1 ORDER BY priority ASC"
            )
            return [EmailRule.from_dict(dict(row)) for row in cursor.fetchall()]

    def get_all_email_rules(self) -> List[EmailRule]:
        """Get all email rules (including inactive), ordered by priority."""
        with self._get_connection() as conn:
            cursor = conn.execute(
                "SELECT * FROM email_rules ORDER BY priority ASC"
            )
            return [EmailRule.from_dict(dict(row)) for row in cursor.fetchall()]

    def delete_email_rule(self, rule_id: str) -> bool:
        """Delete an email rule."""
        with self._get_connection() as conn:
            cursor = conn.execute("DELETE FROM email_rules WHERE id = ?", (rule_id,))
            return cursor.rowcount > 0

    def increment_email_rule_hit(self, rule_id: str) -> None:
        """Increment hit count for an email rule."""
        with self._get_connection() as conn:
            conn.execute("""
                UPDATE email_rules
                SET hit_count = hit_count + 1, last_hit = ?, updated_at = ?
                WHERE id = ?
            """, (datetime.utcnow().isoformat(), datetime.utcnow().isoformat(), rule_id))

    def record_email_rule_false_positive(self, rule_id: str) -> None:
        """Record a false positive for an email rule."""
        with self._get_connection() as conn:
            conn.execute("""
                UPDATE email_rules
                SET false_positives = false_positives + 1, updated_at = ?
                WHERE id = ?
            """, (datetime.utcnow().isoformat(), rule_id))

    # Statistics
    def get_stats(self, hours: int = 24) -> Dict[str, Any]:
        """Get email processing statistics."""
        with self._get_connection() as conn:
            stats = {}

            # Total emails in period (excluding deleted/ignored and spam)
            cursor = conn.execute("""
                SELECT COUNT(*) as total FROM emails
                WHERE datetime(received_at) > datetime('now', ? || ' hours')
                AND state NOT IN ('ignored', 'spam_detected')
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

    def get_advanced_stats(self, hours: int = 24) -> Dict[str, Any]:
        """Get detailed analytics for the dashboard."""
        with self._get_connection() as conn:
            stats = {}

            # Basic stats
            basic = self.get_stats(hours)
            stats.update(basic)

            # Emails by category
            cursor = conn.execute("""
                SELECT category, COUNT(*) as count FROM emails
                WHERE datetime(received_at) > datetime('now', ? || ' hours')
                AND category IS NOT NULL
                GROUP BY category
            """, (f"-{hours}",))
            stats["by_category"] = {row["category"]: row["count"] for row in cursor.fetchall()}

            # Emails by mailbox
            cursor = conn.execute("""
                SELECT mailbox, COUNT(*) as count FROM emails
                WHERE datetime(received_at) > datetime('now', ? || ' hours')
                GROUP BY mailbox
            """, (f"-{hours}",))
            stats["by_mailbox"] = {row["mailbox"]: row["count"] for row in cursor.fetchall()}

            # Auto-sent emails
            cursor = conn.execute("""
                SELECT COUNT(*) as auto_sent FROM emails
                WHERE handled_by = 'ai_auto'
                AND datetime(sent_at) > datetime('now', ? || ' hours')
            """, (f"-{hours}",))
            stats["auto_sent"] = cursor.fetchone()["auto_sent"]

            # VIP emails
            cursor = conn.execute("""
                SELECT COUNT(*) as vip_count FROM emails
                WHERE is_vip = 1
                AND datetime(received_at) > datetime('now', ? || ' hours')
            """, (f"-{hours}",))
            stats["vip_emails"] = cursor.fetchone()["vip_count"]

            # Average response time (for sent emails)
            cursor = conn.execute("""
                SELECT AVG(
                    (julianday(sent_at) - julianday(received_at)) * 24 * 60
                ) as avg_response_minutes
                FROM emails
                WHERE state = 'sent'
                AND sent_at IS NOT NULL
                AND datetime(sent_at) > datetime('now', ? || ' hours')
            """, (f"-{hours}",))
            result = cursor.fetchone()
            stats["avg_response_minutes"] = round(result["avg_response_minutes"] or 0, 1)

            # Top senders
            cursor = conn.execute("""
                SELECT sender_email, sender_name, COUNT(*) as count
                FROM emails
                WHERE datetime(received_at) > datetime('now', ? || ' hours')
                GROUP BY sender_email
                ORDER BY count DESC
                LIMIT 10
            """, (f"-{hours}",))
            stats["top_senders"] = [
                {"email": row["sender_email"], "name": row["sender_name"], "count": row["count"]}
                for row in cursor.fetchall()
            ]

            # Hourly distribution (last 24 hours)
            cursor = conn.execute("""
                SELECT strftime('%H', received_at) as hour, COUNT(*) as count
                FROM emails
                WHERE datetime(received_at) > datetime('now', '-24 hours')
                GROUP BY hour
                ORDER BY hour
            """)
            stats["hourly_distribution"] = {row["hour"]: row["count"] for row in cursor.fetchall()}

            # Priority distribution
            cursor = conn.execute("""
                SELECT priority, COUNT(*) as count
                FROM emails
                WHERE datetime(received_at) > datetime('now', ? || ' hours')
                GROUP BY priority
                ORDER BY priority
            """, (f"-{hours}",))
            stats["by_priority"] = {str(row["priority"]): row["count"] for row in cursor.fetchall()}

            # Meeting emails
            cursor = conn.execute("""
                SELECT COUNT(*) as meeting_count FROM emails
                WHERE category = 'meeting'
                AND datetime(received_at) > datetime('now', ? || ' hours')
            """, (f"-{hours}",))
            stats["meeting_emails"] = cursor.fetchone()["meeting_count"]

            return stats

    def get_sender_stats(self, sender_email: str, hours: int = 168) -> Dict[str, Any]:
        """Get stats for a specific sender."""
        with self._get_connection() as conn:
            stats = {}

            # Total emails from sender
            cursor = conn.execute("""
                SELECT COUNT(*) as total FROM emails
                WHERE sender_email = ?
                AND datetime(received_at) > datetime('now', ? || ' hours')
            """, (sender_email, f"-{hours}"))
            stats["total_emails"] = cursor.fetchone()["total"]

            # Categories
            cursor = conn.execute("""
                SELECT category, COUNT(*) as count FROM emails
                WHERE sender_email = ?
                AND datetime(received_at) > datetime('now', ? || ' hours')
                GROUP BY category
            """, (sender_email, f"-{hours}"))
            stats["by_category"] = {row["category"]: row["count"] for row in cursor.fetchall()}

            # Spam rate
            cursor = conn.execute("""
                SELECT COUNT(*) as spam FROM emails
                WHERE sender_email = ?
                AND category = 'spam_candidate'
                AND datetime(received_at) > datetime('now', ? || ' hours')
            """, (sender_email, f"-{hours}"))
            spam_count = cursor.fetchone()["spam"]
            stats["spam_rate"] = round(spam_count / max(1, stats["total_emails"]) * 100, 1)

            return stats

    # Muted senders operations
    def mute_sender(self, email_pattern: str, reason: Optional[str] = None) -> None:
        """Mute a sender (never show in Teams notifications)."""
        with self._get_connection() as conn:
            conn.execute("""
                INSERT OR REPLACE INTO muted_senders (id, email_pattern, muted_at, reason)
                VALUES (?, ?, ?, ?)
            """, (str(uuid.uuid4()), email_pattern.lower(), datetime.utcnow().isoformat(), reason))

    def unmute_sender(self, email_pattern: str) -> bool:
        """Unmute a sender."""
        with self._get_connection() as conn:
            cursor = conn.execute(
                "DELETE FROM muted_senders WHERE email_pattern = ?",
                (email_pattern.lower(),)
            )
            return cursor.rowcount > 0

    def is_sender_muted(self, sender_email: str) -> bool:
        """Check if a sender is muted (exact match or domain match)."""
        sender_lower = sender_email.lower()
        sender_domain = sender_lower.split('@')[-1] if '@' in sender_lower else ""

        with self._get_connection() as conn:
            # Check exact email match
            cursor = conn.execute(
                "SELECT 1 FROM muted_senders WHERE email_pattern = ?",
                (sender_lower,)
            )
            if cursor.fetchone():
                return True

            # Check domain match
            if sender_domain:
                cursor = conn.execute(
                    "SELECT 1 FROM muted_senders WHERE email_pattern = ?",
                    (sender_domain,)
                )
                if cursor.fetchone():
                    return True

            return False

    def get_muted_senders(self) -> List[Dict[str, Any]]:
        """Get all muted senders."""
        with self._get_connection() as conn:
            cursor = conn.execute(
                "SELECT email_pattern as pattern, muted_at, reason FROM muted_senders ORDER BY muted_at DESC"
            )
            return [dict(row) for row in cursor.fetchall()]

    # Settings operations (key-value store for app state)
    def get_setting(self, key: str) -> Optional[str]:
        """Get a setting value."""
        with self._get_connection() as conn:
            # Ensure settings table exists
            conn.execute("""
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT,
                    updated_at TEXT
                )
            """)
            cursor = conn.execute(
                "SELECT value FROM settings WHERE key = ?",
                (key,)
            )
            row = cursor.fetchone()
            return row["value"] if row else None

    def set_setting(self, key: str, value: str) -> None:
        """Set a setting value."""
        with self._get_connection() as conn:
            # Ensure settings table exists
            conn.execute("""
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT,
                    updated_at TEXT
                )
            """)
            conn.execute("""
                INSERT OR REPLACE INTO settings (key, value, updated_at)
                VALUES (?, ?, ?)
            """, (key, value, datetime.utcnow().isoformat()))

    def save_summary_mapping(self, mapping: Dict[int, str]) -> None:
        """Save the summary email mapping (number -> email_id) to persist across restarts."""
        import json
        # Convert int keys to str for JSON
        str_mapping = {str(k): v for k, v in mapping.items()}
        self.set_setting("summary_email_mapping", json.dumps(str_mapping))

    def get_summary_mapping(self) -> Dict[int, str]:
        """Get the summary email mapping from database."""
        import json
        value = self.get_setting("summary_email_mapping")
        if not value:
            return {}
        try:
            str_mapping = json.loads(value)
            # Convert str keys back to int
            return {int(k): v for k, v in str_mapping.items()}
        except (json.JSONDecodeError, ValueError):
            return {}

    def clear_summary_mapping(self) -> None:
        """Clear the summary email mapping."""
        self.set_setting("summary_email_mapping", "{}")

    # Email queries by category
    def get_emails_by_category(
        self,
        category,  # EmailCategory
        states: Optional[List] = None,  # List[EmailState]
        limit: int = 100
    ) -> List[EmailRecord]:
        """Get emails by category, optionally filtered by states."""
        with self._get_connection() as conn:
            if states:
                state_values = [s.value for s in states]
                placeholders = ", ".join(["?" for _ in state_values])
                cursor = conn.execute(
                    f"""SELECT * FROM emails
                    WHERE category = ? AND state IN ({placeholders})
                    ORDER BY received_at DESC LIMIT ?""",
                    [category.value] + state_values + [limit]
                )
            else:
                cursor = conn.execute(
                    "SELECT * FROM emails WHERE category = ? ORDER BY received_at DESC LIMIT ?",
                    (category.value, limit)
                )
            return [EmailRecord.from_dict(dict(row)) for row in cursor.fetchall()]

    def get_fyi_emails_last_24h(self, limit: int = 50) -> List[EmailRecord]:
        """
        Get all FYI/newsletter emails from the last 24 hours for morning summary.
        Excludes emails that have been archived, deleted, or ignored by the user.
        """
        with self._get_connection() as conn:
            cursor = conn.execute("""
                SELECT * FROM emails
                WHERE category IN ('fyi', 'newsletter')
                AND state NOT IN ('archived', 'ignored', 'sent', 'error', 'spam_detected')
                AND datetime(received_at) > datetime('now', '-24 hours')
                ORDER BY received_at DESC
                LIMIT ?
            """, (limit,))
            return [EmailRecord.from_dict(dict(row)) for row in cursor.fetchall()]

    def get_old_fyi_emails_to_archive(self, older_than_hours: int = 48, limit: int = 100) -> List[EmailRecord]:
        """
        Get FYI/newsletter emails older than specified hours that should be auto-archived.
        Only returns emails in 'fyi_notified' or 'acknowledged' state (already processed).
        Does not include emails that need action or were already archived.
        """
        with self._get_connection() as conn:
            cursor = conn.execute("""
                SELECT * FROM emails
                WHERE category IN ('fyi', 'newsletter')
                AND state IN ('fyi_notified', 'acknowledged', 'new')
                AND datetime(received_at) < datetime('now', ? || ' hours')
                ORDER BY received_at ASC
                LIMIT ?
            """, (f"-{older_than_hours}", limit))
            return [EmailRecord.from_dict(dict(row)) for row in cursor.fetchall()]

    def archive_old_fyi_emails(self, older_than_hours: int = 48) -> int:
        """
        Auto-archive old FYI/newsletter emails.
        Returns the number of emails archived.
        """
        with self._get_connection() as conn:
            cursor = conn.execute("""
                UPDATE emails
                SET state = 'archived', updated_at = datetime('now')
                WHERE category IN ('fyi', 'newsletter')
                AND state IN ('fyi_notified', 'acknowledged', 'new')
                AND datetime(received_at) < datetime('now', ? || ' hours')
            """, (f"-{older_than_hours}",))
            return cursor.rowcount

    def get_auto_sent_emails_last_24h(self, limit: int = 20) -> List[EmailRecord]:
        """Get emails that were auto-sent in the last 24 hours."""
        with self._get_connection() as conn:
            cursor = conn.execute("""
                SELECT * FROM emails
                WHERE handled_by = 'ai_auto'
                AND state = 'sent'
                AND datetime(sent_at) > datetime('now', '-24 hours')
                ORDER BY sent_at DESC
                LIMIT ?
            """, (limit,))
            return [EmailRecord.from_dict(dict(row)) for row in cursor.fetchall()]
