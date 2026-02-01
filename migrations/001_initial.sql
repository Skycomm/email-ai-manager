-- Email AI Manager - Initial Schema
-- Version: 001

-- Emails table - stores all tracked emails
CREATE TABLE IF NOT EXISTS emails (
    id TEXT PRIMARY KEY,
    message_id TEXT NOT NULL,
    mailbox TEXT NOT NULL,
    thread_id TEXT,

    -- Sender info
    sender_email TEXT NOT NULL,
    sender_name TEXT,

    -- Recipients
    to_recipients TEXT,  -- JSON array
    cc_recipients TEXT,  -- JSON array

    -- Content
    subject TEXT NOT NULL,
    body_preview TEXT,
    body_full TEXT,  -- May be encrypted
    received_at TEXT NOT NULL,
    has_attachments INTEGER DEFAULT 0,
    importance TEXT DEFAULT 'normal',

    -- AI Processing
    state TEXT NOT NULL DEFAULT 'new',
    category TEXT,
    priority INTEGER DEFAULT 3,
    spam_score INTEGER DEFAULT 0,
    summary TEXT,

    -- Draft handling
    current_draft TEXT,
    draft_versions TEXT,  -- JSON array
    draft_mode TEXT DEFAULT 'professional',
    approval_token TEXT,

    -- Teams tracking
    teams_message_id TEXT,
    teams_thread_id TEXT,

    -- Audit
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    sent_at TEXT,
    handled_by TEXT DEFAULT 'pending',

    -- Error handling
    error_message TEXT,
    retry_count INTEGER DEFAULT 0,

    -- Phase 4 columns
    is_vip INTEGER DEFAULT 0,
    is_auto_sent INTEGER DEFAULT 0,
    response_time_minutes INTEGER,

    -- Ensure unique message per mailbox
    UNIQUE(message_id, mailbox)
);

-- Audit log table - append-only log of all actions
CREATE TABLE IF NOT EXISTS audit_log (
    id TEXT PRIMARY KEY,
    email_id TEXT,
    timestamp TEXT NOT NULL,
    agent TEXT NOT NULL,
    action TEXT NOT NULL,
    details TEXT,  -- JSON object
    user_command TEXT,
    success INTEGER DEFAULT 1,
    error TEXT,
    FOREIGN KEY (email_id) REFERENCES emails(id)
);

-- Spam rules table - learned spam patterns
CREATE TABLE IF NOT EXISTS spam_rules (
    id TEXT PRIMARY KEY,
    rule_type TEXT NOT NULL,  -- sender, domain, subject_keyword, pattern
    pattern TEXT NOT NULL,
    action TEXT DEFAULT 'archive',  -- archive, delete, digest
    confidence INTEGER DEFAULT 50,  -- 0-100
    hit_count INTEGER DEFAULT 0,
    false_positives INTEGER DEFAULT 0,
    created_at TEXT NOT NULL,
    last_hit TEXT,
    is_active INTEGER DEFAULT 1
);

-- Processed messages tracking - prevent duplicate processing
CREATE TABLE IF NOT EXISTS processed_messages (
    message_id TEXT NOT NULL,
    mailbox TEXT NOT NULL,
    processed_at TEXT NOT NULL,
    PRIMARY KEY (message_id, mailbox)
);

-- VIP senders (Phase 2)
CREATE TABLE IF NOT EXISTS vip_senders (
    id TEXT PRIMARY KEY,
    email_pattern TEXT NOT NULL,  -- exact email or domain pattern
    name TEXT,
    priority_boost INTEGER DEFAULT 2,  -- How much to increase priority
    auto_category TEXT,  -- Override category for this sender
    created_at TEXT NOT NULL,
    is_active INTEGER DEFAULT 1
);

-- Daily digest tracking
CREATE TABLE IF NOT EXISTS digest_entries (
    id TEXT PRIMARY KEY,
    email_id TEXT NOT NULL,
    digest_date TEXT NOT NULL,  -- YYYY-MM-DD
    included INTEGER DEFAULT 0,
    FOREIGN KEY (email_id) REFERENCES emails(id)
);

-- Muted senders (senders that won't trigger Teams notifications)
CREATE TABLE IF NOT EXISTS muted_senders (
    id TEXT PRIMARY KEY,
    email_pattern TEXT UNIQUE NOT NULL,  -- email address or domain
    muted_at TEXT NOT NULL,
    reason TEXT
);

-- Application settings
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT,
    updated_at TEXT NOT NULL
);

-- Indexes for common queries
CREATE INDEX IF NOT EXISTS idx_emails_state ON emails(state);
CREATE INDEX IF NOT EXISTS idx_emails_mailbox ON emails(mailbox);
CREATE INDEX IF NOT EXISTS idx_emails_received ON emails(received_at);
CREATE INDEX IF NOT EXISTS idx_emails_approval_token ON emails(approval_token);
CREATE INDEX IF NOT EXISTS idx_emails_category ON emails(category);
CREATE INDEX IF NOT EXISTS idx_emails_priority ON emails(priority);

CREATE INDEX IF NOT EXISTS idx_audit_email_id ON audit_log(email_id);
CREATE INDEX IF NOT EXISTS idx_audit_timestamp ON audit_log(timestamp);
CREATE INDEX IF NOT EXISTS idx_audit_agent ON audit_log(agent);

CREATE INDEX IF NOT EXISTS idx_spam_rules_type ON spam_rules(rule_type);
CREATE INDEX IF NOT EXISTS idx_spam_rules_active ON spam_rules(is_active);

CREATE INDEX IF NOT EXISTS idx_digest_date ON digest_entries(digest_date);
