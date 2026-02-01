-- Email AI Manager - Follow-up Tracking
-- Version: 002

-- Add follow-up columns to emails table
ALTER TABLE emails ADD COLUMN follow_up_at TEXT;
ALTER TABLE emails ADD COLUMN follow_up_note TEXT;
ALTER TABLE emails ADD COLUMN follow_up_reminded_count INTEGER DEFAULT 0;

-- Add thread context and auto-send columns if not present
ALTER TABLE emails ADD COLUMN thread_context TEXT;
ALTER TABLE emails ADD COLUMN auto_send_eligible INTEGER DEFAULT 0;

-- Index for finding pending follow-ups
CREATE INDEX IF NOT EXISTS idx_emails_follow_up ON emails(follow_up_at) WHERE follow_up_at IS NOT NULL;
CREATE INDEX IF NOT EXISTS idx_emails_state_follow_up ON emails(state) WHERE state = 'follow_up';
