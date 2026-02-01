-- Email AI Manager - Phase 4 Schema Updates
-- Version: 002
-- Adds: is_vip, is_auto_sent, response_time_minutes columns

-- Add Phase 4 columns to emails table
ALTER TABLE emails ADD COLUMN is_vip INTEGER DEFAULT 0;
ALTER TABLE emails ADD COLUMN is_auto_sent INTEGER DEFAULT 0;
ALTER TABLE emails ADD COLUMN response_time_minutes INTEGER;
