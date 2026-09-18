ALTER TABLE sessions ADD COLUMN ip_address TEXT;
ALTER TABLE sessions ADD COLUMN user_agent TEXT;
ALTER TABLE sessions ADD COLUMN last_active_at TEXT;
UPDATE sessions SET last_active_at = created_at WHERE last_active_at IS NULL;
