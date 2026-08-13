-- Preserve the complete legacy UI shape during the controlled D1 read
-- rollout. Rows missing from a completed shadow snapshot are hidden rather
-- than deleted, preserving rollback and auditability.
ALTER TABLE tasks ADD COLUMN legacy_payload TEXT NOT NULL DEFAULT '{}';
ALTER TABLE tasks ADD COLUMN legacy_present INTEGER NOT NULL DEFAULT 1
  CHECK (legacy_present IN (0, 1));

INSERT OR IGNORE INTO feature_flags(name, value, updated_by_email, updated_at)
VALUES ('shared_storage_visible_read_users', '', 'migration', CURRENT_TIMESTAMP);

CREATE INDEX IF NOT EXISTS idx_tasks_owner_source_present
  ON tasks(owner_email, source_type, legacy_present, updated_at);
