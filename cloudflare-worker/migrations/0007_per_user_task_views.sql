-- Materialized per-user task views preserve the exact contents of each
-- employee's legacy OneDrive list without allowing a recipient copy to
-- overwrite the canonical task owner's row.
CREATE TABLE IF NOT EXISTS user_task_views (
  user_email TEXT NOT NULL,
  app_task_id TEXT NOT NULL,
  legacy_payload TEXT NOT NULL DEFAULT '{}',
  present INTEGER NOT NULL DEFAULT 1 CHECK (present IN (0, 1)),
  updated_at TEXT NOT NULL,
  PRIMARY KEY (user_email, app_task_id)
);

CREATE INDEX IF NOT EXISTS idx_user_task_views_present_updated
  ON user_task_views(user_email, present, updated_at);
