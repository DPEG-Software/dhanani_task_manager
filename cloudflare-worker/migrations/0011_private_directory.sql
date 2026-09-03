-- Private authenticated directory and task-visibility policy.
-- Actual employee identities and access rules are deliberately seeded out of
-- band so they never become part of the public application repository.
CREATE TABLE IF NOT EXISTS directory_users (
  email TEXT PRIMARY KEY COLLATE NOCASE,
  display_name TEXT NOT NULL DEFAULT '',
  role_title TEXT NOT NULL DEFAULT '',
  department_key TEXT NOT NULL DEFAULT '',
  is_admin INTEGER NOT NULL DEFAULT 0 CHECK (is_admin IN (0, 1)),
  is_principal INTEGER NOT NULL DEFAULT 0 CHECK (is_principal IN (0, 1)),
  wednesday_review INTEGER NOT NULL DEFAULT 0 CHECK (wednesday_review IN (0, 1)),
  active INTEGER NOT NULL DEFAULT 1 CHECK (active IN (0, 1)),
  updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS directory_access_rules (
  id TEXT PRIMARY KEY,
  viewer_email TEXT NOT NULL COLLATE NOCASE,
  rule_type TEXT NOT NULL CHECK (rule_type IN ('department', 'assigner', 'recipient')),
  group_key TEXT NOT NULL DEFAULT '',
  target_value TEXT NOT NULL COLLATE NOCASE,
  updated_at TEXT NOT NULL,
  UNIQUE(viewer_email, rule_type, group_key, target_value)
);

CREATE INDEX IF NOT EXISTS idx_directory_access_viewer
  ON directory_access_rules(viewer_email, rule_type, group_key);
