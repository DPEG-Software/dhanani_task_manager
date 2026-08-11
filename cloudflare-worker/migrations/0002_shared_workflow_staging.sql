PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS tasks (
  id TEXT PRIMARY KEY,
  owner_email TEXT NOT NULL,
  title TEXT NOT NULL,
  summary TEXT NOT NULL DEFAULT '',
  department_name TEXT NOT NULL DEFAULT 'Needs Department',
  priority TEXT NOT NULL DEFAULT 'Normal',
  due_date TEXT,
  status TEXT NOT NULL DEFAULT 'Pending',
  source_type TEXT NOT NULL DEFAULT 'manual',
  source_message_id TEXT,
  source_conversation_id TEXT,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL,
  completed_at TEXT,
  cancelled_at TEXT,
  version INTEGER NOT NULL DEFAULT 1
);

CREATE INDEX IF NOT EXISTS idx_tasks_owner_updated
  ON tasks(owner_email, updated_at);
CREATE INDEX IF NOT EXISTS idx_tasks_department_status
  ON tasks(department_name, status);
CREATE INDEX IF NOT EXISTS idx_tasks_due_status
  ON tasks(due_date, status);

CREATE TABLE IF NOT EXISTS proof_submissions (
  id TEXT PRIMARY KEY,
  assignment_id TEXT NOT NULL,
  app_task_id TEXT NOT NULL,
  submitter_email TEXT NOT NULL,
  submitter_name TEXT NOT NULL DEFAULT '',
  note TEXT NOT NULL DEFAULT '',
  status TEXT NOT NULL DEFAULT 'pending'
    CHECK (status IN ('pending', 'approved', 'changes_requested')),
  submitted_at TEXT NOT NULL,
  reviewed_at TEXT,
  reviewer_email TEXT,
  review_reason TEXT NOT NULL DEFAULT '',
  idempotency_key TEXT UNIQUE,
  legacy_notification_id TEXT UNIQUE,
  FOREIGN KEY (assignment_id) REFERENCES assignments(id)
);

CREATE INDEX IF NOT EXISTS idx_proofs_assignment_submitted
  ON proof_submissions(assignment_id, submitted_at);
CREATE INDEX IF NOT EXISTS idx_proofs_task_submitter
  ON proof_submissions(app_task_id, submitter_email, submitted_at);
CREATE INDEX IF NOT EXISTS idx_proofs_status_submitted
  ON proof_submissions(status, submitted_at);

CREATE TABLE IF NOT EXISTS proof_files (
  id TEXT PRIMARY KEY,
  submission_id TEXT NOT NULL,
  drive_provider TEXT NOT NULL DEFAULT 'onedrive',
  drive_item_id TEXT,
  file_name TEXT NOT NULL,
  mime_type TEXT,
  size_bytes INTEGER,
  web_url TEXT NOT NULL,
  created_at TEXT NOT NULL,
  FOREIGN KEY (submission_id) REFERENCES proof_submissions(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_proof_files_submission
  ON proof_files(submission_id);

CREATE TABLE IF NOT EXISTS reminders (
  id TEXT PRIMARY KEY,
  assignment_id TEXT NOT NULL,
  sender_email TEXT NOT NULL,
  recipient_email TEXT NOT NULL,
  created_at TEXT NOT NULL,
  idempotency_key TEXT UNIQUE,
  FOREIGN KEY (assignment_id) REFERENCES assignments(id)
);

CREATE INDEX IF NOT EXISTS idx_reminders_assignment_created
  ON reminders(assignment_id, created_at);
CREATE INDEX IF NOT EXISTS idx_reminders_recipient_created
  ON reminders(recipient_email, created_at);

CREATE TABLE IF NOT EXISTS departments (
  id TEXT PRIMARY KEY,
  name TEXT NOT NULL COLLATE NOCASE UNIQUE,
  active INTEGER NOT NULL DEFAULT 1 CHECK (active IN (0, 1)),
  created_by_email TEXT NOT NULL,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS department_members (
  user_email TEXT PRIMARY KEY,
  user_name TEXT NOT NULL DEFAULT '',
  department_id TEXT NOT NULL,
  updated_by_email TEXT NOT NULL,
  updated_at TEXT NOT NULL,
  version INTEGER NOT NULL DEFAULT 1,
  FOREIGN KEY (department_id) REFERENCES departments(id)
);

CREATE INDEX IF NOT EXISTS idx_department_members_department
  ON department_members(department_id, user_email);

CREATE TABLE IF NOT EXISTS user_task_state (
  assignment_id TEXT NOT NULL,
  user_email TEXT NOT NULL,
  messages_seen_count INTEGER NOT NULL DEFAULT 0,
  reminders_seen_count INTEGER NOT NULL DEFAULT 0,
  last_opened_at TEXT,
  updated_at TEXT NOT NULL,
  PRIMARY KEY (assignment_id, user_email),
  FOREIGN KEY (assignment_id) REFERENCES assignments(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS discussion_notes (
  id TEXT PRIMARY KEY,
  owner_email TEXT NOT NULL,
  title TEXT NOT NULL DEFAULT '',
  body TEXT NOT NULL DEFAULT '',
  note_date TEXT,
  archived_at TEXT,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL,
  version INTEGER NOT NULL DEFAULT 1
);

CREATE INDEX IF NOT EXISTS idx_discussion_notes_owner_date
  ON discussion_notes(owner_email, note_date);

CREATE TABLE IF NOT EXISTS task_events (
  id TEXT PRIMARY KEY,
  assignment_id TEXT,
  app_task_id TEXT,
  actor_email TEXT NOT NULL,
  event_type TEXT NOT NULL,
  event_data TEXT NOT NULL DEFAULT '{}',
  created_at TEXT NOT NULL,
  idempotency_key TEXT UNIQUE
);

CREATE INDEX IF NOT EXISTS idx_task_events_assignment_created
  ON task_events(assignment_id, created_at);
CREATE INDEX IF NOT EXISTS idx_task_events_task_created
  ON task_events(app_task_id, created_at);
CREATE INDEX IF NOT EXISTS idx_task_events_actor_created
  ON task_events(actor_email, created_at);

CREATE TABLE IF NOT EXISTS migration_records (
  source_store TEXT NOT NULL,
  source_id TEXT NOT NULL,
  target_table TEXT NOT NULL,
  target_id TEXT NOT NULL,
  source_hash TEXT,
  migrated_at TEXT NOT NULL,
  PRIMARY KEY (source_store, source_id)
);

CREATE TABLE IF NOT EXISTS feature_flags (
  name TEXT PRIMARY KEY,
  value TEXT NOT NULL,
  updated_by_email TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

INSERT OR IGNORE INTO feature_flags(name, value, updated_by_email, updated_at)
VALUES ('shared_storage_read_mode', 'legacy', 'migration', CURRENT_TIMESTAMP);

INSERT OR IGNORE INTO feature_flags(name, value, updated_by_email, updated_at)
VALUES ('shared_storage_dual_write', 'off', 'migration', CURRENT_TIMESTAMP);
