-- Staging-only baseline for the DPEG Task Manager.
-- This reproduces the D1 tables already used by production. It is safe for a
-- brand-new staging database and does not migrate or delete production data.

PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS assignments (
  id TEXT PRIMARY KEY,
  app_task_id TEXT NOT NULL,
  title TEXT NOT NULL,
  summary TEXT NOT NULL DEFAULT '',
  dept TEXT NOT NULL DEFAULT 'Needs Department',
  priority TEXT NOT NULL DEFAULT 'Normal',
  due_date TEXT,
  assigner_email TEXT NOT NULL,
  assigner_name TEXT NOT NULL DEFAULT '',
  recipient_email TEXT NOT NULL,
  recipient_name TEXT NOT NULL DEFAULT '',
  status TEXT NOT NULL DEFAULT 'Assigned',
  progress_note TEXT,
  proof_status TEXT NOT NULL DEFAULT 'none',
  proof_submitted_at TEXT,
  proof_reviewed_at TEXT,
  proof_notification_id TEXT,
  recipient_todo_list_id TEXT,
  recipient_todo_task_id TEXT,
  proof_instructions TEXT NOT NULL DEFAULT '',
  update_alert_at TEXT,
  reminder_count INTEGER NOT NULL DEFAULT 0,
  assigner_message_seen_count INTEGER NOT NULL DEFAULT 0,
  recipient_message_seen_count INTEGER NOT NULL DEFAULT 0,
  recipient_reminder_seen_count INTEGER,
  version INTEGER NOT NULL DEFAULT 1,
  cancel_reason TEXT,
  cancelled_at TEXT,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_assignments_recipient_created
  ON assignments(recipient_email, created_at);
CREATE INDEX IF NOT EXISTS idx_assignments_assigner_created
  ON assignments(assigner_email, created_at);
CREATE INDEX IF NOT EXISTS idx_assignments_task_recipient
  ON assignments(app_task_id, recipient_email);

CREATE TABLE IF NOT EXISTS task_messages (
  id TEXT PRIMARY KEY,
  app_task_id TEXT NOT NULL,
  task_title TEXT NOT NULL DEFAULT '',
  assigner_email TEXT NOT NULL,
  recipient_email TEXT NOT NULL,
  recipient_name TEXT NOT NULL DEFAULT '',
  sender_email TEXT NOT NULL,
  sender_name TEXT NOT NULL DEFAULT '',
  sender_role TEXT NOT NULL CHECK (sender_role IN ('assignor', 'assignee')),
  message TEXT NOT NULL,
  created_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_task_messages_thread
  ON task_messages(app_task_id, recipient_email, created_at);
CREATE INDEX IF NOT EXISTS idx_task_messages_assigner
  ON task_messages(assigner_email, created_at);
CREATE INDEX IF NOT EXISTS idx_task_messages_recipient
  ON task_messages(recipient_email, created_at);
