PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS recurring_schedules (
  id TEXT PRIMARY KEY,
  title TEXT NOT NULL,
  summary TEXT NOT NULL DEFAULT '',
  department_name TEXT NOT NULL DEFAULT 'Needs Department',
  priority TEXT NOT NULL DEFAULT 'Normal',
  proof_instructions TEXT NOT NULL DEFAULT '',
  assigner_email TEXT NOT NULL,
  assigner_name TEXT NOT NULL DEFAULT '',
  recipient_email TEXT NOT NULL,
  recipient_name TEXT NOT NULL DEFAULT '',
  cadence TEXT NOT NULL DEFAULT 'weekly' CHECK (cadence IN ('weekly')),
  interval_weeks INTEGER NOT NULL DEFAULT 1 CHECK (interval_weeks BETWEEN 1 AND 52),
  generation_lead_days INTEGER NOT NULL DEFAULT 4 CHECK (generation_lead_days BETWEEN 0 AND 30),
  next_due_date TEXT NOT NULL,
  active INTEGER NOT NULL DEFAULT 1 CHECK (active IN (0, 1)),
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_recurring_schedules_due
  ON recurring_schedules(active, next_due_date);
CREATE INDEX IF NOT EXISTS idx_recurring_schedules_assigner
  ON recurring_schedules(assigner_email, updated_at);

CREATE TABLE IF NOT EXISTS recurring_occurrences (
  id TEXT PRIMARY KEY,
  schedule_id TEXT NOT NULL,
  period_key TEXT NOT NULL,
  app_task_id TEXT NOT NULL UNIQUE,
  assignment_id TEXT NOT NULL UNIQUE,
  due_date TEXT NOT NULL,
  created_at TEXT NOT NULL,
  UNIQUE(schedule_id, period_key),
  FOREIGN KEY (schedule_id) REFERENCES recurring_schedules(id)
);

CREATE INDEX IF NOT EXISTS idx_recurring_occurrences_schedule
  ON recurring_occurrences(schedule_id, due_date DESC);
