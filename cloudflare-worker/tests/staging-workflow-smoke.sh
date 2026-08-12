#!/bin/sh
set -eu

repo_dir=$(CDPATH= cd -- "$(dirname -- "$0")/../.." && pwd)
test_db=$(mktemp /tmp/dpeg-staging-workflow.XXXXXX.db)
trap 'rm -f "$test_db"' EXIT

sqlite3 "$test_db" < "$repo_dir/cloudflare-worker/migrations/0001_staging_baseline.sql"
sqlite3 "$test_db" < "$repo_dir/cloudflare-worker/migrations/0002_shared_workflow_staging.sql"

sqlite3 "$test_db" <<'SQL'
INSERT INTO tasks
  (id, owner_email, title, summary, department_name, priority, status,
   source_type, created_at, updated_at, version)
VALUES
  ('stg-test-task', 'assigner@dhananipeg.com', 'Synthetic task', '', 'IT',
   'Normal', 'Pending', 'staging_test', '2026-08-11T00:00:00Z',
   '2026-08-11T00:00:00Z', 1);

INSERT INTO assignments
  (id, app_task_id, title, summary, dept, priority, assigner_email,
   assigner_name, recipient_email, recipient_name, status, proof_status,
   created_at, updated_at, version)
VALUES
  ('stg-test-assignment', 'stg-test-task', 'Synthetic task', '', 'IT',
   'Normal', 'assigner@dhananipeg.com', 'Assigner',
   'recipient@dhananipeg.com', 'Recipient', 'Assigned', 'none',
   '2026-08-11T00:00:00Z', '2026-08-11T00:00:00Z', 1);

INSERT INTO reminders
  (id, assignment_id, sender_email, recipient_email, created_at)
VALUES
  ('stg-test-reminder', 'stg-test-assignment', 'assigner@dhananipeg.com',
   'recipient@dhananipeg.com', '2026-08-11T00:00:01Z');

INSERT INTO proof_submissions
  (id, assignment_id, app_task_id, submitter_email, submitter_name, note,
   status, submitted_at)
VALUES
  ('stg-test-proof', 'stg-test-assignment', 'stg-test-task',
   'recipient@dhananipeg.com', 'Recipient', 'Synthetic proof', 'pending',
   '2026-08-11T00:00:02Z');

INSERT INTO proof_files
  (id, submission_id, drive_provider, file_name, web_url, created_at)
VALUES
  ('stg-test-file', 'stg-test-proof', 'staging', 'proof.txt', 'about:blank',
   '2026-08-11T00:00:02Z');

UPDATE tasks
   SET status = 'In Progress', version = version + 1
 WHERE id = 'stg-test-task' AND version = 1;

INSERT INTO assignments
  (id, app_task_id, title, summary, dept, priority, assigner_email,
   assigner_name, recipient_email, recipient_name, status, proof_status,
   created_at, updated_at, version)
VALUES
  ('legacy-only-assignment', 'legacy-only-task', 'Legacy-only task',
   'Imported from an assignment', 'Operations', 'High',
   'assigner@dhananipeg.com', 'Assigner', 'recipient@dhananipeg.com',
   'Recipient', 'In Progress', 'none', '2026-08-10T00:00:00Z',
   '2026-08-11T00:00:00Z', 3);
SQL

sqlite3 "$test_db" < "$repo_dir/cloudflare-worker/migrations/0003_backfill_tasks_from_assignments.sql"

[ "$(sqlite3 "$test_db" "SELECT version || '|' || status FROM tasks WHERE id='stg-test-task';")" = "2|In Progress" ]

[ "$(sqlite3 "$test_db" "UPDATE tasks SET status='Done', version=version+1 WHERE id='stg-test-task' AND version=1; SELECT changes();")" = "0" ]
[ "$(sqlite3 "$test_db" "SELECT COUNT(*) FROM reminders WHERE assignment_id='stg-test-assignment';")" = "1" ]
[ "$(sqlite3 "$test_db" "SELECT COUNT(*) FROM proof_files WHERE submission_id='stg-test-proof';")" = "1" ]
[ "$(sqlite3 "$test_db" "SELECT COUNT(*) FROM tasks WHERE id='stg-test-task';")" = "1" ]
[ "$(sqlite3 "$test_db" "SELECT status || '|' || source_type || '|' || version FROM tasks WHERE id='legacy-only-task';")" = "In Progress|legacy_assignment|3" ]
[ "$(sqlite3 "$test_db" "SELECT COUNT(*) FROM migration_records WHERE source_store='d1_assignments' AND source_id='stg-test-task';")" = "1" ]
[ -z "$(sqlite3 "$test_db" "PRAGMA foreign_key_check;")" ]

echo "staging workflow smoke test passed"
