-- Rehearsal-safe normalization step.
--
-- After a copy of legacy assignment rows is imported into staging, create one
-- normalized task row per app_task_id. Existing tasks always win, making this
-- migration repeatable and safe to run more than once during rehearsals.
INSERT OR IGNORE INTO tasks (
  id, owner_email, title, summary, department_name, priority, due_date, status,
  source_type, created_at, updated_at, completed_at, cancelled_at, version
)
SELECT
  a.app_task_id,
  MIN(lower(a.assigner_email)),
  MAX(COALESCE(NULLIF(a.title, ''), 'Imported task')),
  MAX(COALESCE(a.summary, '')),
  MAX(COALESCE(NULLIF(a.dept, ''), 'Needs Department')),
  MAX(COALESCE(NULLIF(a.priority, ''), 'Normal')),
  MAX(a.due_date),
  CASE
    WHEN MAX(CASE WHEN lower(COALESCE(a.status, '')) = 'cancelled' THEN 1 ELSE 0 END) = 1
      THEN 'Cancelled'
    WHEN MAX(CASE WHEN lower(COALESCE(a.status, '')) = 'done'
                        OR lower(COALESCE(a.proof_status, '')) = 'approved' THEN 1 ELSE 0 END) = 1
      THEN 'Done'
    WHEN MAX(CASE WHEN lower(COALESCE(a.status, '')) = 'in progress' THEN 1 ELSE 0 END) = 1
      THEN 'In Progress'
    ELSE 'Pending'
  END,
  'legacy_assignment',
  MIN(a.created_at),
  MAX(a.updated_at),
  MAX(CASE WHEN lower(COALESCE(a.proof_status, '')) = 'approved' THEN a.proof_reviewed_at END),
  MAX(a.cancelled_at),
  MAX(COALESCE(a.version, 1))
FROM assignments a
WHERE COALESCE(a.app_task_id, '') <> ''
GROUP BY a.app_task_id;

INSERT OR IGNORE INTO migration_records (
  source_store, source_id, target_table, target_id, source_hash, migrated_at
)
SELECT
  'd1_assignments',
  a.app_task_id,
  'tasks',
  a.app_task_id,
  NULL,
  CURRENT_TIMESTAMP
FROM assignments a
WHERE COALESCE(a.app_task_id, '') <> ''
GROUP BY a.app_task_id;
