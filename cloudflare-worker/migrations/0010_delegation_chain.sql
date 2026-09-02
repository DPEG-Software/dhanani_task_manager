ALTER TABLE assignments ADD COLUMN parent_assignment_id TEXT;
ALTER TABLE assignments ADD COLUMN root_assignment_id TEXT;
ALTER TABLE assignments ADD COLUMN delegation_level INTEGER NOT NULL DEFAULT 0;
ALTER TABLE assignments ADD COLUMN delegated_to_email TEXT;
ALTER TABLE assignments ADD COLUMN delegated_to_name TEXT;
ALTER TABLE assignments ADD COLUMN forwarded_review_note TEXT;

CREATE INDEX IF NOT EXISTS idx_assignments_parent ON assignments(parent_assignment_id);
CREATE INDEX IF NOT EXISTS idx_assignments_root ON assignments(root_assignment_id);
