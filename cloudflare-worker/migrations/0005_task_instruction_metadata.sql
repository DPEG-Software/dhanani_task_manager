-- Preserve task-specific instructions independently from summaries and keep
-- Outlook source references during the legacy-to-D1 shadow period.
ALTER TABLE tasks ADD COLUMN task_instruction TEXT NOT NULL DEFAULT '';
ALTER TABLE tasks ADD COLUMN proof_instructions TEXT NOT NULL DEFAULT '';
