ALTER TABLE recurring_schedules ADD COLUMN frequency_unit TEXT NOT NULL DEFAULT 'week';
ALTER TABLE recurring_schedules ADD COLUMN frequency_interval INTEGER NOT NULL DEFAULT 1;
