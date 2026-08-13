-- Additive control for per-user D1 shadow-read verification. Existing users
-- remain on legacy reads until both read mode and their explicit allow-list
-- entry are enabled.
INSERT OR IGNORE INTO feature_flags(name, value, updated_by_email, updated_at)
VALUES ('shared_storage_read_canary_users', '', 'migration', CURRENT_TIMESTAMP);
