-- Optional opaque role key used only by the authenticated staging simulator.
ALTER TABLE directory_users ADD COLUMN staging_profile_key TEXT;
CREATE UNIQUE INDEX IF NOT EXISTS idx_directory_users_staging_profile
  ON directory_users(staging_profile_key) WHERE staging_profile_key IS NOT NULL;
