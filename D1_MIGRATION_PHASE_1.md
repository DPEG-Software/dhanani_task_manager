# D1 Migration — Phase 1 (Staging Foundation)

## Production safety rule

Do not apply these migrations to the production D1 database yet. Phase 1 only
creates a separate staging database and validates its schema. The production
Worker, KV namespace, D1 database, GitHub Pages site, and OneDrive files remain
unchanged.

## Current storage inventory

| Data | Current source | Phase 1 decision |
|---|---|---|
| Task assignments and status | D1 `assignments` | Reproduce in staging baseline |
| Task conversation messages | D1 `task_messages` | Reproduce in staging baseline |
| Proof submissions and reviews | KV `company-state.notifications` | Add staging D1 tables; do not switch reads/writes |
| Department catalog and mappings | KV `company-state` | Add staging D1 tables; do not switch reads/writes |
| Action Log/shared task document | KV `company-state` plus user OneDrive data | Add staging `tasks`; inventory before backfill |
| Discussion notes | KV/OneDrive application document | Add staging table; preserve current behavior |
| Proof files | Employee OneDrive | Keep files in OneDrive; store only metadata/links in D1 |
| Personal contact cache and UI seen state | Browser storage/OneDrive | Keep personal; migrate only shared read state later |
| Short proof links and edit-presence locks | KV | Keep in KV because records are temporary |

## Files created

- `cloudflare-worker/migrations/0001_staging_baseline.sql`
- `cloudflare-worker/migrations/0002_shared_workflow_staging.sql`
- `cloudflare-worker/wrangler.staging.example.jsonc`

## Create staging resources

Run these only after signing in to the correct Cloudflare account:

```bash
npx wrangler d1 create dpeg-task-manager-staging
npx wrangler kv namespace create DPEG_DATA_STAGING
```

Copy `wrangler.staging.example.jsonc` to `wrangler.staging.jsonc`, then replace
both placeholder IDs with the staging IDs returned by Cloudflare. Never put a
production database or KV ID in the staging configuration.

## Apply migrations to staging only

From `cloudflare-worker/`:

```bash
npx wrangler d1 migrations apply dpeg-task-manager-staging \
  --config wrangler.staging.jsonc \
  --remote
```

## Verification queries

```sql
SELECT name FROM sqlite_schema WHERE type = 'table' ORDER BY name;
SELECT name FROM sqlite_schema WHERE type = 'index' ORDER BY name;
SELECT name, value FROM feature_flags ORDER BY name;
PRAGMA foreign_key_check;
```

Expected feature flags:

- `shared_storage_read_mode = legacy`
- `shared_storage_dual_write = off`

These defaults guarantee that creating the staging schema does not switch any
production application behavior.

## Rollback for Phase 1

There is no production rollback because Phase 1 does not touch production. If
the staging schema is incorrect, delete only the explicitly named staging D1
database and recreate it. Do not delete or modify `DPEG_ASSIGNMENTS` or the
production `DPEG_DATA` namespace.

## Gate before Phase 2

Do not begin dual-write implementation until all of the following are true:

1. The migrations apply successfully to a fresh local SQLite database.
2. They apply successfully to the separate staging D1 database.
3. `PRAGMA foreign_key_check` returns no rows.
4. The staging Worker has separate D1 and KV bindings.
5. Production resource IDs have not been copied into staging configuration.
