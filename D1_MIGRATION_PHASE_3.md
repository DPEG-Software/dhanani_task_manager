# D1 Migration — Phase 3 Inventory and Backfill Rehearsal

## Safety status

- Production frontend and Worker behavior remain unchanged.
- Production feature flags have not been enabled.
- All production inventory queries were read-only and reported zero rows written.
- The real-interface staging test passed with two users and realtime updates.
- No paid Cloudflare plan is required for this phase.

## Production inventory (2026-08-12)

| Resource | ID | Inventory |
|---|---|---|
| D1 `dpeg-assignments` | `ad39f4aa-f2cc-488b-8c4b-d0f072ad9309` | 39 assignments, 8 task messages |
| KV `DPEG_DATA` | `4ff9c0fa5eac4512b384cb254f0f7b15` | No keys at inventory time |
| Employee task documents | Microsoft OneDrive | Remain authoritative until a rehearsed backfill is verified |

Production D1 Time Travel bookmark captured before backfill work:

`000006bd-00000000-000050c5-a3580a762e26715dcbedb47ba4d2405e`

Do not restore this bookmark unless a production rollback is explicitly
authorized. Time Travel restore is destructive to changes made after the
bookmark.

## Backfill rehearsal rule

`0003_backfill_tasks_from_assignments.sql` is idempotent. It creates a
normalized `tasks` row only when that `app_task_id` does not already exist and
records the source in `migration_records`. It never updates or deletes an
existing task.

The next rehearsal copies production D1 data into staging under a controlled
snapshot, applies the normalization migration, and compares counts. It must not
write to production or enable `shared_storage_dual_write` or
`shared_storage_read_mode`.

## Offline production-snapshot rehearsal result

Completed successfully on 2026-08-12 using a Git-ignored local export:

| Check | Result |
|---|---:|
| Source assignments | 39 |
| Source task messages | 8 |
| Distinct source task IDs | 39 |
| Normalized tasks created | 39 |
| Migration records created | 39 |
| Missing required task fields | 0 |
| Orphan task messages | 0 |
| Assignments missing a normalized task | 0 |
| Foreign-key errors | 0 |

Normalized status distribution: 12 Pending, 5 In Progress, 21 Done and 1
Cancelled.

Production was queried after the export and remained responsive with 39
assignments. No production rows were written, and the production Worker and
frontend were not deployed.

## Gate before production schema expansion

The next step creates the normalized tables in production with both feature
flags defaulting to `off`/`legacy`. Although this does not switch application
behavior, applying a D1 migration can briefly make the database unavailable.
Schedule it for a low-usage maintenance window and take a fresh Time Travel
bookmark immediately beforehand. Do not perform it during normal employee use.

## Production schema expansion result

Completed successfully during an authorized maintenance window on 2026-08-12.
The database migration statements completed in milliseconds.

- Original assignments preserved: 39
- Original task messages preserved: 8
- Normalized tasks created: 39
- Migration records created: 39
- Missing required fields: 0
- Orphan messages: 0
- Assignments without normalized tasks: 0
- Foreign-key errors: 0
- Pending migrations: 0
- Production Worker availability check: HTTP 204
- `shared_storage_dual_write`: `off`
- `shared_storage_read_mode`: `legacy`

Pre-migration rollback bookmark:

`000006bd-00000000-000050c5-a3580a762e26715dcbedb47ba4d2405e`

No frontend or production Worker code was deployed. The application continues
using its pre-migration behavior until dual-write is separately implemented,
tested, and explicitly enabled.
