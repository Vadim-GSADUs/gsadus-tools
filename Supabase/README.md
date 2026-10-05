# Supabase

`egress-probe.mjs` is the egress guard for the one Supabase project every repo shares,
`gsadus-web-catalog` (`rfguhdxdrzbqcoiulyuv`). It is read-only by construction: one helper opens a
`READ ONLY` transaction, runs the fixed statements in the file and always rolls back.

- Budget, method and each repo's guard: Vault `wiki/curated/supabase-egress-budget.md`.
- The rule it enforces: `C:\GSADUs\AGENTS.md` → Rules for AI Agents.

| File | Purpose |
|---|---|
| `egress-probe.mjs` | The probe. Node ≥ 20; `pg` is its only dependency. |
| `Register-EgressCheckTask.ps1` | Registers (or `-Unregister`s) the daily `\GSADUs\supabase-egress-check` task. |
| `.state\snapshots.json` | Local baselines saved by `check`. Gitignored and per machine. |
| `.state\check.log` | The scheduled task's output. Gitignored. |

```powershell
egress-probe check                    # shell-profile function (pwsh); installs pg on first use
egress-probe check --since latest     # what the window since the last check cost (a test loop, a deploy)
egress-probe roles                    # totals by role since the stats reset
egress-probe top --role pm_service    # one role's statements by estimated bytes
egress-probe snapshots                # the local baselines (no database call)
node C:/GSADUs/Tools/Supabase/egress-probe.mjs check   # same thing from Git Bash / any shell
```

Add `--json` for the raw payload.

## What `check` does

1. **Picks a baseline:** the newest local snapshot at least `--min-age-h` old (default 1), from
   the same stats epoch. With no such snapshot it uses the pg_stat_statements reset instead.
2. **Reports the change since then:** calls, rows and estimated bytes by role, plus the
   statements that grew most.
3. **Saves a new snapshot**, unless you pass `--no-save`.
4. **Applies the limit:** exits **2** when client-facing traffic runs above `--max-mb-day`
   (default **500** MB/day billed-equivalent: about twice the traffic measured before the
   2026-09-29 fixes and 6% of the ~8,300 MB/day Pro share, an early warning rather than a quota
   line (owner 2026-10-05); it was 100 MB/day on the Free plan).

Exit codes: 0 ok, 1 error, 2 over the limit.

**Snapshots store rows, never bytes.** Row widths come from `pg_stats` and change whenever a
table is analyzed. So a delta in bytes is always (rows now − rows then) × today's width, summed
in the database. The snapshot holds the 300 statements with the largest estimates. A statement
older than the baseline but missing from it is counted at its lower bound and marked `≥`.

## The daily task

Owner decision, 2026-09-29: the task runs once a day, on **this PC only** (gsadus-vadim).
- **Register or remove it:**
  `pwsh -File C:\GSADUs\Tools\Supabase\Register-EgressCheckTask.ps1 [-Unregister]`.
- **When it runs:** `check` at 08:00, or at the next start if the PC was off.
- **Where the result goes:** its output is appended to `.state\check.log`.
- **How to read it:** `LastTaskResult` 2 means it was over the limit. The log's last block says
  which statements grew.
- **Why daily:** each run is the next one's 24 h baseline.

## How bytes are estimated

pg_stat_statements counts calls and rows, never bytes, so the probe estimates.

- **Row width.** Each statement is sized at rows × the row width of its first `FROM` (or `COPY`)
  relation, taken from `pg_stats`.
  - Fallback: on-disk bytes per row.
  - When no relation matches (a view without stats, a catalog query), 100 B per row.
- **Writes.** A write without `RETURNING` counts as zero.
- **The estimate is a floor.** Text encoding and protocol messages add bytes. Billed ≈ estimate
  × 1.5 (`WIRE_FACTOR`), from the 2026-09-29 calibration: about 3.5 GB estimated since the 09-09
  reset against 5.86 GB billed since 09-06.
- **Internal roles are listed but never count toward the limit.** These are `supabase_*`,
  `pgbouncer` and `authenticator`. Their queries run next to the database: Auth, Realtime,
  PostgREST's schema cache, Supavisor's auth query.

The billed cycle total is not in the public Management API; v1 exposes API request counts only.
The meter is the org usage page: https://supabase.com/dashboard/org/luemjnmkgzhppwrgncsa/usage

**The probe costs egress too.** One `check` returns about 15 KB, most of it the 300-statement
baseline map. Prefer one `check` over repeated `top` calls.

**Never run `pg_stat_statements_reset()`.** Every session and every baseline reads the same
counters.

## Auth

- **Where the DSN comes from.** `SUPABASE_DB_URL` (the `postgres` role, Shared Pooler session
  mode), read at call time from Doppler `core/prd`. `EGRESS_PROBE_DB_URL` in the environment
  overrides it. Never print either.
- **TLS.** Verified against the Supabase root CA pinned beside the probe
  (`supabase-root-ca-2021.pem`, expires 2031-04-26).
- **Read-only.** The `postgres` role can write; the read-only transaction is what stops it here.
- **If it ever runs unattended, give it its own role.** That means off this machine or with
  credentials shared beyond the owner's Doppler. The role holds `pg_read_all_stats` and nothing
  else, is created by a WebCatalog migration the way `helpdesk_agent` was, and has its own
  Doppler secret.
