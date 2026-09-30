// Supabase egress probe for the one shared project (gsadus-web-catalog) - READ-ONLY BY
// CONSTRUCTION: `readOnly()` is the only path to the database. It opens a READ ONLY
// transaction, runs statements defined as constants in this file (never caller-supplied SQL),
// and always rolls back. Keep it that way.
//
// Why it exists: the org is on the Free plan (5 GB egress per cycle, cycles start on the 6th)
// and every repo spends the same budget through the same Shared Pooler. Spec, budget math and
// the patterns that keep reads cheap: Vault wiki/curated/supabase-egress-budget.md.
//
// What it measures: pg_stat_statements counts calls and rows, never bytes. The probe estimates
// bytes as rows x the row width of the statement's first FROM (or COPY) relation, from pg_stats
// (fallback: on-disk bytes per row; 100 B when no relation matches). Writes without RETURNING
// count as zero. The estimate is a floor: text encoding and per-message protocol overhead make
// the billed figure larger. WIRE_FACTOR is the observed ratio (calibrated 2026-09-29: ~3.5 GB
// estimated since the 09-09 stats reset against 5.86 GB billed since 09-06).
// The billed cycle total is not in the public Management API (v1 exposes API request counts
// only); the org usage page is the meter.
//
// `check` saves a small snapshot (role totals plus the top statements' row counts) under .state\
// so the next run reports a rate since that snapshot instead of the average since the stats reset.
// One check fetches about 15 KB. Diagnostics cost egress too: prefer `check` over repeated `top`.
//
// Auth (never printed): EGRESS_PROBE_DB_URL from the environment if set, otherwise
// SUPABASE_DB_URL read at call time from Doppler core/prd. TLS to the pooler is verified
// against the Supabase root CA pinned in ..\Helpdesk\.
//
// Usage (node >= 20, any cwd; `egress-probe` is the pwsh shell-profile function):
//   node C:/GSADUs/Tools/Supabase/egress-probe.mjs check  [--max-mb-day 100] [--since latest|<snapshot>] [--min-age-h 1] [--limit 10] [--no-save]
//   node C:/GSADUs/Tools/Supabase/egress-probe.mjs roles                        # totals by role since the stats reset
//   node C:/GSADUs/Tools/Supabase/egress-probe.mjs top    [--role webapp_service] [--limit 20]
//   node C:/GSADUs/Tools/Supabase/egress-probe.mjs snapshots                    # local baselines (no DB call)
//   add --json for the raw payload.
// Exit: 0 ok · 1 error · 2 check over the limit.
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const CA_FILE = path.join(HERE, '..', 'Helpdesk', 'supabase-root-ca-2021.pem');
const STATE_FILE = path.join(HERE, '.state', 'snapshots.json');
const DOPPLER = { project: 'core', config: 'prd', name: 'SUPABASE_DB_URL' };
const PROJECT = 'gsadus-web-catalog (rfguhdxdrzbqcoiulyuv)';
const USAGE_PAGE = 'https://supabase.com/dashboard/org/luemjnmkgzhppwrgncsa/usage';
// pg-connection-string turns these into an `ssl` object that would replace the verifying one.
const DSN_SSL_KEYS = ['ssl', 'sslmode', 'sslrootcert', 'sslcert', 'sslkey', 'sslnegotiation'];

const BUDGET_BYTES = 5e9; // Free plan: 5 GB uncached egress per cycle, all services together
const CYCLE_START_DAY = 6;
const WIRE_FACTOR = 1.5; // billed bytes per estimated byte (calibration above)
const FALLBACK_WIDTH = 100;
const KEEP = 300; // statements kept per snapshot as the next baseline
const SNAPSHOTS_KEPT = 60;
const DEFAULT_MAX_MB_DAY = 100; // ~60% of the 167 MB/day budget; the rest covers auth, storage, spikes
// Platform-side roles: their queries run next to the database (Auth, Realtime, PostgREST's
// schema cache, Supavisor's auth_query) and are not billed as egress to a client.
const INTERNAL = /^(supabase_\w+|pgbouncer|authenticator)$/;

const args = process.argv.slice(2);
const command = args[0];
const json = args.includes('--json');
function flag(name, fallback) {
  const i = args.indexOf(name);
  return i !== -1 && args[i + 1] !== undefined && !args[i + 1].startsWith('--') ? args[i + 1] : fallback;
}
function numFlag(name, fallback) {
  const v = Number(flag(name, fallback));
  if (!Number.isFinite(v) || v < 0) throw new Error(`${name} needs a non-negative number`);
  return v;
}

// -- Statements (the only SQL this file runs) -----------------------------------------------
const EST_CTE = String.raw`
rel as (
  select n.nspname, c.relname,
         coalesce(st.w, case when c.reltuples > 0 then (pg_relation_size(c.oid) / c.reltuples)::int end) as w
    from pg_class c
    join pg_namespace n on n.oid = c.relnamespace
    left join (select schemaname, tablename, sum(avg_width)::int as w from pg_stats group by 1, 2) st
      on st.schemaname = n.nspname and st.tablename = c.relname
   where c.relkind in ('r', 'v', 'm', 'p', 'f')
     and n.nspname not in ('pg_catalog', 'information_schema', 'pg_toast')
),
stmt as (
  select s.userid::text || ':' || s.queryid::text as key, r.rolname as role, s.calls, s.rows,
         s.query, s.stats_since,
         s.query ~* '^\s*(insert|update|delete|merge)\M' and s.query !~* '\mreturning\M' as no_result,
         coalesce(regexp_match(s.query, '^\s*copy\s+(?:"?(\w+)"?\.)?"?(\w+)"?', 'i'),
                  regexp_match(s.query, '\mfrom\s+(?:"?(\w+)"?\.)?"?(\w+)"?', 'i')) as m
    from pg_stat_statements s
    join pg_roles r on r.oid = s.userid
),
est as (
  select stmt.key, stmt.role, stmt.calls, stmt.rows, stmt.query, stmt.stats_since,
         x.nspname || '.' || x.relname as rel, x.w,
         case when stmt.no_result then 0 else coalesce(x.w, ${FALLBACK_WIDTH}) end as eff_w,
         (case when stmt.no_result then 0 else stmt.rows * coalesce(x.w, ${FALLBACK_WIDTH}) end)::bigint as est
    from stmt
    left join lateral (
      select rel.nspname, rel.relname, rel.w from rel
       where rel.relname = stmt.m[2] and (stmt.m[1] is null or rel.nspname = stmt.m[1])
       order by rel.nspname = 'public' desc
       limit 1) x on true
)`;

// Deltas since a baseline. The snapshot keeps ROWS per statement, never bytes: widths come from
// pg_stats and move whenever a table is analyzed, so bytes are always rows x today's width.
// $1 prior map {key: rows} · $2 prior taken_at (null: no baseline, delta = cumulative)
// $3 prior cutoff: the smallest estimate kept. A statement older than the baseline but missing
// from the map was below it then, so its growth is at least its estimate now minus the cutoff.
const DELTA_CTE = String.raw`${EST_CTE},
d as (
  select est.*,
         case
           when $2::timestamptz is null or est.stats_since > $2::timestamptz then est.rows
           when $1::jsonb ? est.key then
             case when est.rows >= ($1::jsonb ->> est.key)::bigint
                  then est.rows - ($1::jsonb ->> est.key)::bigint else est.rows end
         end as delta_rows,
         case
           when $2::timestamptz is null then 'cumulative'
           when est.stats_since > $2::timestamptz then 'new'
           when $1::jsonb ? est.key then 'exact'
           else 'floor'
         end as basis
    from est
),
dd as (
  select d.*, coalesce(d.delta_rows * d.eff_w, greatest(d.est - $3::bigint, 0))::bigint as delta from d
)`;

const NO_BASELINE = ['{}', null, 0];

const SQL = {
  meta: `select (select stats_reset from pg_stat_statements_info) as stats_reset, now() as now`,
  roles: `with ${DELTA_CTE}
    select role, sum(calls)::bigint as calls, sum(rows)::bigint as rows, sum(est)::bigint as est,
           sum(delta)::bigint as delta
      from dd group by role order by est desc`,
  // $4 limit · $5 role filter (null: all) · $6 rank by delta (else by cumulative est)
  stmts: String.raw`with ${DELTA_CTE}
    select role, calls, rows, rel, w, est, delta_rows, delta, basis,
           left(regexp_replace(query, '\s+', ' ', 'g'), 160) as query
      from dd
     where $5::text is null or role = $5::text
     order by case when $6::boolean then delta else est end desc, est desc
     limit $4`,
  // $1 keep
  keep: `with ${EST_CTE}
    select coalesce(jsonb_object_agg(key, rows), '{}'::jsonb) as map, min(est)::bigint as cutoff
      from (select key, rows, est from est where est > 0 order by est desc limit $1) k`,
};

// -- Connection -------------------------------------------------------------------------------
function resolveDsn() {
  if (process.env.EGRESS_PROBE_DB_URL) return process.env.EGRESS_PROBE_DB_URL;
  const r = spawnSync('doppler',
    ['secrets', 'get', DOPPLER.name, '--project', DOPPLER.project, '--config', DOPPLER.config, '--plain'],
    { encoding: 'utf8', timeout: 15000, windowsHide: true });
  if (r.error || r.status !== 0 || !r.stdout.trim()) {
    throw new Error(`${DOPPLER.name} is not in the environment and the Doppler read failed ` +
      `(${DOPPLER.project}/${DOPPLER.config}); run \`doppler login\``);
  }
  return r.stdout.trim();
}

async function readOnly(work) {
  let pg;
  try { pg = (await import('pg')).default; }
  catch { throw new Error(`dependencies are not installed; run \`npm ci --prefix ${HERE}\``); }
  const target = new URL(resolveDsn());
  for (const key of DSN_SSL_KEYS) target.searchParams.delete(key);
  const client = new pg.Client({
    connectionString: target.toString(),
    ssl: { ca: fs.readFileSync(CA_FILE, 'utf8'), rejectUnauthorized: true },
    connectionTimeoutMillis: 10000,
    application_name: 'egress-probe',
  });
  await client.connect();
  try {
    await client.query('begin transaction read only');
    await client.query("set local statement_timeout = '30s'");
    return await work((name, params = []) => client.query(SQL[name], params).then((r) => r.rows));
  } finally {
    await client.query('rollback').catch(() => {});
    await client.end().catch(() => {});
  }
}

// -- Snapshots (local, per machine) -------------------------------------------------------------
function loadSnapshots() {
  try { return JSON.parse(fs.readFileSync(STATE_FILE, 'utf8')); } catch { return []; }
}
function saveSnapshot(snap) {
  const all = [...loadSnapshots(), snap].slice(-SNAPSHOTS_KEPT);
  fs.mkdirSync(path.dirname(STATE_FILE), { recursive: true });
  const tmp = `${STATE_FILE}.${process.pid}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(all));
  fs.renameSync(tmp, STATE_FILE);
}
// A baseline must come from the same stats epoch: a reset makes every counter restart.
function pickBaseline(snaps, meta, since, minAgeH) {
  const same = snaps.filter((s) => s.stats_reset === iso(meta.stats_reset));
  if (since === 'latest') return same.at(-1) ?? null;
  if (since) {
    const hit = snaps.find((s) => s.taken_at === since || s.taken_at.startsWith(since));
    if (!hit) throw new Error(`no snapshot "${since}" (run \`snapshots\` to list them)`);
    if (hit.stats_reset !== iso(meta.stats_reset)) throw new Error(`snapshot ${hit.taken_at} predates the last stats reset`);
    return hit;
  }
  const cutoff = new Date(meta.now).getTime() - minAgeH * 3600e3;
  return same.filter((s) => new Date(s.taken_at).getTime() <= cutoff).at(-1) ?? null;
}

// -- Formatting -------------------------------------------------------------------------------
const iso = (d) => (d ? new Date(d).toISOString() : null);
const MB = (b) => Number(b) / 1e6;
const fmtMB = (b) => MB(b).toLocaleString('en-US', { maximumFractionDigits: MB(b) < 10 ? 1 : 0, minimumFractionDigits: MB(b) < 10 ? 1 : 0 });
const fmtN = (n) => Number(n).toLocaleString('en-US');
const pad = (s, n) => String(s).padEnd(n);
const lpad = (s, n) => String(s).padStart(n);
const short = (d) => iso(d).slice(0, 16).replace('T', ' ') + 'Z';
function span(ms) {
  const h = ms / 3600e3;
  return h < 48 ? `${h.toFixed(1)} h` : `${(h / 24).toFixed(1)} d`;
}
function cycle(now) {
  const d = new Date(now);
  let start = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), CYCLE_START_DAY));
  if (start > d) start = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() - 1, CYCLE_START_DAY));
  const end = new Date(Date.UTC(start.getUTCFullYear(), start.getUTCMonth() + 1, CYCLE_START_DAY));
  const days = (end - start) / 86400e3;
  const day = Math.floor((d - start) / 86400e3) + 1;
  return { start, end, days, day, perDay: BUDGET_BYTES / days };
}
function cycleLine(now) {
  const c = cycle(now);
  return `cycle ${iso(c.start).slice(0, 10)} → ${iso(c.end).slice(0, 10)}, day ${c.day}/${c.days} · ` +
    `budget 5 GB ≈ ${fmtMB(c.perDay)} MB/day for everything · billed total: ${USAGE_PAGE}`;
}
function stmtLines(rows, deltaMode) {
  return rows.map((r) => {
    const amount = deltaMode ? r.delta : r.est;
    const mark = deltaMode ? { exact: ' ', new: '+', floor: '≥', cumulative: ' ' }[r.basis] : ' ';
    const width = r.w != null ? `${r.w} B/row` : `~${FALLBACK_WIDTH} B/row`;
    const rows = deltaMode ? (r.delta_rows != null ? `${fmtN(r.delta_rows)} rows Δ` : 'rows Δ ?') : `${fmtN(r.rows)} rows`;
    return `${mark}${lpad(fmtMB(amount), 8)} MB  ${pad(r.role, 16)} ${lpad(rows, 17)}  ${lpad(fmtN(r.calls), 9)} calls  ` +
      `${r.rel ?? '(no table matched)'} ${width}\n            ${r.query}`;
  });
}

// -- Commands ---------------------------------------------------------------------------------
const commands = {
  async roles() {
    const { meta, roles } = await readOnly(async (q) => ({ meta: (await q('meta'))[0], roles: await q('roles', NO_BASELINE) }));
    const days = (new Date(meta.now) - new Date(meta.stats_reset)) / 86400e3;
    const lines = [
      `${PROJECT} · pg_stat_statements since ${short(meta.stats_reset)} (${span(new Date(meta.now) - new Date(meta.stats_reset))})`,
      cycleLine(meta.now),
      `${pad('ROLE', 34)}${lpad('CALLS', 12)}${lpad('ROWS', 14)}${lpad('EST MB', 10)}${lpad('MB/DAY', 9)}`,
      ...roles.map((r) => `${pad(r.role + (INTERNAL.test(r.role) ? ' (internal)' : ''), 34)}${lpad(fmtN(r.calls), 12)}` +
        `${lpad(fmtN(r.rows), 14)}${lpad(fmtMB(r.est), 10)}${lpad(fmtMB(r.est / days), 9)}`),
      `EST MB = rows × row width, a floor; billed ≈ ×${WIRE_FACTOR}. Internal roles are platform-side, not client egress.`,
    ];
    out(lines.join('\n'), { meta, roles });
  },

  async top() {
    const limit = numFlag('--limit', 20);
    const role = flag('--role', null);
    const { meta, rows } = await readOnly(async (q) => ({
      meta: (await q('meta'))[0],
      rows: await q('stmts', [...NO_BASELINE, limit, role, false]),
    }));
    const lines = [
      `${PROJECT} · top statements by estimated bytes since ${short(meta.stats_reset)}${role ? ` · role ${role}` : ''}`,
      ...stmtLines(rows, false),
    ];
    out(lines.join('\n'), { meta, rows });
  },

  async check() {
    const maxMbDay = numFlag('--max-mb-day', DEFAULT_MAX_MB_DAY);
    const minAgeH = numFlag('--min-age-h', 1);
    const limit = numFlag('--limit', 10);
    const since = flag('--since', null);
    const snaps = loadSnapshots();
    const result = await readOnly(async (q) => {
      const meta = (await q('meta'))[0];
      const base = pickBaseline(snaps, meta, since, minAgeH);
      const baseline = base ? [JSON.stringify(base.stmts), base.taken_at, base.cutoff] : NO_BASELINE;
      const roles = await q('roles', baseline);
      const rows = await q('stmts', [...baseline, limit, null, true]);
      const keep = (await q('keep', [KEEP]))[0];
      return { meta, base, roles, rows, keep };
    });
    const { meta, base, roles, rows, keep } = result;
    const now = new Date(meta.now);
    const from = base ? new Date(base.taken_at) : new Date(meta.stats_reset);
    const days = Math.max((now - from) / 86400e3, 1 / 1440);
    // Calls and rows by role are exact counter differences; bytes are the server-side sum of each
    // statement's row delta x today's width (DELTA_CTE).
    const deltaRoles = roles.map((r) => {
      const p = base?.roles?.[r.role];
      const restarted = p && Number(r.rows) < p.rows;
      const d = (k) => (p && !restarted ? Number(r[k]) - p[k] : Number(r[k]));
      return { role: r.role, internal: INTERNAL.test(r.role), calls: d('calls'), rows: d('rows'), est: Number(r.delta) };
    }).filter((r) => r.calls || r.rows || r.est);
    const clientEst = deltaRoles.filter((r) => !r.internal).reduce((s, r) => s + r.est, 0);
    const billedPerDay = (clientEst * WIRE_FACTOR) / days;
    const over = MB(billedPerDay) > maxMbDay;
    const shortWindow = (now - from) < 3600e3;

    if (!args.includes('--no-save')) {
      saveSnapshot({
        taken_at: iso(meta.now),
        stats_reset: iso(meta.stats_reset),
        roles: Object.fromEntries(roles.map((r) => [r.role, { calls: Number(r.calls), rows: Number(r.rows), est: Number(r.est) }])),
        stmts: keep.map,
        cutoff: Number(keep.cutoff ?? 0),
      });
    }

    const verdict = over ? 'OVER' : 'OK';
    const lines = [
      `${PROJECT} · egress check ${short(meta.now)}`,
      base ? `baseline: snapshot ${short(base.taken_at)} (${span(now - from)} ago)`
        : `baseline: stats reset ${short(meta.stats_reset)} (${span(now - from)} ago; no snapshot ≥ ${minAgeH} h old on this machine)`,
      cycleLine(meta.now),
      `${pad('ROLE', 34)}${lpad('CALLS Δ', 12)}${lpad('ROWS Δ', 14)}${lpad('EST MB Δ', 10)}${lpad('MB/DAY', 9)}`,
      ...deltaRoles.map((r) => `${pad(r.role + (r.internal ? ' (internal)' : ''), 34)}${lpad(fmtN(r.calls), 12)}` +
        `${lpad(fmtN(r.rows), 14)}${lpad(fmtMB(r.est), 10)}${lpad(fmtMB(r.est / days), 9)}`),
      `client-facing: ${fmtMB(clientEst)} MB estimated → ${fmtMB(billedPerDay)} MB/day billed-equivalent (×${WIRE_FACTOR}); ` +
        `limit ${maxMbDay} MB/day → ${verdict}${shortWindow ? ' (window under 1 h: the rate is noisy)' : ''}`,
      `top statements since the baseline (+ new since then, ≥ lower bound: was below the kept top ${KEEP}):`,
      ...stmtLines(rows.filter((r) => Number(r.delta) > 0), true),
    ];
    out(lines.join('\n'), { meta, baseline: base ? { taken_at: base.taken_at } : null, roles: deltaRoles, top: rows,
      client_est_bytes: clientEst, billed_equivalent_mb_per_day: MB(billedPerDay), limit_mb_per_day: maxMbDay, verdict });
    if (over) process.exitCode = 2;
  },

  async snapshots() {
    const snaps = loadSnapshots();
    const lines = snaps.map((s) => {
      const est = Object.entries(s.roles).filter(([r]) => !INTERNAL.test(r)).reduce((t, [, v]) => t + v.est, 0);
      return `${s.taken_at}  stats since ${s.stats_reset.slice(0, 10)}  client-facing est ${fmtMB(est)} MB`;
    });
    out(lines.length ? lines.join('\n') : `(no snapshots yet: \`check\` saves one to ${STATE_FILE})`,
      snaps.map(({ stmts, ...rest }) => rest));
  },
};

function out(compact, raw) {
  if (json) console.log(JSON.stringify(raw, null, 2));
  else console.log(compact);
}

if (!commands[command]) {
  console.error('Usage: node egress-probe.mjs <check|roles|top|snapshots> [--max-mb-day N] [--since latest|<snapshot>] ' +
    '[--role <role>] [--limit N] [--no-save] [--json]');
  process.exit(1);
}
try {
  await commands[command]();
} catch (err) {
  console.error(`ERROR: ${err.message}`);
  process.exit(1);
}
