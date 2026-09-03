// Read-only Sentry probe for the whole GSADUs org — GET-ONLY BY CONSTRUCTION: the
// single `get()` helper is the only network path in this file and it never sends a
// body or a non-GET method (that is the security property; keep it that way). Issue
// status changes, assignments, and comments stay in the Sentry UI or the claude.ai
// Sentry connector.
//
// Why this lives in Tools, not in one app repo: one Sentry org (`gsadus`) holds every
// project (webapp, pm, ...), and one personal token reads all of them. An agent in any
// repo runs it by absolute path (or the `sentry-probe` shell-profile function) and passes
// `--project <slug>`; nothing here depends on the caller's cwd or repo.
//
// Why it exists at all: the Sentry MCP connector is fine for exploration but costs a
// dozen tool rounds and ~100KB of prose per question; one call here returns the same
// facts as compact lines.
//
// Auth (never printed): SENTRY_AUTH_TOKEN from the environment if set, otherwise read
// at call time from Doppler (`webapp/dev`) — the one place the token lives. Personal
// token scopes: project:read + event:read + org:read + project:releases.
//
// Usage (node >= 18; run from any directory):
//   node C:/GSADUs/Tools/Sentry/sentry-probe.mjs projects
//   node C:/GSADUs/Tools/Sentry/sentry-probe.mjs issues  [--project pm] [--query "is:unresolved"] [--period 14d] [--limit 25]
//   node C:/GSADUs/Tools/Sentry/sentry-probe.mjs issue   <SHORT_ID>          # e.g. WEBAPP-6 / PM-10 — detail + latest event
//   node C:/GSADUs/Tools/Sentry/sentry-probe.mjs events  [--project webapp] [--query "transaction:/site-check"] [--period 14d]
//                                                        [--fields "timestamp,title,browser.name"] [--limit 50]
//   add --json to any command for the raw API payload.
import { spawnSync } from 'node:child_process';

const ORG = process.env.SENTRY_ORG || 'gsadus';
const BASE = process.env.SENTRY_API_BASE || 'https://us.sentry.io/api/0';
const DOPPLER = { project: process.env.SENTRY_TOKEN_DOPPLER_PROJECT || 'webapp', config: process.env.SENTRY_TOKEN_DOPPLER_CONFIG || 'dev' };

const args = process.argv.slice(2);
const command = args[0];
const json = args.includes('--json');
function flag(name, fallback) {
  const i = args.indexOf(name);
  return i !== -1 && args[i + 1] !== undefined ? args[i + 1] : fallback;
}

function resolveToken() {
  if (process.env.SENTRY_AUTH_TOKEN) return process.env.SENTRY_AUTH_TOKEN;
  const r = spawnSync(
    'doppler',
    ['secrets', 'get', 'SENTRY_AUTH_TOKEN', '--project', DOPPLER.project, '--config', DOPPLER.config, '--plain'],
    { encoding: 'utf8' },
  );
  if (r.error || r.status !== 0 || !r.stdout.trim()) {
    throw new Error(
      `SENTRY_AUTH_TOKEN not in env and Doppler read failed (${DOPPLER.project}/${DOPPLER.config}). ` +
        'Run `doppler login`, or export SENTRY_AUTH_TOKEN for this shell.',
    );
  }
  return r.stdout.trim();
}
const TOKEN = resolveToken();

async function get(path, params = {}) {
  const url = new URL(`${BASE}${path}`);
  for (const [k, v] of Object.entries(params)) {
    if (v === undefined || v === null || v === '') continue;
    if (Array.isArray(v)) v.forEach((x) => url.searchParams.append(k, x));
    else url.searchParams.set(k, String(v));
  }
  const res = await fetch(url, { method: 'GET', headers: { Authorization: `Bearer ${TOKEN}` } });
  if (!res.ok) {
    const body = await res.text().catch(() => '');
    throw new Error(`Sentry ${res.status} ${path}: ${body.slice(0, 300)}`);
  }
  return res.json();
}

function out(compact, raw) {
  if (json) console.log(JSON.stringify(raw, null, 2));
  else console.log(compact);
}

async function projectId(slug) {
  if (!slug) return undefined;
  const projects = await get(`/organizations/${ORG}/projects/`);
  const hit = projects.find((p) => p.slug === slug || p.id === slug);
  if (!hit) {
    throw new Error(`No project "${slug}" in org ${ORG} (have: ${projects.map((p) => p.slug).join(', ')})`);
  }
  return hit.id;
}

function tagMap(event) {
  return Object.fromEntries((event?.tags ?? []).map((t) => [t.key, t.value]));
}

function topFrames(event, n = 6) {
  const exc = (event?.entries ?? []).find((e) => e.type === 'exception');
  const values = exc?.data?.values ?? [];
  const last = values[values.length - 1];
  const frames = last?.stacktrace?.frames ?? [];
  return frames
    .slice(-n)
    .reverse()
    .map((f) => `${f.filename ?? f.module ?? '?'}:${f.lineNo ?? '?'}:${f.colNo ?? '?'} ${f.function ?? ''}`.trim());
}

const commands = {
  async projects() {
    const projects = await get(`/organizations/${ORG}/projects/`);
    out(projects.map((p) => `${p.slug}\t${p.id}\t${p.platform ?? ''}`).join('\n'), projects);
  },

  async issues() {
    const project = await projectId(flag('--project'));
    const issues = await get(`/organizations/${ORG}/issues/`, {
      query: flag('--query', 'is:unresolved'),
      statsPeriod: flag('--period', '14d'),
      project,
      limit: flag('--limit', '25'),
      sort: flag('--sort', 'date'),
    });
    const lines = issues.map((i) =>
      [i.shortId, i.status, `n=${i.count}`, `users=${i.userCount}`, i.firstSeen, i.lastSeen, i.culprit, i.title].join('\t'),
    );
    out(lines.length ? lines.join('\n') : '(no issues)', issues);
  },

  async issue() {
    const shortId = args[1];
    if (!shortId || shortId.startsWith('--')) throw new Error('issue <SHORT_ID> required, e.g. WEBAPP-6 or PM-10');
    const { groupId } = await get(`/organizations/${ORG}/shortids/${shortId}/`);
    const [issue, event] = await Promise.all([
      get(`/organizations/${ORG}/issues/${groupId}/`),
      get(`/organizations/${ORG}/issues/${groupId}/events/latest/`),
    ]);
    const tags = tagMap(event);
    const pick = ['environment', 'release', 'transaction', 'url', 'browser', 'os', 'device', 'handled', 'mechanism', 'level'];
    const frames = topFrames(event);
    const compact = [
      `${issue.shortId}  ${issue.title}`,
      `status=${issue.status}  count=${issue.count}  users=${issue.userCount}  platform=${issue.platform}  project=${issue.project?.slug}`,
      `first=${issue.firstSeen}  last=${issue.lastSeen}  culprit=${issue.culprit}`,
      `latest event ${event.eventID} @ ${event.dateCreated}`,
      ...pick.filter((k) => tags[k] !== undefined).map((k) => `  ${k}=${tags[k]}`),
      'top frames (innermost first):',
      ...(frames.length ? frames.map((f) => `  ${f}`) : ['  (no stacktrace)']),
    ];
    out(compact.join('\n'), { issue, latestEvent: event });
  },

  async events() {
    const project = await projectId(flag('--project'));
    const fields = flag(
      '--fields',
      'timestamp,title,issue,project,environment,release,transaction,url,browser.name,os.name,device.family',
    ).split(',');
    const payload = await get(`/organizations/${ORG}/events/`, {
      field: fields,
      query: flag('--query', ''),
      statsPeriod: flag('--period', '14d'),
      project,
      dataset: 'errors',
      sort: flag('--sort', '-timestamp'),
      per_page: flag('--limit', '50'),
    });
    const rows = payload.data ?? [];
    const lines = [fields.join('\t'), ...rows.map((r) => fields.map((f) => r[f] ?? '').join('\t'))];
    out(rows.length ? lines.join('\n') : '(no events)', payload);
  },
};

if (!commands[command]) {
  console.error('Usage: node sentry-probe.mjs <projects|issues|issue <SHORT_ID>|events> [--project <slug>] [options] [--json]');
  process.exit(1);
}
try {
  await commands[command]();
} catch (err) {
  console.error(`ERROR: ${err.message}`);
  process.exit(1);
}
