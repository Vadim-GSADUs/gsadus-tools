// Staff helpdesk command: the owner-side tool for HD-n tickets. Spec: Vault
// wiki/curated/helpdesk.md. Schema: WebCatalog pipeline/supabase/migrations/0113_helpdesk_init.sql.
//
// It connects as the scoped `helpdesk_agent` role, which can select, insert and update but
// never delete. The database enforces the lifecycle itself: allowed transitions,
// append-only history, immutable submissions. This tool adds three things:
//   - ergonomics;
//   - the claim guard: only the session that claimed a ticket moves it on, or the owner;
//   - the owner confirmation window for approve/reject (confirm.ps1). It makes every approval
//     a deliberate owner click, in either harness and any permission mode. It is not a security
//     boundary; confirm.ps1 explains why and what the hard boundary will be.
//
// Auth (never printed): HELPDESK_DB_URL from the environment if set, otherwise read at call
// time from Doppler (core/prd), the one place it lives. TLS to the Supabase pooler is
// verified against the bundled Supabase root CA, the same pin PM uses.
//
// Every change a reporter should see is then posted to the ticket's thread in the "Tech
// Requests" Chat space as the GSADUs staff bot (chat.mjs). The database change is the record:
// a failed post is reported but never undoes it.
//
// Usage (node >= 20, any cwd):  node C:/GSADUs/Tools/Helpdesk/helpdesk.mjs <command> ...
// or the `helpdesk` shell-profile function. `helpdesk help` lists the commands.
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { chatConfig, openingText, postToThread, say } from './chat.mjs';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const CA_FILE = path.join(HERE, 'supabase-root-ca-2021.pem');
const CONFIRM_SCRIPT = path.join(HERE, 'confirm.ps1');
const DOPPLER = { project: 'core', config: 'prd', name: 'HELPDESK_DB_URL' };
const WORKSPACE = 'c:\\gsadus';
const SKILL = 'C:\\GSADUs\\.claude\\skills\\helpdesk\\SKILL.md';
const LOOPBACK = new Set(['127.0.0.1', 'localhost', '::1', '[::1]']);
// pg-connection-string turns these into an `ssl` object that would replace the verifying one.
const DSN_SSL_KEYS = ['ssl', 'sslmode', 'sslrootcert', 'sslcert', 'sslkey', 'sslnegotiation'];

export const PRODUCTS = ['webapp', 'pm', 'pngtools', 'pyrevit', 'studio', 'it', 'other'];
export const CATEGORIES = ['broken', 'confusing', 'idea', 'access'];
export const SEVERITIES = ['blocking', 'annoying', 'nice-to-have'];
export const ASSIGNEES = ['claude', 'codex', 'owner'];
const CLOSED = ['resolved', 'rejected', 'duplicate'];
const STATUS_ORDER = ['new', 'reopened', 'triaged', 'needs_info', 'approved', 'in_progress', 'fix_ready',
  'queued', 'resolved', 'rejected', 'duplicate'];
const MAX_SCREENSHOT = 2 * 1024 * 1024;
const EXT = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/webp': 'webp' };
const BOOLEAN_FLAGS = new Set(['json', 'all', 'no-files']);

export class UsageError extends Error {}

// -- Arguments ------------------------------------------------------------------------------
export function parseArgs(argv) {
  const positional = [];
  const flags = new Map();
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (!arg.startsWith('--')) { positional.push(arg); continue; }
    const eq = arg.indexOf('=');
    let name;
    let value;
    if (eq !== -1) { name = arg.slice(2, eq); value = arg.slice(eq + 1); }
    else if (BOOLEAN_FLAGS.has(arg.slice(2))) { name = arg.slice(2); value = true; }
    else {
      name = arg.slice(2);
      value = argv[++i];
      if (value === undefined) throw new UsageError(`--${name} needs a value`);
    }
    if (!flags.has(name)) flags.set(name, []);
    flags.get(name).push(value);
  }
  return { positional, flags };
}

const one = (flags, name) => { const v = flags.get(name); return v ? v[v.length - 1] : undefined; };
const all = (flags, name) => flags.get(name) || [];
function required(flags, name, hint) {
  const v = one(flags, name);
  if (v === undefined || v === true || String(v).trim() === '') {
    throw new UsageError(`--${name} is required${hint ? ` (${hint})` : ''}`);
  }
  return String(v);
}
function oneOf(value, allowed, name) {
  if (!allowed.includes(value)) throw new UsageError(`--${name} must be one of ${allowed.join(' | ')} (got ${value})`);
  return value;
}

export function parseTicket(value) {
  const m = /^(?:hd-?)?(\d+)$/i.exec(String(value ?? '').trim());
  if (!m) throw new UsageError(`expected a ticket number like 14 or HD-14, got ${value ?? 'nothing'}`);
  return Number(m[1]);
}

// Actor recorded on history rows and in claimed_by: <who>@<machine>, e.g. claude:RoseIbis@gsadus-vadim.
// Required for every command that changes a ticket, so an agent can never act as the owner by omission.
export function actorFrom(flags, env = process.env, host = os.hostname()) {
  const who = one(flags, 'as') || env.HELPDESK_ACTOR;
  if (!who || who === true) throw new UsageError('--as is required: claude:<AgentName>, codex:<AgentName>, or owner');
  const actor = String(who).includes('@') ? String(who) : `${who}@${host.toLowerCase()}`;
  if (!/^[A-Za-z0-9._:-]+@[A-Za-z0-9._-]+$/.test(actor)) {
    throw new UsageError(`--as must look like claude:AgentName or codex:AgentName (got ${who})`);
  }
  return actor;
}
const isOwner = (actor) => actor.startsWith('owner@');
const harnessOf = (actor) => actor.split('@')[0].split(':')[0];

// -- Connection -------------------------------------------------------------------------------
export function isLoopback(url) {
  try { return LOOPBACK.has(new URL(url).hostname); } catch { return false; }
}

function resolveDsn() {
  if (process.env.HELPDESK_DB_URL) return process.env.HELPDESK_DB_URL;
  const r = spawnSync('doppler',
    ['secrets', 'get', DOPPLER.name, '--project', DOPPLER.project, '--config', DOPPLER.config, '--plain'],
    { encoding: 'utf8', timeout: 15000, windowsHide: true });
  if (r.error || r.status !== 0 || !r.stdout.trim()) {
    throw new Error(`${DOPPLER.name} is not in the environment and the Doppler read failed ` +
      `(${DOPPLER.project}/${DOPPLER.config}); run \`doppler login\``);
  }
  return r.stdout.trim();
}

async function connect(url) {
  let pg;
  try { pg = (await import('pg')).default; }
  catch { throw new Error(`dependencies are not installed; run \`npm ci --prefix ${HERE}\``); }
  const target = new URL(url);
  for (const key of DSN_SSL_KEYS) target.searchParams.delete(key);
  const ssl = isLoopback(url) ? false : { ca: fs.readFileSync(CA_FILE, 'utf8'), rejectUnauthorized: true };
  const client = new pg.Client({
    connectionString: target.toString(), ssl, connectionTimeoutMillis: 10000, application_name: 'helpdesk-cli',
  });
  await client.connect();
  return client;
}

// One transaction per command, with helpdesk.actor set so trigger-written history names the actor.
async function tx(client, actor, work) {
  await client.query('begin');
  try {
    await client.query("select set_config('helpdesk.actor', $1, true)", [actor]);
    const result = await work();
    await client.query('commit');
    return result;
  } catch (e) {
    await client.query('rollback').catch(() => {});
    throw e;
  }
}

// -- Owner confirmation -----------------------------------------------------------------------
// The test override only works against a loopback database, so it can never approve a real ticket.
export function testConfirmAllowed(url, env = process.env) {
  return isLoopback(url) && ['yes', 'no'].includes(env.HELPDESK_TEST_CONFIRM);
}

function confirmOrThrow(action, text, url) {
  if (testConfirmAllowed(url)) {
    if (process.env.HELPDESK_TEST_CONFIRM === 'yes') return;
    throw new UsageError(`${action} cancelled by the owner`);
  }
  const r = spawnSync('pwsh',
    ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', CONFIRM_SCRIPT, '-Action', action],
    { env: { ...process.env, HELPDESK_CONFIRM_TEXT: text }, stdio: 'ignore', timeout: 300000 });
  if (r.status === 0) return;
  if (r.status === 1) throw new UsageError(`${action} cancelled by the owner`);
  if (r.status === 2) throw new UsageError(`${action} not confirmed: the confirmation window timed out`);
  throw new UsageError(`${action} not confirmed: the owner confirmation window could not be shown ` +
    `(${r.error?.message || `exit ${r.status}`}). Run it from a desktop session on the owner's PC.`);
}

// -- Formatting -------------------------------------------------------------------------------
const hd = (n) => `HD-${n}`;
const oneLine = (s, max) => {
  const t = String(s ?? '').replace(/\s+/g, ' ').trim();
  return t.length > max ? `${t.slice(0, max - 1)}…` : t;
};
const indent = (s) => String(s).split(/\r?\n/).map((l) => `    ${l}`).join('\n');
const local = (d) => new Date(d).toLocaleString('sv-SE', { hour12: false }).slice(0, 16);
function age(d) {
  const s = (Date.now() - new Date(d).getTime()) / 1000;
  if (s < 3600) return `${Math.max(1, Math.round(s / 60))}m`;
  if (s < 86400) return `${Math.round(s / 3600)}h`;
  return `${Math.round(s / 86400)}d`;
}

export function summarize(rows) {
  const get = (status) => rows.find((r) => r.status === status) || { n: 0, blocking: 0 };
  const fresh = get('new').n + get('reopened').n;
  const freshBlocking = get('new').blocking + get('reopened').blocking;
  const parts = [];
  if (fresh) parts.push(`${fresh} new${freshBlocking ? ` (${freshBlocking} blocking)` : ''}`);
  if (get('triaged').n) parts.push(`${get('triaged').n} to approve`);
  if (get('fix_ready').n) parts.push(`${get('fix_ready').n} fix-ready for review`);
  if (get('approved').n) parts.push(`${get('approved').n} approved, unclaimed`);
  if (get('in_progress').n) parts.push(`${get('in_progress').n} in progress`);
  if (get('needs_info').n) parts.push(`${get('needs_info').n} waiting on the reporter`);
  return parts.length ? `Helpdesk: ${parts.join(' · ')}` : '';
}

// -- Database helpers -------------------------------------------------------------------------
async function load(c, n) {
  const { rows } = await c.query('select * from helpdesk.ticket where number = $1', [n]);
  if (!rows.length) throw new UsageError(`${hd(n)} does not exist`);
  return rows[0];
}

// `sets` keys are constants from this file, never user input.
async function update(c, n, sets) {
  const cols = Object.keys(sets);
  if (!cols.length) return;
  await c.query(`update helpdesk.ticket set ${cols.map((k, i) => `${k} = $${i + 2}`).join(', ')} where number = $1`,
    [n, ...cols.map((k) => sets[k])]);
}

// Move only if the ticket is still in one of `from`, so a concurrent change is reported, not overwritten.
async function transition(c, n, from, to, sets = {}) {
  const cols = Object.keys(sets);
  const assign = ['status = $3', ...cols.map((k, i) => `${k} = $${i + 4}`)].join(', ');
  const { rows } = await c.query(
    `update helpdesk.ticket set ${assign} where number = $1 and status = any($2::text[]) returning *`,
    [n, from, to, ...cols.map((k) => sets[k])]);
  if (!rows.length) {
    const t = await load(c, n);
    throw new UsageError(`${hd(n)} is ${t.status}; moving it to ${to} needs ${from.join(' | ')}`);
  }
  return rows[0];
}

const addEvent = (c, n, actor, kind, body) => c.query(
  'insert into helpdesk.ticket_event (ticket, actor, kind, body) values ($1, $2, $3, $4)', [n, actor, kind, body ?? null]);

function linksFrom(flags) {
  return [...all(flags, 'pr').map((ref) => ({ kind: 'pr', ref })), ...all(flags, 'commit').map((ref) => ({ kind: 'commit', ref }))];
}
async function addLinks(c, n, actor, links) {
  if (!links.length) return;
  await c.query('update helpdesk.ticket set links = links || $2::jsonb where number = $1', [n, JSON.stringify(links)]);
  await addEvent(c, n, actor, 'link', links.map((l) => `${l.kind}: ${l.ref}`).join('\n'));
}

function assertWorker(t, actor) {
  if (t.claimed_by && t.claimed_by !== actor && !isOwner(actor)) {
    throw new UsageError(`${hd(t.number)} is claimed by ${t.claimed_by}; only that session or the owner moves it on`);
  }
}

function readScreenshot(file) {
  const data = fs.readFileSync(file);
  const type = data.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) ? 'image/png'
    : data[0] === 0xff && data[1] === 0xd8 && data[2] === 0xff ? 'image/jpeg'
      : data.subarray(0, 4).toString('latin1') === 'RIFF' && data.subarray(8, 12).toString('latin1') === 'WEBP' ? 'image/webp'
        : null;
  if (!type) throw new UsageError(`${file} is not a PNG, JPEG or WebP image`);
  if (data.length > MAX_SCREENSHOT) throw new UsageError(`${file} is over 2 MB`);
  return { type, data, name: path.basename(file) };
}

// -- Chat thread ------------------------------------------------------------------------------
// Opens the thread first if the ticket has none yet (filed here, or PM's post failed), then
// records the thread name the first time it is known.
async function postUpdate(c, cfg, n, actor, text) {
  const t = await load(c, n);
  let thread = t.chat_thread;
  if (!thread) thread = await postToThread(cfg, n, openingText(t));
  if (text) thread = (await postToThread(cfg, n, text)) || thread;
  if (!t.chat_thread && thread) {
    await tx(c, actor, () => c.query(
      'update helpdesk.ticket set chat_thread = $2 where number = $1 and chat_thread is null', [n, thread]));
  }
}

async function announce(c, url, { n, actor, text }) {
  try {
    const cfg = chatConfig(process.env, isLoopback(url));
    if (cfg) await postUpdate(c, cfg, n, actor, text);
  } catch (e) {
    console.error(`${hd(n)}: the change is saved, but posting it to the Chat thread failed (${oneLine(e.message, 160)})`);
    process.exitCode = 3;
  }
}

// -- Commands ---------------------------------------------------------------------------------
async function statusLine(c) {
  const { rows } = await c.query(
    `select status, count(*)::int as n, (count(*) filter (where severity = 'blocking'))::int as blocking
       from helpdesk.ticket where status <> all($1::text[]) group by status`, [[...CLOSED, 'queued']]);
  return summarize(rows);
}

async function hookStatus() {
  let input = {};
  try { const raw = fs.readFileSync(0, 'utf8'); input = raw.trim() ? JSON.parse(raw) : {}; } catch { /* no input */ }
  const cwd = String(input.cwd || process.cwd()).toLowerCase();
  if (cwd !== WORKSPACE && !cwd.startsWith(`${WORKSPACE}\\`) && !cwd.startsWith('c:/gsadus')) return;
  let line;
  let c;
  try {
    c = await connect(resolveDsn());
    line = await statusLine(c);
  } catch (e) {
    line = `Helpdesk status unavailable (${oneLine(e.message, 140)})`;
  } finally {
    await c?.end().catch(() => {});
  }
  if (!line) return;
  const context = `${line}. These are staff tickets (HD-n) about GSADUs tools; list them with \`helpdesk list\` ` +
    `(or \`node C:/GSADUs/Tools/Helpdesk/helpdesk.mjs list\`), and read ${SKILL} before acting on one. ` +
    'Mention this to the user when relevant; work tickets only when the user asks.';
  process.stdout.write(JSON.stringify({
    hookSpecificOutput: { hookEventName: input.hook_event_name || 'SessionStart', additionalContext: context },
  }));
}

async function cmdStatus(c) {
  console.log((await statusLine(c)) || 'Helpdesk: nothing open');
}

async function cmdList(c, flags) {
  const statuses = one(flags, 'status') ? String(one(flags, 'status')).split(',').map((s) => s.trim()) : null;
  const product = one(flags, 'product') ?? null;
  const { rows } = await c.query(
    `select number, status, severity, product, category, reporter_email, note, created_at, assignee, claimed_by
       from helpdesk.ticket
      where ($1::boolean or status <> all($2::text[]))
        and ($3::text[] is null or status = any($3::text[]))
        and ($4::text is null or product = $4)
      order by array_position($5::text[], status),
               array_position(array['blocking', 'annoying', 'nice-to-have'], severity), number`,
    [one(flags, 'all') === true || !!statuses, CLOSED, statuses, product, STATUS_ORDER]);
  if (one(flags, 'json') === true) { console.log(JSON.stringify(rows, null, 2)); return; }
  if (!rows.length) { console.log('No tickets.'); return; }
  for (const r of rows) {
    console.log([
      hd(r.number).padEnd(7), r.status.padEnd(11), r.severity.padEnd(12), `${r.product}/${r.category}`.padEnd(17),
      age(r.created_at).padStart(4), r.reporter_email.split('@')[0].padEnd(12), oneLine(r.note, 70),
      r.assignee ? `→${r.assignee}` : '', r.claimed_by ? `[${r.claimed_by}]` : '',
    ].join('  ').trimEnd());
  }
}

async function cmdShow(c, flags, n) {
  const t = await load(c, n);
  const { rows: events } = await c.query(
    'select at, actor, kind, from_status, to_status, body from helpdesk.ticket_event where ticket = $1 order by id', [n]);
  const { rows: shots } = await c.query(
    'select id, content_type, filename, byte_size, created_at, purged_at, data from helpdesk.attachment where ticket = $1 order by id', [n]);
  const saved = new Map();
  if (one(flags, 'no-files') !== true && shots.some((s) => s.data)) {
    const dir = path.join(os.tmpdir(), 'helpdesk', hd(n));
    fs.mkdirSync(dir, { recursive: true });
    for (const s of shots.filter((x) => x.data)) {
      const file = path.join(dir, `${s.id}.${EXT[s.content_type]}`);
      fs.writeFileSync(file, s.data);
      saved.set(s.id, file);
    }
  }
  const attachments = shots.map(({ data, ...s }) => ({ ...s, path: saved.get(s.id) ?? null }));
  if (one(flags, 'json') === true) { console.log(JSON.stringify({ ticket: t, events, attachments }, null, 2)); return; }
  const lines = [
    `${hd(n)}  [${t.status}]  ${t.product}/${t.category}  ${t.severity}`,
    `Reporter: ${t.reporter_email}   Filed: ${local(t.created_at)} (${age(t.created_at)} ago)`,
  ];
  const context = [
    t.page_url && `Page: ${t.page_url}`, t.build_sha && `Build: ${t.build_sha}`, t.sentry_event_id && `Sentry: ${t.sentry_event_id}`,
    t.viewport && `Viewport: ${t.viewport}`, t.user_agent && `Browser: ${oneLine(t.user_agent, 90)}`,
  ].filter(Boolean);
  if (context.length) lines.push(...context);
  const work = [
    t.assignee && `Assignee: ${t.assignee}`, t.claimed_by && `Claimed by: ${t.claimed_by} since ${local(t.claimed_at)}`,
    t.duplicate_of && `Duplicate of: ${hd(t.duplicate_of)}`, t.chat_thread && `Chat thread: ${t.chat_thread}`,
    t.resolved_at && `Resolved: ${local(t.resolved_at)}`,
  ].filter(Boolean);
  if (work.length) lines.push(...work);
  lines.push('Note:', indent(t.note));
  if (t.triage_note) lines.push('Triage:', indent(t.triage_note));
  if (t.resolution) lines.push('Resolution:', indent(t.resolution));
  if (t.links.length) lines.push(`Links: ${t.links.map((l) => `${l.kind}: ${l.ref}`).join('; ')}`);
  if (attachments.length) {
    lines.push('Screenshots:');
    for (const a of attachments) {
      lines.push(`    ${a.path ?? (a.purged_at ? `#${a.id} purged ${local(a.purged_at)}` : `#${a.id} (not saved; --no-files)`)}` +
        `  ${a.content_type}, ${Math.ceil(a.byte_size / 1024)} KB`);
    }
  }
  lines.push('History:');
  for (const e of events) {
    const what = e.kind === 'status' ? `${e.from_status} → ${e.to_status}` : e.kind;
    lines.push(`    ${local(e.at)}  ${what.padEnd(24)}  ${e.actor}${e.body ? `\n${indent(indent(e.body))}` : ''}`);
  }
  console.log(lines.join('\n'));
}

async function cmdFile(c, flags, actor) {
  const reporter = String(one(flags, 'reporter') || process.env.HELPDESK_REPORTER || '').toLowerCase();
  if (!reporter) throw new UsageError('--reporter <name@gsadus.com> is required (or set HELPDESK_REPORTER)');
  const product = oneOf(required(flags, 'product', PRODUCTS.join(' | ')), PRODUCTS, 'product');
  const category = oneOf(required(flags, 'category', CATEGORIES.join(' | ')), CATEGORIES, 'category');
  const severity = oneOf(required(flags, 'severity', SEVERITIES.join(' | ')), SEVERITIES, 'severity');
  const note = required(flags, 'note');
  const shots = all(flags, 'screenshot').map(readScreenshot);
  const n = await tx(c, actor, async () => {
    const { rows } = await c.query(
      `insert into helpdesk.ticket (reporter_email, note, page_url, product, category, severity)
       values ($1, $2, $3, $4, $5, $6) returning number`,
      [reporter, note, one(flags, 'url') ?? null, product, category, severity]);
    const number = rows[0].number;
    for (const s of shots) {
      await c.query('insert into helpdesk.attachment (ticket, content_type, filename, byte_size, data) values ($1, $2, $3, $4, $5)',
        [number, s.type, s.name, s.data.length, s.data]);
    }
    await addEvent(c, number, actor, 'comment', `Filed through the helpdesk command by ${actor}.`);
    return number;
  });
  console.log(`${hd(n)} filed (${product}/${category}, ${severity})${shots.length ? ` with ${shots.length} screenshot(s)` : ''}`);
  return { n, actor, text: null };
}

async function cmdTriage(c, flags, n, actor) {
  const note = required(flags, 'note', 'affected code, related errors, repro steps, fix plan, size');
  const corrections = {};
  if (one(flags, 'product')) corrections.product = oneOf(one(flags, 'product'), PRODUCTS, 'product');
  if (one(flags, 'category')) corrections.category = oneOf(one(flags, 'category'), CATEGORIES, 'category');
  if (one(flags, 'severity')) corrections.severity = oneOf(one(flags, 'severity'), SEVERITIES, 'severity');
  const moved = await tx(c, actor, async () => {
    const t = await load(c, n);
    if (t.status === 'triaged') await update(c, n, { triage_note: note, ...corrections });
    // queued → triaged: the QUEUE item was demoted to parkinglot at a handoff close.
    else await transition(c, n, ['new', 'needs_info', 'reopened', 'queued'], 'triaged', { triage_note: note, ...corrections });
    await addEvent(c, n, actor, 'triage', note);
    return t.status !== 'triaged';
  });
  console.log(`${hd(n)} triaged`);
  return moved ? { n, actor, text: say.triaged() } : null;
}

async function cmdAsk(c, flags, n, actor) {
  const question = required(flags, 'question');
  await tx(c, actor, async () => {
    const t = await load(c, n);
    if (t.status === 'in_progress') assertWorker(t, actor);
    await transition(c, n, ['new', 'triaged', 'reopened', 'in_progress'], 'needs_info');
    await addEvent(c, n, actor, 'comment', `Question for the reporter: ${question}`);
  });
  console.log(`${hd(n)} needs info from the reporter`);
  return { n, actor, text: say.asked(question) };
}

const OWNER_DECIDES = ['new', 'triaged', 'needs_info', 'reopened'];

function describe(t, extra) {
  return [`${t.product} · ${t.category} · ${t.severity} · from ${t.reporter_email}`, '', "Reporter's note:", t.note,
    ...(t.triage_note ? ['', 'Triage:', t.triage_note] : []), ...(extra ? ['', extra] : [])].join('\n');
}

async function cmdApprove(c, flags, n, url) {
  const to = oneOf(required(flags, 'to', ASSIGNEES.join(' | ')), ASSIGNEES, 'to');
  const note = one(flags, 'note');
  const t = await load(c, n);
  const from = [...OWNER_DECIDES, 'queued', 'approved'];
  if (!from.includes(t.status)) throw new UsageError(`${hd(n)} is ${t.status}; only unworked tickets can be approved`);
  confirmOrThrow('Approve', `Approve ${hd(n)} for ${to}?\n\n${describe(t, note && `Your note: ${note}`)}`, url);
  const owner = `owner@${os.hostname().toLowerCase()}`;
  await tx(c, owner, async () => {
    if (t.status === 'approved') {
      await transition(c, n, ['approved'], 'approved', { assignee: to });
      await addEvent(c, n, owner, 'comment', `Reassigned to ${to}.`);
    } else {
      await transition(c, n, [t.status], 'approved', { assignee: to });
    }
    if (note) await addEvent(c, n, owner, 'comment', note);
  });
  console.log(`${hd(n)} approved for ${to}`);
  return { n, actor: owner, text: t.status === 'approved' ? say.reassigned(to) : say.approved(to) };
}

async function cmdReject(c, flags, n, url) {
  const reason = required(flags, 'reason');
  const t = await load(c, n);
  if (!OWNER_DECIDES.includes(t.status)) throw new UsageError(`${hd(n)} is ${t.status}; only undecided tickets can be rejected`);
  confirmOrThrow('Reject', `Reject ${hd(n)}?\n\nReason: ${reason}\n\n${describe(t)}`, url);
  const owner = `owner@${os.hostname().toLowerCase()}`;
  await tx(c, owner, () => transition(c, n, [t.status], 'rejected', { resolution: reason }));
  console.log(`${hd(n)} rejected`);
  return { n, actor: owner, text: say.rejected(reason) };
}

async function cmdDup(c, flags, n, actor) {
  const of = parseTicket(required(flags, 'of', 'the original ticket'));
  if (of === n) throw new UsageError('a ticket cannot duplicate itself');
  await tx(c, actor, async () => {
    await load(c, of);
    await transition(c, n, OWNER_DECIDES, 'duplicate', { duplicate_of: of });
    await addEvent(c, n, actor, 'comment', `Duplicate of ${hd(of)}.`);
  });
  console.log(`${hd(n)} marked duplicate of ${hd(of)}`);
  return { n, actor, text: say.duplicate(of) };
}

async function cmdPark(c, flags, n, actor) {
  const ref = required(flags, 'ref', 'the QUEUE.md entry, e.g. PM/docs/QUEUE.md');
  const note = one(flags, 'note');
  await tx(c, actor, async () => {
    const t = await load(c, n);
    if (t.status === 'in_progress') assertWorker(t, actor);
    await transition(c, n, ['new', 'triaged', 'approved', 'in_progress'], 'queued');
    await addLinks(c, n, actor, [{ kind: 'queue', ref }]);
    if (note) await addEvent(c, n, actor, 'comment', note);
  });
  console.log(`${hd(n)} parked in ${ref}`);
  return { n, actor, text: say.queued(ref) };
}

async function cmdClaim(c, n, actor) {
  const t = await load(c, n);
  if (t.assignee === 'owner' && !isOwner(actor)) throw new UsageError(`${hd(n)} is assigned to the owner`);
  if (t.assignee && t.assignee !== 'owner' && !isOwner(actor) && harnessOf(actor) !== t.assignee) {
    throw new UsageError(`${hd(n)} is assigned to ${t.assignee}; claim it with --as ${t.assignee}:<AgentName>`);
  }
  await tx(c, actor, async () => {
    const { rows } = await c.query(
      `update helpdesk.ticket set status = 'in_progress', claimed_by = $2, claimed_at = now()
        where number = $1 and status = 'approved' returning number`, [n, actor]);
    if (!rows.length) {
      const now = await load(c, n);
      throw new UsageError(`${hd(n)} is ${now.status}${now.claimed_by ? `, claimed by ${now.claimed_by}` : ''}; ` +
        'only approved, unclaimed tickets can be claimed');
    }
  });
  console.log(`${hd(n)} claimed by ${actor}`);
  return { n, actor, text: say.claimed(harnessOf(actor)) };
}

async function cmdRelease(c, flags, n, actor) {
  await tx(c, actor, async () => {
    assertWorker(await load(c, n), actor);
    await transition(c, n, ['in_progress'], 'approved');
    if (one(flags, 'note')) await addEvent(c, n, actor, 'comment', one(flags, 'note'));
  });
  console.log(`${hd(n)} released back to approved`);
  return { n, actor, text: say.released() };
}

// Review sent it back: fix_ready → in_progress, the same claim continues.
async function cmdRework(c, flags, n, actor) {
  const note = required(flags, 'note', 'what the review found');
  await tx(c, actor, async () => {
    assertWorker(await load(c, n), actor);
    await transition(c, n, ['fix_ready'], 'in_progress');
    await addEvent(c, n, actor, 'comment', note);
  });
  console.log(`${hd(n)} back in progress`);
  return { n, actor, text: say.rework() };
}

async function cmdFixReady(c, flags, n, actor) {
  const summary = required(flags, 'summary', 'what changed and how it was verified');
  await tx(c, actor, async () => {
    assertWorker(await load(c, n), actor);
    await transition(c, n, ['in_progress'], 'fix_ready');
    await addEvent(c, n, actor, 'comment', summary);
    await addLinks(c, n, actor, linksFrom(flags));
  });
  console.log(`${hd(n)} fix ready for review`);
  return { n, actor, text: say.fixReady() };
}

async function cmdResolve(c, flags, n, actor) {
  const resolution = required(flags, 'resolution');
  await tx(c, actor, async () => {
    const t = await load(c, n);
    if (t.status === 'fix_ready') assertWorker(t, actor);
    await transition(c, n, ['fix_ready', 'queued'], 'resolved', { resolution });
    await addLinks(c, n, actor, linksFrom(flags));
  });
  console.log(`${hd(n)} resolved`);
  return { n, actor, text: say.resolved(resolution) };
}

async function cmdReopen(c, flags, n, actor) {
  const note = required(flags, 'note', 'why it is not fixed');
  await tx(c, actor, async () => {
    await transition(c, n, CLOSED, 'reopened');
    await addEvent(c, n, actor, 'comment', note);
  });
  console.log(`${hd(n)} reopened`);
  return { n, actor, text: say.reopened(note) };
}

async function cmdComment(c, flags, positional, n, actor) {
  const body = one(flags, 'body') ?? positional.slice(1).join(' ');
  if (!String(body).trim()) throw new UsageError('comment text is required');
  await tx(c, actor, async () => { await load(c, n); await addEvent(c, n, actor, 'comment', String(body)); });
  console.log(`${hd(n)} comment added`);
}

// A message to the reporter and staff in the ticket's thread, posted as the bot. Unlike the
// status lines, the post is the point, so it must succeed before the history records it.
async function cmdReply(c, flags, positional, n, actor, url) {
  const text = String(one(flags, 'body') ?? positional.slice(1).join(' ')).trim();
  if (!text) throw new UsageError('reply text is required');
  await load(c, n);
  const cfg = chatConfig(process.env, isLoopback(url));
  if (!cfg) throw new UsageError('Chat posting is off (HELPDESK_CHAT=off, or a test database without Chat settings)');
  await postUpdate(c, cfg, n, actor, say.reply(harnessOf(actor), text));
  await tx(c, actor, () => addEvent(c, n, actor, 'comment', `Posted to the Chat thread: ${text}`));
  console.log(`${hd(n)} reply posted to the Chat thread`);
}

async function cmdLink(c, flags, n, actor) {
  const kind = oneOf(required(flags, 'kind', 'pr | commit | queue | other'), ['pr', 'commit', 'queue', 'other'], 'kind');
  const ref = required(flags, 'ref');
  await tx(c, actor, async () => { await load(c, n); await addLinks(c, n, actor, [{ kind, ref }]); });
  console.log(`${hd(n)} linked ${kind}: ${ref}`);
}

const HELP = `helpdesk: staff tickets (HD-n) for GSADUs tools. Rules: ${SKILL}

Read
  status                          one-line summary of open work
  list [--all] [--status s1,s2] [--product p] [--json]
  show <n> [--json] [--no-files]  ticket, history, screenshots saved to %TEMP%\\helpdesk\\HD-n

Owner decisions (a confirmation window on the owner's desktop must be clicked)
  approve <n> --to claude|codex|owner [--note text]
  reject <n> --reason text

Triage and work (--as claude:<AgentName> | codex:<AgentName> | owner is required)
  triage <n> --note text [--product p] [--category c] [--severity s]
  ask <n> --question text          needs_info; the question goes to the history
  dup <n> --of <m>
  park <n> --ref <repo>/docs/QUEUE.md [--note text]
  claim <n>                        approved → in_progress, one session only
  release <n> [--note text]        give a claimed ticket back (in_progress → approved)
  fix-ready <n> --summary text [--pr ref] [--commit sha]
  rework <n> --note text           review sent it back (fix_ready → in_progress)
  resolve <n> --resolution text [--commit sha] [--pr ref]
  reopen <n> --note text
  comment <n> <text>               history only (for the owner and agents)
  reply <n> <text>                 posted to the ticket's Chat thread as the bot, and recorded
  link <n> --kind pr|commit|queue|other --ref ref

Chat: every change a reporter should see is posted to the ticket's thread in the Tech Requests
space. Exit code 3 means the change is saved but that post failed. HELPDESK_CHAT=off skips posts.

Filing (on a reporter's behalf; also needs --as)
  file --reporter name@gsadus.com --product ${PRODUCTS.join('|')}
       --category ${CATEGORIES.join('|')} --severity ${SEVERITIES.join('|')}
       --note text [--url page] [--screenshot file.png ...]

Hook
  status --hook claude|codex      SessionStart context line (reads hook JSON on stdin)`;

export async function main(argv) {
  const [command, ...rest] = argv;
  if (!command || command === 'help' || command === '--help' || command === '-h') { console.log(HELP); return; }
  const { positional, flags } = parseArgs(rest);
  if (command === 'status' && one(flags, 'hook')) { await hookStatus(); return; }
  const handlers = {
    status: (c) => cmdStatus(c),
    list: (c) => cmdList(c, flags),
    show: (c) => cmdShow(c, flags, parseTicket(positional[0])),
    file: (c) => cmdFile(c, flags, actorFrom(flags)),
    triage: (c) => cmdTriage(c, flags, parseTicket(positional[0]), actorFrom(flags)),
    ask: (c) => cmdAsk(c, flags, parseTicket(positional[0]), actorFrom(flags)),
    approve: (c, url) => cmdApprove(c, flags, parseTicket(positional[0]), url),
    reject: (c, url) => cmdReject(c, flags, parseTicket(positional[0]), url),
    dup: (c) => cmdDup(c, flags, parseTicket(positional[0]), actorFrom(flags)),
    park: (c) => cmdPark(c, flags, parseTicket(positional[0]), actorFrom(flags)),
    claim: (c) => cmdClaim(c, parseTicket(positional[0]), actorFrom(flags)),
    release: (c) => cmdRelease(c, flags, parseTicket(positional[0]), actorFrom(flags)),
    'fix-ready': (c) => cmdFixReady(c, flags, parseTicket(positional[0]), actorFrom(flags)),
    rework: (c) => cmdRework(c, flags, parseTicket(positional[0]), actorFrom(flags)),
    resolve: (c) => cmdResolve(c, flags, parseTicket(positional[0]), actorFrom(flags)),
    reopen: (c) => cmdReopen(c, flags, parseTicket(positional[0]), actorFrom(flags)),
    comment: (c) => cmdComment(c, flags, positional, parseTicket(positional[0]), actorFrom(flags)),
    reply: (c, url) => cmdReply(c, flags, positional, parseTicket(positional[0]), actorFrom(flags), url),
    link: (c) => cmdLink(c, flags, parseTicket(positional[0]), actorFrom(flags)),
  };
  const handler = handlers[command];
  if (!handler) throw new UsageError(`unknown command ${command}; see \`helpdesk help\``);
  const url = resolveDsn();
  const client = await connect(url);
  try {
    const news = await handler(client, url);
    if (news) await announce(client, url, news);
  } finally { await client.end().catch(() => {}); }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main(process.argv.slice(2)).catch((e) => {
    if (e instanceof UsageError) { console.error(e.message); process.exitCode = 2; return; }
    // P0001 = the schema's own guards (transitions, immutability); 23514 = check constraints; 42501 = privileges.
    const refused = ['P0001', '23514', '42501', '23503'].includes(e.code);
    console.error(refused ? `refused by the database: ${e.message}` : `helpdesk: ${e.message}`);
    process.exitCode = 1;
  });
}
