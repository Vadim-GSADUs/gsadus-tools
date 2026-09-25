// Tests for the helpdesk command. Unit tests always run. Integration tests drive the real CLI
// against a disposable PostgreSQL that has WebCatalog migration 0113 applied, and skip unless
// HELPDESK_TEST_ADMIN_URL names that fixture:
//   docker run -d --rm --name helpdesk-cli-pg -e POSTGRES_HOST_AUTH_METHOD=trust \
//       -e POSTGRES_DB=helpdesk_cli_test -p 127.0.0.1:55444:5432 postgres:17.11
//   $env:HELPDESK_TEST_ADMIN_URL = 'postgresql://postgres@127.0.0.1:55444/helpdesk_cli_test'
//   npm test --prefix C:\GSADUs\Tools\Helpdesk
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { after, before, describe, test } from 'node:test';
import { fileURLToPath } from 'node:url';
import {
  actorFrom, parseArgs, parseTicket, summarize, testConfirmAllowed, UsageError,
} from './helpdesk.mjs';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const CLI = path.join(HERE, 'helpdesk.mjs');
const MIGRATION = path.resolve(HERE, '../../WebCatalog/pipeline/supabase/migrations/0113_helpdesk_init.sql');
const HOST = os.hostname().toLowerCase();

describe('unit', () => {
  test('parseArgs keeps repeated flags, booleans and = values', () => {
    const { positional, flags } = parseArgs(['14', '--screenshot', 'a.png', '--screenshot', 'b.png', '--json', '--to=codex']);
    assert.deepEqual(positional, ['14']);
    assert.deepEqual(flags.get('screenshot'), ['a.png', 'b.png']);
    assert.deepEqual(flags.get('json'), [true]);
    assert.deepEqual(flags.get('to'), ['codex']);
    assert.throws(() => parseArgs(['--note']), UsageError);
  });

  test('parseTicket accepts 14, HD-14 and hd14 only', () => {
    assert.equal(parseTicket('14'), 14);
    assert.equal(parseTicket('HD-14'), 14);
    assert.equal(parseTicket('hd14'), 14);
    for (const bad of [undefined, '', 'HD-', '14a', '-3']) assert.throws(() => parseTicket(bad), UsageError);
  });

  test('actorFrom requires an identity and appends the machine', () => {
    assert.throws(() => actorFrom(new Map(), {}), /--as is required/);
    assert.equal(actorFrom(new Map([['as', ['claude:RoseIbis']]]), {}, 'GSADUS-VADIM'), 'claude:RoseIbis@gsadus-vadim');
    assert.equal(actorFrom(new Map(), { HELPDESK_ACTOR: 'owner' }, 'vg-home'), 'owner@vg-home');
    assert.throws(() => actorFrom(new Map([['as', ['bad name']]]), {}), UsageError);
  });

  test('the test confirmation override only works against a loopback database', () => {
    assert.equal(testConfirmAllowed('postgresql://u:p@127.0.0.1:55444/x', { HELPDESK_TEST_CONFIRM: 'yes' }), true);
    assert.equal(testConfirmAllowed('postgresql://u:p@aws-1-us-west-1.pooler.supabase.com:5432/postgres',
      { HELPDESK_TEST_CONFIRM: 'yes' }), false);
    assert.equal(testConfirmAllowed('postgresql://u:p@127.0.0.1:55444/x', { HELPDESK_TEST_CONFIRM: 'maybe' }), false);
  });

  test('summarize names only what needs attention', () => {
    assert.equal(summarize([]), '');
    assert.equal(summarize([
      { status: 'new', n: 2, blocking: 1 }, { status: 'reopened', n: 1, blocking: 0 },
      { status: 'triaged', n: 1, blocking: 0 }, { status: 'fix_ready', n: 1, blocking: 0 },
    ]), 'Helpdesk: 3 new (1 blocking) · 1 to approve · 1 fix-ready for review');
  });
});

const ADMIN = process.env.HELPDESK_TEST_ADMIN_URL;

describe('integration (disposable PostgreSQL)', { skip: !ADMIN && 'HELPDESK_TEST_ADMIN_URL not set' }, () => {
  let pg;
  let admin;
  let agentUrl;

  before(async () => {
    const target = new URL(ADMIN);
    if (target.hostname !== '127.0.0.1' || target.port !== '55444' || target.pathname !== '/helpdesk_cli_test') {
      throw new Error('Only the disposable helpdesk CLI fixture is allowed');
    }
    pg = (await import('pg')).default;
    admin = new pg.Client({ connectionString: ADMIN });
    await admin.connect();
    await admin.query(`drop schema if exists helpdesk cascade; drop schema if exists auth cascade;
      create schema auth; create table auth.users (id uuid primary key)`);
    await admin.query(`do $$ declare r text; begin
      foreach r in array array['anon', 'authenticated', 'service_role'] loop
        if not exists (select 1 from pg_roles where rolname = r) then execute format('create role %I nologin', r); end if;
      end loop;
      foreach r in array array['pm_service', 'webapp_service'] loop
        if not exists (select 1 from pg_roles where rolname = r) then execute format('create role %I login bypassrls', r); end if;
      end loop;
      if exists (select 1 from pg_roles where rolname = 'helpdesk_agent') then
        execute 'drop owned by helpdesk_agent'; execute 'drop role helpdesk_agent';
      end if; end $$`);
    await admin.query(fs.readFileSync(MIGRATION, 'utf8'));
    // What provision_service_roles.py does for this role (--read-only-schemas helpdesk).
    await admin.query(`alter role helpdesk_agent with login bypassrls password 'fixture-only';
      grant select on all tables in schema helpdesk to helpdesk_agent`);
    agentUrl = `postgresql://helpdesk_agent:fixture-only@127.0.0.1:${target.port}${target.pathname}`;
  });

  after(async () => { await admin?.end(); });

  function run(args, { confirm = 'yes', input, env = {} } = {}) {
    const childEnv = { ...process.env, HELPDESK_DB_URL: agentUrl, HELPDESK_TEST_CONFIRM: confirm, ...env };
    delete childEnv.HELPDESK_ACTOR;
    const r = spawnSync(process.execPath, [CLI, ...args], { env: childEnv, encoding: 'utf8', input });
    return { code: r.status, out: r.stdout, err: r.stderr };
  }
  function ok(args, opts) {
    const r = run(args, opts);
    assert.equal(r.code, 0, `helpdesk ${args.join(' ')} failed: ${r.err}`);
    return r.out;
  }
  function file(note = 'The estimate total is blank', extra = []) {
    const out = ok(['file', '--as', 'owner', '--reporter', 'staff@gsadus.com', '--product', 'webapp', '--category',
      'broken', '--severity', 'annoying', '--note', note, ...extra]);
    return Number(/HD-(\d+) filed/.exec(out)[1]);
  }
  const ticket = async (n) => (await admin.query('select * from helpdesk.ticket where number = $1', [n])).rows[0];

  test('files a ticket with a screenshot and shows it back byte for byte', async () => {
    const png = Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), Buffer.alloc(64, 7)]);
    const shot = path.join(os.tmpdir(), `helpdesk-test-${process.pid}.png`);
    fs.writeFileSync(shot, png);
    const n = file('Chart legend overlaps the axis', ['--url', 'https://example.test/estimate', '--screenshot', shot]);
    const shown = JSON.parse(ok(['show', String(n), '--json']));
    assert.equal(shown.ticket.status, 'new');
    assert.equal(shown.ticket.page_url, 'https://example.test/estimate');
    assert.equal(shown.attachments.length, 1);
    assert.deepEqual(fs.readFileSync(shown.attachments[0].path), png);
    assert.deepEqual(shown.events.map((e) => [e.kind, e.actor]),
      [['created', 'staff@gsadus.com'], ['comment', `owner@${HOST}`]]);
    assert.match(ok(['show', `HD-${n}`]), /Chart legend overlaps the axis[\s\S]*Screenshots:/);
  });

  test('changes need an identity; bad input is refused before the database', () => {
    const n = file();
    assert.match(run(['triage', String(n), '--note', 'x']).err, /--as is required/);
    assert.match(run(['file', '--as', 'owner', '--reporter', 'a@gsadus.com', '--product', 'crm', '--category', 'broken',
      '--severity', 'annoying', '--note', 'x']).err, /--product must be one of/);
    const bogus = path.join(os.tmpdir(), `helpdesk-test-${process.pid}.txt`);
    fs.writeFileSync(bogus, 'not an image');
    assert.match(run(['file', '--as', 'owner', '--reporter', 'a@gsadus.com', '--product', 'pm', '--category', 'idea',
      '--severity', 'nice-to-have', '--note', 'x', '--screenshot', bogus]).err, /not a PNG, JPEG or WebP/);
  });

  test('full lifecycle: triage, owner approval, one claim, review, resolve', async () => {
    const n = file();
    ok(['triage', String(n), '--as', 'claude:Triager', '--note', 'EstimateTotal renders before rooms load', '--severity', 'blocking']);
    ok(['approve', String(n), '--to', 'codex', '--note', 'Go ahead']);
    assert.match(run(['claim', String(n), '--as', 'claude:Other']).err, /assigned to codex/);
    ok(['claim', String(n), '--as', 'codex:First']);
    assert.match(run(['claim', String(n), '--as', 'codex:Second']).err, /in_progress, claimed by codex:First@/);
    assert.match(run(['fix-ready', String(n), '--as', 'codex:Second', '--summary', 'x']).err, /claimed by codex:First@/);
    ok(['fix-ready', String(n), '--as', 'codex:First', '--summary', 'Guarded the empty room list; tests pass', '--pr', 'gsadus-web-app#12']);
    ok(['rework', String(n), '--as', 'owner', '--note', 'Also cover the PDF export']);
    ok(['fix-ready', String(n), '--as', 'codex:First', '--summary', 'PDF export covered too']);
    ok(['resolve', String(n), '--as', 'codex:First', '--resolution', 'Fixed in production', '--commit', 'abc1234']);
    const t = await ticket(n);
    assert.equal(t.status, 'resolved');
    assert.equal(t.severity, 'blocking');
    assert.equal(t.claimed_by, null);
    assert.deepEqual(t.links, [{ kind: 'pr', ref: 'gsadus-web-app#12' }, { kind: 'commit', ref: 'abc1234' }]);
    const moves = (await admin.query(
      "select actor, from_status, to_status from helpdesk.ticket_event where ticket = $1 and kind = 'status' order by id", [n])).rows;
    assert.deepEqual(moves.map((m) => `${m.from_status}>${m.to_status} ${m.actor}`), [
      `new>triaged claude:Triager@${HOST}`,
      `triaged>approved owner@${HOST}`,
      `approved>in_progress codex:First@${HOST}`,
      `in_progress>fix_ready codex:First@${HOST}`,
      `fix_ready>in_progress owner@${HOST}`,
      `in_progress>fix_ready codex:First@${HOST}`,
      `fix_ready>resolved codex:First@${HOST}`,
    ]);
  });

  test('a declined confirmation leaves the ticket untouched', async () => {
    const n = file();
    const r = run(['approve', String(n), '--to', 'claude'], { confirm: 'no' });
    assert.equal(r.code, 2);
    assert.match(r.err, /Approve cancelled by the owner/);
    assert.match(run(['reject', String(n), '--reason', 'x'], { confirm: 'no' }).err, /Reject cancelled/);
    assert.equal((await ticket(n)).status, 'new');
  });

  test('reject, reopen, duplicate, park and resolve from the queue', async () => {
    const a = file('One');
    const b = file('Two');
    const q = file('Three');
    ok(['reject', String(a), '--reason', 'Works as designed']);
    assert.equal((await ticket(a)).resolution, 'Works as designed');
    ok(['reopen', String(a), '--as', 'owner', '--note', 'Reporter showed a second case']);
    ok(['dup', String(b), '--as', 'claude:Triager', '--of', String(a)]);
    assert.equal((await ticket(b)).duplicate_of, a);
    ok(['park', String(q), '--as', 'claude:Triager', '--ref', 'PM/docs/QUEUE.md']);
    assert.equal((await ticket(q)).status, 'queued');
    ok(['resolve', String(q), '--as', 'claude:Shipper', '--resolution', 'Shipped with the QUEUE slice']);
    assert.equal((await ticket(q)).status, 'resolved');
    const demoted = file('Four');
    ok(['park', String(demoted), '--as', 'claude:Triager', '--ref', 'WebApp/docs/QUEUE.md']);
    ok(['triage', String(demoted), '--as', 'claude:Closer', '--note', 'QUEUE item demoted to parkinglot']);
    assert.equal((await ticket(demoted)).status, 'triaged');
    assert.match(run(['resolve', String(a), '--as', 'claude:X', '--resolution', 'x']).err,
      /is reopened; moving it to resolved needs fix_ready \| queued/);
  });

  test('status line and SessionStart hook', async () => {
    await admin.query('truncate helpdesk.attachment, helpdesk.ticket_event, helpdesk.ticket');
    assert.equal(ok(['status']).trim(), 'Helpdesk: nothing open');
    assert.equal(ok(['status', '--hook', 'claude'], { input: JSON.stringify({ cwd: 'C:\\GSADUs\\PM' }) }), '');
    const n = file();
    ok(['triage', String(file()), '--as', 'claude:T', '--note', 'x', '--severity', 'blocking']);
    assert.equal(ok(['status']).trim(), 'Helpdesk: 1 new · 1 to approve');
    const hook = JSON.parse(ok(['status', '--hook', 'codex'],
      { input: JSON.stringify({ cwd: 'C:\\GSADUs\\PM', hook_event_name: 'SessionStart' }) }));
    assert.equal(hook.hookSpecificOutput.hookEventName, 'SessionStart');
    assert.match(hook.hookSpecificOutput.additionalContext, /^Helpdesk: 1 new · 1 to approve\. .*SKILL\.md/);
    assert.equal(ok(['status', '--hook', 'claude'], { input: JSON.stringify({ cwd: 'D:\\Elsewhere' }) }), '');
    const broken = run(['status', '--hook', 'claude'], {
      input: JSON.stringify({ cwd: 'C:\\GSADUs' }), env: { HELPDESK_DB_URL: 'postgresql://nobody:x@127.0.0.1:1/none' },
    });
    assert.equal(broken.code, 0);
    assert.match(JSON.parse(broken.out).hookSpecificOutput.additionalContext, /^Helpdesk status unavailable/);
    assert.match(ok(['list']), new RegExp(`HD-${n}\\s+new`));
  });
});
