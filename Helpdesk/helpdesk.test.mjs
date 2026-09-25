// Tests for the helpdesk command. Unit tests always run. Integration tests drive the real CLI
// against a disposable PostgreSQL that has the WebCatalog helpdesk migrations applied, and skip unless
// HELPDESK_TEST_ADMIN_URL names that fixture:
//   docker run -d --rm --name helpdesk-cli-pg -e POSTGRES_HOST_AUTH_METHOD=trust \
//       -e POSTGRES_DB=helpdesk_cli_test -p 127.0.0.1:55444:5432 postgres:17.11
//   $env:HELPDESK_TEST_ADMIN_URL = 'postgresql://postgres@127.0.0.1:55444/helpdesk_cli_test'
//   npm test --prefix C:\GSADUs\Tools\Helpdesk
import assert from 'node:assert/strict';
import { spawn, spawnSync } from 'node:child_process';
import crypto from 'node:crypto';
import fs from 'node:fs';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import { after, before, describe, test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { chatConfig, openingText, say } from './chat.mjs';
import {
  actorFrom, parseArgs, parseTicket, summarize, testConfirmAllowed, UsageError,
} from './helpdesk.mjs';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const CLI = path.join(HERE, 'helpdesk.mjs');
// Every helpdesk migration (0113 onward), in order; HELPDESK_MIGRATIONS points elsewhere, e.g. a worktree.
const MIGRATIONS_DIR = process.env.HELPDESK_MIGRATIONS
  || path.resolve(HERE, '../../WebCatalog/pipeline/supabase/migrations');
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

  test('Chat messages: the opening card and staff-facing lines', () => {
    assert.equal(openingText({ number: 7, product: 'pm', category: 'broken', severity: 'blocking',
      note: 'Save button\n  does nothing', reporter_email: 'staff@gsadus.com', page_url: 'https://pm.test/x' }),
    '*HD-7* · pm · broken · blocking\nSave button does nothing\nReported by staff@gsadus.com on https://pm.test/x');
    assert.equal(say.approved('codex'), 'Approved and assigned to Codex.');
    assert.equal(say.approved('owner'), 'Approved; the owner will handle it.');
    assert.equal(say.queued('PM/docs/QUEUE.md'), 'Parked for a larger work session (PM backlog).');
    assert.equal(say.reply('claude', 'Which browser?'), 'Claude: Which browser?');
  });

  test('a test database never posts to the real Chat space', () => {
    const settings = { GOOGLE_CHAT_APP_SA_JSON: '{}', HELPDESK_CHAT_SPACE: 'spaces/X' };
    assert.equal(chatConfig(settings, true), null);
    assert.equal(chatConfig({ ...settings, HELPDESK_CHAT_API: 'https://chat.googleapis.com',
      HELPDESK_CHAT_TOKEN_URI: 'https://oauth2.googleapis.com/token' }, true), null);
    assert.equal(chatConfig({ ...settings, HELPDESK_CHAT: 'off' }, false), null);
    const local = chatConfig({ ...settings, HELPDESK_CHAT_API: 'http://127.0.0.1:9/', HELPDESK_CHAT_TOKEN_URI: 'http://127.0.0.1:9/token' }, true);
    assert.equal(local.api, 'http://127.0.0.1:9');
    assert.equal(chatConfig({ ...settings, HELPDESK_CHAT_API: 'https://evil.test' }, false).api, 'https://chat.googleapis.com');
    assert.throws(() => chatConfig({ ...settings, HELPDESK_CHAT_SPACE: 'spaces/X/threads/y' }, false), /incomplete/);
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
    const migrations = fs.readdirSync(MIGRATIONS_DIR).filter((f) => /^\d{4}_helpdesk_.*\.sql$/.test(f)).sort();
    assert.ok(migrations.includes('0113_helpdesk_init.sql'), `no 0113 in ${MIGRATIONS_DIR}`);
    for (const f of migrations) await admin.query(fs.readFileSync(path.join(MIGRATIONS_DIR, f), 'utf8'));
    // What provision_service_roles.py does for this role (--read-only-schemas helpdesk).
    await admin.query(`alter role helpdesk_agent with login bypassrls password 'fixture-only';
      grant select on all tables in schema helpdesk to helpdesk_agent`);
    agentUrl = `postgresql://helpdesk_agent:fixture-only@127.0.0.1:${target.port}${target.pathname}`;
  });

  after(async () => { await admin?.end(); });

  function childEnv(confirm, env) {
    const e = { ...process.env, HELPDESK_DB_URL: agentUrl, HELPDESK_TEST_CONFIRM: confirm };
    for (const name of ['HELPDESK_ACTOR', 'HELPDESK_CHAT', 'GOOGLE_CHAT_APP_SA_JSON', 'HELPDESK_CHAT_SPACE',
      'HELPDESK_CHAT_API', 'HELPDESK_CHAT_TOKEN_URI']) delete e[name];
    return { ...e, ...env };
  }
  function run(args, { confirm = 'yes', input, env = {} } = {}) {
    const r = spawnSync(process.execPath, [CLI, ...args], { env: childEnv(confirm, env), encoding: 'utf8', input });
    return { code: r.status, out: r.stdout, err: r.stderr };
  }
  // For tests that serve HTTP in this process: spawnSync would block the server.
  function runAsync(args, { confirm = 'yes', env = {} } = {}) {
    return new Promise((resolve) => {
      const child = spawn(process.execPath, [CLI, ...args], { env: childEnv(confirm, env) });
      let out = '';
      let err = '';
      child.stdout.on('data', (d) => { out += d; });
      child.stderr.on('data', (d) => { err += d; });
      child.on('close', (code) => resolve({ code, out, err }));
    });
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
    for (const product of ['crm', 'pyrevit']) {
      assert.match(run(['file', '--as', 'owner', '--reporter', 'a@gsadus.com', '--product', product, '--category', 'broken',
        '--severity', 'annoying', '--note', 'x']).err, /--product must be one of webapp \| pm \| it/);
    }
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

  test('Chat: opens the thread, posts each visible change, records the thread, survives a failed post', async () => {
    const posts = [];
    let failPosts = false;
    const server = http.createServer((req, res) => {
      let body = '';
      req.on('data', (d) => { body += d; });
      req.on('end', () => {
        const reply = (status, json) => { res.writeHead(status, { 'content-type': 'application/json' }); res.end(JSON.stringify(json)); };
        if (req.url === '/token') {
          const assertion = new URLSearchParams(body).get('assertion');
          const claims = JSON.parse(Buffer.from(assertion.split('.')[1], 'base64url'));
          reply(200, { access_token: `token-for-${claims.iss}-${claims.scope.split('/').pop()}`, expires_in: 3600 });
          return;
        }
        if (failPosts) { reply(503, { error: { message: 'Chat is down' } }); return; }
        const msg = JSON.parse(body);
        posts.push({ url: req.url, auth: req.headers.authorization, text: msg.text, key: msg.thread.threadKey });
        reply(200, { name: 'spaces/TEST/messages/m', thread: { name: `spaces/TEST/threads/t-${msg.thread.threadKey}` } });
      });
    });
    await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
    const base = `http://127.0.0.1:${server.address().port}`;
    const { privateKey } = crypto.generateKeyPairSync('rsa', { modulusLength: 2048,
      privateKeyEncoding: { type: 'pkcs8', format: 'pem' }, publicKeyEncoding: { type: 'spki', format: 'pem' } });
    const env = {
      GOOGLE_CHAT_APP_SA_JSON: JSON.stringify({ client_email: 'bot@fixture', private_key: privateKey, private_key_id: 'k1' }),
      HELPDESK_CHAT_SPACE: 'spaces/TEST', HELPDESK_CHAT_API: base, HELPDESK_CHAT_TOKEN_URI: `${base}/token`,
    };
    try {
      const go = async (args) => {
        const r = await runAsync(args, { env });
        assert.equal(r.code, 0, `helpdesk ${args.join(' ')} failed: ${r.err}`);
        return r.out;
      };
      const n = Number(/HD-(\d+) filed/.exec(await go(['file', '--as', 'owner', '--reporter', 'staff@gsadus.com',
        '--product', 'pm', '--category', 'confusing', '--severity', 'annoying', '--note', 'Where is the save button?']))[1]);
      await go(['triage', String(n), '--as', 'claude:T', '--note', 'Internal: SaveBar hidden under 900px']);
      await go(['triage', String(n), '--as', 'claude:T', '--note', 'Refined internal note']);
      await go(['comment', String(n), '--as', 'claude:T', 'Internal only']);
      await go(['reply', String(n), '--as', 'claude:T', 'Which screen size are you on?']);
      await go(['approve', String(n), '--to', 'codex']);
      assert.deepEqual(posts.map((p) => p.text), [
        `*HD-${n}* · pm · confusing · annoying\nWhere is the save button?\nReported by staff@gsadus.com`,
        'Triaged; waiting for the owner to decide.',
        'Claude: Which screen size are you on?',
        'Approved and assigned to Codex.',
      ]);
      assert.ok(posts.every((p) => p.key === `HD-${n}` && p.url === '/v1/spaces/TEST/messages?messageReplyOption=REPLY_MESSAGE_FALLBACK_TO_NEW_THREAD'));
      assert.ok(posts.every((p) => p.auth === 'Bearer token-for-bot@fixture-chat.bot'));
      assert.equal((await ticket(n)).chat_thread, `spaces/TEST/threads/t-HD-${n}`);
      const history = (await admin.query("select body from helpdesk.ticket_event where ticket = $1 and kind = 'comment' order by id", [n])).rows;
      assert.ok(history.some((e) => e.body === 'Posted to the Chat thread: Which screen size are you on?'));

      failPosts = true;
      const claimed = await runAsync(['claim', String(n), '--as', 'codex:Worker'], { env });
      assert.equal(claimed.code, 3);
      assert.match(claimed.err, /the change is saved, but posting it to the Chat thread failed \(Chat 503: Chat is down\)/);
      assert.equal((await ticket(n)).status, 'in_progress');
      const replied = await runAsync(['reply', String(n), '--as', 'codex:Worker', 'Looking now'], { env });
      assert.equal(replied.code, 1);
      const recorded = (await admin.query("select count(*)::int as n from helpdesk.ticket_event where body like '%Looking now%'")).rows[0].n;
      assert.equal(recorded, 0);
    } finally {
      server.close();
    }
  });
});
