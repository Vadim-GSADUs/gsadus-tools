// Posts ticket updates to the helpdesk's Google Chat space as the "GSADUs staff" Chat app (PM's
// bot). Spec: Vault wiki/curated/helpdesk.md → Chat. One thread per ticket, keyed by HD-n, so PM
// (which opens the thread when a ticket is submitted) and this command reply into the same one.
//
// Auth (never printed): the pm-chat-app@ service-account key GOOGLE_CHAT_APP_SA_JSON and the
// space name HELPDESK_CHAT_SPACE, from the environment if both are set, otherwise read at call
// time from Doppler core/prd. The app authenticates as itself with the chat.bot scope and can
// post only in spaces it is a member of. Posts are text only; screenshots never go to Chat.
import { spawnSync } from 'node:child_process';
import crypto from 'node:crypto';

const SCOPE = 'https://www.googleapis.com/auth/chat.bot';
const GOOGLE = { api: 'https://chat.googleapis.com', token: 'https://oauth2.googleapis.com/token' };
const NAMES = ['GOOGLE_CHAT_APP_SA_JSON', 'HELPDESK_CHAT_SPACE'];
const WHO = { claude: 'Claude', codex: 'Codex', owner: 'The owner' };
const LOOPBACK = new Set(['127.0.0.1', 'localhost', '::1', '[::1]']);

const oneLine = (s, max) => {
  const t = String(s ?? '').replace(/\s+/g, ' ').trim();
  return t.length > max ? `${t.slice(0, max - 1)}…` : t;
};
const nameOf = (harness) => WHO[harness] || harness;
const isLoopbackUrl = (u) => { try { return LOOPBACK.has(new URL(u).hostname); } catch { return false; } };

// Where to post, or null when posting is off. The endpoint overrides are honored only when they
// point at loopback, and a loopback (test) database posts only to such a local stand-in, so a
// test can never reach the real space.
export function chatConfig(env = process.env, loopbackDb = false) {
  if (env.HELPDESK_CHAT === 'off') return null;
  const override = (name) => (isLoopbackUrl(env[name]) ? env[name].replace(/\/$/, '') : null);
  const api = override('HELPDESK_CHAT_API');
  const tokenUri = override('HELPDESK_CHAT_TOKEN_URI');
  if (loopbackDb && !(api && tokenUri)) return null;
  let values;
  if (env.GOOGLE_CHAT_APP_SA_JSON && env.HELPDESK_CHAT_SPACE) {
    values = { key: env.GOOGLE_CHAT_APP_SA_JSON, space: env.HELPDESK_CHAT_SPACE };
  } else {
    const r = spawnSync('doppler', ['secrets', 'get', ...NAMES, '--project', 'core', '--config', 'prd', '--json'],
      { encoding: 'utf8', timeout: 15000, windowsHide: true });
    if (r.error || r.status !== 0) throw new Error('the Chat settings could not be read from Doppler (core/prd); run `doppler login`');
    const got = JSON.parse(r.stdout);
    values = { key: got.GOOGLE_CHAT_APP_SA_JSON?.computed, space: got.HELPDESK_CHAT_SPACE?.computed };
  }
  if (!values.key || !/^spaces\/[^/]+$/.test(values.space || '')) throw new Error('the Chat settings are incomplete');
  return { sa: JSON.parse(values.key), space: values.space, api: api || GOOGLE.api, tokenUri: tokenUri || GOOGLE.token };
}

// -- Messages (Chat text: *bold*, one line each where possible) --------------------------------
export function openingText(t) {
  return [
    `*HD-${t.number}* · ${t.product} · ${t.category} · ${t.severity}`,
    oneLine(t.note, 300),
    `Reported by ${t.reporter_email}${t.page_url ? ` on ${t.page_url}` : ''}`,
  ].join('\n');
}

// Staff-facing line for each change. Triage notes and fix summaries stay in the history: they
// are written for the owner, not the reporter.
export const say = {
  triaged: () => 'Triaged; waiting for the owner to decide.',
  asked: (question) => `Question for the reporter: ${question}`,
  approved: (assignee) => (assignee === 'owner' ? 'Approved; the owner will handle it.' : `Approved and assigned to ${nameOf(assignee)}.`),
  reassigned: (assignee) => `Reassigned to ${nameOf(assignee)}.`,
  rejected: (reason) => `Closed without a change: ${reason}`,
  duplicate: (of) => `Duplicate of HD-${of}; follow that thread.`,
  queued: (ref) => `Parked for a larger work session (${String(ref).split(/[\\/]/)[0]} backlog).`,
  claimed: (harness) => `Work started (${nameOf(harness)}).`,
  released: () => 'Work paused; the ticket is back in the queue.',
  fixReady: () => 'A fix is ready for the owner to review.',
  rework: () => 'The review asked for more work.',
  resolved: (resolution) => `Resolved: ${resolution}`,
  reopened: (note) => `Reopened: ${note}`,
  reply: (harness, text) => `${nameOf(harness)}: ${text}`,
};

// -- Google APIs ----------------------------------------------------------------------------------
let cachedToken = null;

async function accessToken(cfg) {
  if (cachedToken && cachedToken.exp > Date.now() + 60000) return cachedToken.value;
  const now = Math.floor(Date.now() / 1000);
  const part = (o) => Buffer.from(JSON.stringify(o)).toString('base64url');
  const unsigned = `${part({ alg: 'RS256', typ: 'JWT', kid: cfg.sa.private_key_id })}.${part({
    iss: cfg.sa.client_email, scope: SCOPE, aud: GOOGLE.token, iat: now, exp: now + 3600,
  })}`;
  const signature = crypto.sign('RSA-SHA256', Buffer.from(unsigned), cfg.sa.private_key).toString('base64url');
  const res = await fetch(cfg.tokenUri, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: `${unsigned}.${signature}` }),
    signal: AbortSignal.timeout(15000),
  });
  const body = await res.json().catch(() => ({}));
  if (!res.ok || !body.access_token) throw new Error(`token ${res.status}: ${body.error_description || body.error || res.statusText}`);
  cachedToken = { value: body.access_token, exp: Date.now() + (body.expires_in || 3600) * 1000 };
  return cachedToken.value;
}

// Replies into the thread with this key, starting it if it does not exist yet. Returns the thread name.
export async function postMessage(cfg, threadKey, text) {
  const token = await accessToken(cfg);
  const res = await fetch(`${cfg.api}/v1/${cfg.space}/messages?messageReplyOption=REPLY_MESSAGE_FALLBACK_TO_NEW_THREAD`, {
    method: 'POST',
    headers: { authorization: `Bearer ${token}`, 'content-type': 'application/json; charset=utf-8' },
    body: JSON.stringify({ text, thread: { threadKey } }),
    signal: AbortSignal.timeout(15000),
  });
  const body = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(`Chat ${res.status}: ${body.error?.message || res.statusText}`);
  return body.thread?.name;
}

export const postToThread = (cfg, number, text) => postMessage(cfg, `HD-${number}`, text);
