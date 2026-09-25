# helpdesk — staff tickets (HD-n), worked by agents

The owner-side command for the GSADUs staff helpdesk. Staff file tickets about any GSADUs
tool (from PM's form, once that ships). This command is how the owner and agent sessions on
either PC read, triage, approve, claim, fix and close them.

- Spec and owner decisions: Vault `wiki/curated/helpdesk.md`.
- Rules for agents: `C:\GSADUs\.claude\skills\helpdesk\SKILL.md`.
- Schema: WebCatalog `pipeline/supabase/migrations/0113_helpdesk_init.sql`.

## Run

```powershell
helpdesk list                      # pwsh: shell-profile function (installs dependencies on first use)
node C:/GSADUs/Tools/Helpdesk/helpdesk.mjs list   # any shell, any cwd
helpdesk help                      # every command
```

The first run on a machine needs `npm ci --prefix C:\GSADUs\Tools\Helpdesk`. The `helpdesk`
function does this for you.

## How it connects

- It connects as the scoped `helpdesk_agent` role: select, insert and update on the
  `helpdesk` schema, never delete, and no access to any other schema.
- The DSN is `HELPDESK_DB_URL` in Doppler `core/prd`, read at call time and never printed or
  rendered to a file. An `HELPDESK_DB_URL` environment variable overrides it; the tests use
  that.
- TLS to the Supabase pooler is verified against `supabase-root-ca-2021.pem`. That is the same
  public root PM pins in `PM/lib/supabase-ca.ts`, and it expires 2031-04-26.

## Chat thread

Every change a reporter should see is posted to the ticket's thread in the **Tech Requests**
Google Chat space (`chat.mjs`), as the **GSADUs staff** bot: PM's Chat app, authenticated as
the `pm-chat-app@` service account with the `chat.bot` scope.
- One thread per ticket, keyed `HD-<n>`. The first post opens it with the ticket card, and
  `chat_thread` records its name.
- Triage notes, fix summaries and `comment` stay in the history; they're written for the owner.
  `reply <n> <text>` is the way to talk to the reporter.
- Screenshots are never posted.
- The key (`GOOGLE_CHAT_APP_SA_JSON`) and the space (`HELPDESK_CHAT_SPACE`) come from Doppler
  `core/prd` at call time, like the DSN.
- The database change is the record. If its post fails, the command says so and exits 3;
  `HELPDESK_CHAT=off` skips posting.

## What the database enforces, and what this command adds

The database enforces:
- which status moves are allowed;
- that history is append-only;
- that the reporter's submission never changes;
- that screenshots are purged, never deleted.

The command adds three things:
- **An actor on every change.** `--as claude:<AgentName>`, `codex:<AgentName>` or `owner` is
  required, and it is recorded as `<who>@<machine>`.
- **The claim guard.** An approved ticket goes to the harness the owner assigned. The first
  `claim` wins, and only that session, or the owner, moves it on.
- **Owner confirmation.** `approve` and `reject` open a topmost window on the owner's desktop.
  The command waits up to 4 minutes for a click; Cancel is the default. Agent permission
  rules match command text, so they can't guard this tool, while a window works in both
  harnesses and every permission mode. It stops accidental and unprompted approvals. It isn't
  a security boundary against an agent set on going around it (`confirm.ps1` explains). The
  hard boundary is the Chat Approve button checked against the owner's Google account, a
  later step.

## SessionStart status line

`helpdesk status --hook claude|codex` prints the SessionStart context line, such as
`Helpdesk: 2 new (1 blocking) · 1 to approve`. It stays silent when nothing is open or when
the session is outside `C:\GSADUs`. Install it per machine:

```powershell
node C:\GSADUs\.claude\hooks\helpdesk\install.mjs           # show the plan
node C:\GSADUs\.claude\hooks\helpdesk\install.mjs --apply   # write it (backups first)
```

Codex also needs the new hook trusted once, in its `/hooks` review.

## Tests

```powershell
npm test --prefix C:\GSADUs\Tools\Helpdesk     # unit tests; integration tests skip without a fixture
docker run -d --rm --name helpdesk-cli-pg -e POSTGRES_HOST_AUTH_METHOD=trust `
    -e POSTGRES_DB=helpdesk_cli_test -p 127.0.0.1:55444:5432 postgres:17.11
$env:HELPDESK_TEST_ADMIN_URL = 'postgresql://postgres@127.0.0.1:55444/helpdesk_cli_test'
npm test --prefix C:\GSADUs\Tools\Helpdesk     # now also drives the real CLI against migration 0113
docker stop helpdesk-cli-pg
```

- The integration suite refuses any target other than that fixture.
- It applies every helpdesk migration (`0113` onward) from the WebCatalog checkout next to
  this repo; `HELPDESK_MIGRATIONS` points it at another folder, such as a worktree.
- It answers the confirmation window through `HELPDESK_TEST_CONFIRM`, which only works against
  a loopback database.
- Its Chat test serves a stand-in for Google's token and Chat endpoints on loopback. A loopback
  database never posts anywhere else, so the tests can't reach the real space.
