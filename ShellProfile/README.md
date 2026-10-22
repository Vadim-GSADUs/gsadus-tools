# ShellProfile

The GSADUs PowerShell profile — `wip` / `unwip` cross-machine sync logic, tracked
in git so both machines run identical code (no more manual `scp` and no drift).

## Files

| File | Purpose |
|---|---|
| `profile.ps1` | The real logic. Edit and commit this. |
| `Install-Profile.ps1` | Writes a tiny shim into the machine's `$PROFILE` that dot-sources `profile.ps1`. Run once per machine. |

## How it works

Each machine's `$PROFILE` is a 3-line shim pointing at
`C:\GSADUs\Tools\ShellProfile\profile.ps1`. Because the logic lives in the repo,
a change only needs to be committed once — every machine picks it up on the next
`unwip-all` (which pulls the Tools repo) and the next new shell. `$PROFILE` itself
never needs editing again.

## Setup on a new machine

```powershell
# after the Tools repo is cloned to C:\GSADUs\Tools
& C:\GSADUs\Tools\ShellProfile\Install-Profile.ps1
. $PROFILE        # load it into the current session
```

## Commands

| Command | Action |
|---|---|
| `wip` / `unwip` | Save / pick up working state for the current repo. |
| `wip-all` / `unwip-all` | Same across every GSADUs repo. |
| `end-day` | `wip-all` then lock the screen. |
| `Register-StartupUnwip` / `Unregister-StartupUnwip` | Add/remove the at-logon `unwip-all` scheduled task. |
| `pull-env` | Render every repo's env file from Doppler (values never echoed). |
| `pull-env -RepoPath .` | Refresh only this registered checkout's env, including linked worktrees. |
| `init-worktree` | Initialize this linked worktree's missing env and npm dependencies. |
| `sentry-probe` | Read-only Sentry issues/events for any project in the org — see `..\Sentry\README.md`. |
| `egress-probe` | Read-only Supabase egress guard for the shared project; `check` exits 2 over the limit — see `..\Supabase\README.md`. |
| `remote` | Moonlight/Sunshine streams between the two PCs: what is live, then a menu of what applies now — see [Streams between the PCs](#streams-between-the-pcs-remote). |

## Streams between the PCs (`remote`)

`remote` manages the Moonlight/Sunshine streams between VG-HOME and GSADUS-VADIM without an
agent. With no argument it prints what is live in both directions, then a numbered menu of
only the actions that apply now (Enter cancels). Each action is also a subcommand, and they
tab-complete:

| Command | Action |
|---|---|
| `remote status` | What is live, both ways. No menu. |
| `remote connect` | Stream the other PC's Desktop. Moonlight starts detached (through WMI, as Explorer would), so it prints nothing in the terminal and doesn't depend on it. |
| `remote disconnect` | End this PC's stream: a clean `Moonlight quit`, then close any Moonlight left here. |
| `remote kick` | End the other PC's stream into this PC: close its Moonlight over SSH, then wait up to 20 s for Sunshine here to drop the stream. |
| `remote restart` | Restart Sunshine on this PC (one UAC prompt). It ends any stream into this PC without SSH; the menu offers it when SSH is down. |

How it decides:

- **The other PC** is the other row of `$GSADUsStreamPCs`, picked by `$env:COMPUTERNAME`; there
  is no per-machine config. Moonlight gets the name it paired each PC under (`VG-Home`,
  `GSADUS-VADIM`).
- **A stream into a PC** is read from that PC's `C:\Program Files\Sunshine\config\sunshine.log`:
  the last `CLIENT CONNECTED` or `CLIENT DISCONNECTED` since the service started. **A stream
  from a PC** is its Moonlight process.
- **The other PC's half** comes from one SSH call (key auth, `BatchMode`, a 5 s connect timeout
  and a 10 s deadline; a healthy call takes about 1.5 s), so `remote` never hangs. With SSH down
  it says so and still shows this PC's half.
- **No credentials:** no stored password, and no Sunshine web API (it needs the owner's web
  login). SSH can't drive the other PC's desktop, so `kick` closes its Moonlight instead of
  quitting cleanly.
- **Over SSH**, `connect` refuses (a Moonlight started there has no desktop), and `disconnect`
  skips the quit and only closes Moonlight.

Moonlight is the portable copy at `%USERPROFILE%\Apps\Moonlight\Moonlight.exe` on each PC, and
Sunshine is a service on each. `setup.ps1` installs neither; their setup and settings are in
the Vault (`wiki/curated/remote-access.md` → Sunshine + Moonlight).

Offline checks, with mocks only (they start and end no stream):

```powershell
pwsh -NoProfile -File C:\GSADUs\Tools\ShellProfile\Test-RemoteCommand.ps1
```

## Sync model

Real work lives on `main` (pushed/pulled normally). Each machine snapshots its
working tree to its own `wip/<hostname>` branch — only that machine writes that
ref, so a force-push can never clobber the other machine. `unwip` fast-forwards
`main` then adopts the newest *other* machine's `wip/*` branch as uncommitted
changes. The autostash is only dropped after its content is confirmed already
safe on this machine's own `wip/<host>` ref, so local work is never lost.

## Agent worktrees (Claude and Codex)

Install once after updating Tools locally:

```powershell
pwsh -NoProfile -File C:\GSADUs\Tools\ShellProfile\Install-WorktreeSetup.ps1
```

This installs a small `post-checkout` hook in the shared `.git/hooks` of WebApp,
PM, WebCatalog and PNGTools. Both harnesses' ordinary `git worktree add` calls
run it automatically on the initial checkout. No repo-local `.codex` directory,
global `core.hooksPath`, harness hook trust change, or sandbox relaxation is needed.
Existing unrelated hooks and configured hook paths cause the installer to refuse
replacement. Reinstallation is idempotent; `-Uninstall` removes only managed hooks.

The hook resolves the main repo through Git's common directory, so worktree names
and locations do not matter. It uses the existing Doppler render table and authenticated
Windows user. It renders a missing env directly from Doppler and runs `npm ci` when a
root `package-lock.json` exists and `node_modules` is missing. Each worktree gets its own
dependencies. It neither copies secrets nor links the main checkout's dependencies.
Env writes must be gitignored and untracked; failed downloads leave existing files intact.
The existing PowerShell profile restores private npm authentication from the managed
user `.npmrc`. Enrollment is still `doppler login` and a normal `pull-env` once per machine.

Existing worktrees, env files, dependencies, and ordinary branch switches are left
alone. For a previously created or `--no-checkout` worktree, run:

```powershell
init-worktree
# Explicit refresh after a Doppler change:
pull-env -RepoPath .
```

From Git Bash or a shell without the profile:

```sh
pwsh -NoProfile -File C:/GSADUs/Tools/ShellProfile/Initialize-Worktree.ps1
```

A setup failure returns nonzero but Git has already created the checkout. Fix the
reported prerequisite and rerun `init-worktree`. An incomplete managed npm install
is marked in the worktree's Git metadata and retried; pre-existing dependencies are
preserved. After changing a lockfile, run `npm ci` yourself in that worktree. Python
virtual environments are not provisioned by this npm setup.

Git's [post-checkout contract](https://git-scm.com/docs/githooks#_post_checkout)
includes `git worktree add` unless `--no-checkout` is used. Codex's native
[local environments](https://learn.chatgpt.com/docs/environments/local-environment)
require a repo `.codex` folder, which this workspace explicitly disallows. The Git hook
supplies the shared automatic setup without changing that owner decision.

### Validation commands

```powershell
pwsh -NoProfile -File C:\GSADUs\Tools\ShellProfile\Test-WorktreeEnvironment.ps1
pwsh -NoProfile -File C:\GSADUs\Tools\ShellProfile\Test-DisposablePostgres.ps1
```

The first uses fake data and disposable repositories. The second starts PostgreSQL
17.11 with no network, published port, host mount or production credentials, checks a
transaction and rollback, verifies graceful exit 0, and removes only its own container.
It may download the official PostgreSQL image and leaves that image cached.

### Docker Desktop

Start Docker Desktop only with the helper, and stop it with `docker desktop stop`. A
Docker started from an agent shell runs inside the Claude or Codex app and breaks its
next start ([SETUP-VERIFICATION.md](SETUP-VERIFICATION.md) → Docker repair). The
helper launches Docker through Explorer and waits for the engine:

```powershell
pwsh -NoProfile -File C:\GSADUs\Tools\ShellProfile\Start-DockerDesktop.ps1
```

The local installation and Docker repair evidence is in [SETUP-VERIFICATION.md](SETUP-VERIFICATION.md).

## Desktop app containers (Claude and Codex)

The Claude and Codex desktop apps are MSIX packages. Agent tool shells, hooks and anything
they start run inside the app's container. The Claude app's Terminal panel, Task Scheduler,
Explorer and WSL run outside it (checked 2026-10-05). Inside a container, the user's
`AppData` is a merged view; `%TEMP%` and the profile outside `AppData` are not redirected:

- A new top-level folder created inside lands in the app's private copy,
  `%LOCALAPPDATA%\Packages\<package>\LocalCache\Roaming` or `…\Local`
  (`Claude_pzs8sxrjxfjjc`, `OpenAI.Codex_2p2nqsd0c76g0`).
- Inside, a file in the private copy hides the real file of the same name, and new files
  written to a folder that has a private copy go to the private copy.
- A folder that exists only in the real `AppData` is shared: reads, edits and new files from
  inside go to the real folder.
- Outside processes, and the other app, never see a private copy.

**CLI logins kept in files under `AppData`** (the Vercel CLI keeps its login in
`%APPDATA%\xdg.data\com.vercel.cli\auth.json`). Suppose that CLI first ran inside an app. A
later login from outside then updates only the real file, and that app's agent shells keep
the old login or none. Logins kept in the Windows Credential Manager (`gh`) are not
redirected. To share one login, log in from outside, then delete the app's private copy of
that CLI's folder. Delete its private parent folders too once they are empty, or the next
new folder under them lands in the private copy again. Agent shells then read and refresh
the real file. Copying the real files into the private copy works only until the next
login or token refresh on either side. List both packages' `LocalCache\Roaming` and
`LocalCache\Local` to see which tools hold a private copy. VG-HOME, 2026-10-05: the Claude
app's private `Roaming\xdg.data` (only the Vercel login) was deleted, and `vercel whoami`
from an agent shell then read the real login.
