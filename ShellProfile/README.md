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
