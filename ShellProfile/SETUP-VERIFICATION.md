# Agent environment setup — 2026-09-15

Machine: `gsadus-vadim`, user `Vadim`. Tools changes are committed locally; no push,
deployment, live database mutation, or production role change was performed.

## Findings and permission decision

- Both harnesses run as the same Windows user. `doppler me` succeeded directly from
  Codex. The pwsh 7 profile shim already loads `C:\GSADUs\Tools\ShellProfile\profile.ps1`
  and restores `NODE_AUTH_TOKEN` from the existing managed user npm configuration.
- Codex's saved desktop permission selection is `full-access`, matching this task's
  effective unrestricted local access. Its config also contains an unelevated Windows
  sandbox setting for sandboxed modes, but that setting is not the cause of this failure.
- Claude's inspected global/workspace/repo settings contain read/npm command permissions
  and Agent Mail hooks, but no automatic Doppler or dependency setup for worktrees.
  Claude commonly works inside `C:\GSADUs` or `.claude/worktrees`; Codex's app creates
  worktrees under the user profile. Neither location inherits ignored files from Git.
- The old `pull-env` always targeted the four main checkouts. The missing seam was
  worktree routing and initialization, not a missing Doppler identity or authorization.
- **Permission changes: none.** No ACLs, Doppler scopes, Codex sandbox settings,
  Claude permissions, hook trust hashes, or production roles were changed. The new
  repo-local Git hooks are the only installed automatic execution configuration.

## Installed shared workflow

`Install-WorktreeSetup.ps1` installed `post-checkout` under `.git/hooks` in WebApp,
PM, WebCatalog and PNGTools. The hook invokes the canonical Tools script with pwsh 7,
only for initial linked-worktree checkout. Other hooks and ordinary branch switches
remain independent. No feature worktree was edited.

`pull-env -RepoPath <checkout>` resolves the repository through Git metadata and renders
only its approved Doppler config to that checkout. `init-worktree` preserves existing
env/dependencies and installs missing npm dependencies locally with `npm ci`.

## Verification

- Offline regression script passed: routing from a linked checkout with spaces in its
  path, main checkout preservation, missing-only behavior, download-failure preservation,
  refusal of unignored and tracked secret destinations.
- A fresh detached WebApp worktree at commit `35aa212` was created using plain
  `git worktree add`, with no environment overrides. The installed hook rendered its
  ignored `.env.local` from Doppler and installed private npm dependencies successfully.
- A fresh pwsh process launched by this Codex task restored npm auth, confirmed
  `node_modules` was a real local directory, reran `init-worktree` successfully, and ran
  **`npm run test:unit`: 126 files, 2,426 tests passed**. The checkout stayed Git-clean.
- Re-running the installer twice reported all four hooks already installed, without
  replacing them. All PowerShell files parsed successfully and `git diff --check` passed.
- The disposable PostgreSQL 17.11 test started with `--network none`, tmpfs storage,
  no published ports, and no application secrets. SQL insert/assert/rollback passed;
  PostgreSQL stopped with exit 0 and the test container was removed.
- Scope of harness verification: the installed Git hook and fresh Codex shell were
  exercised. A separate desktop conversation and a new Claude model session were not
  launched. Frontends that deliberately use `--no-checkout` or disable Git hooks need
  the documented `init-worktree` command after materializing the checkout.

The unit-test log is outside the repo at
`%LOCALAPPDATA%\Temp\gsadus-worktree-verification-20260915\unit-tests.log`.

## Docker repair

Docker Desktop 4.67.0 failed before starting its engine because stale AF_UNIX socket
entries could not be accessed (Windows error 1920). The first error named
`%LOCALAPPDATA%\Docker\run\dockerInference`; after that runtime directory was
quarantined, startup exposed the same error at
`%LOCALAPPDATA%\docker-secrets-engine\engine.sock`.

Individual socket reads/renames failed. With the backend stopped, the containing
runtime directories were renamed, preserving all entries. The Secrets Engine directory
was verified to contain only `engine.sock`. Both locations had to be clear in the same
startup: the failed intermediate attempt had recreated the first socket.

Preserved directories (local time in names):

- `%LOCALAPPDATA%\Docker\run.stale-20260915-162312`
- `%LOCALAPPDATA%\Docker\run.stale-20260915-162739`
- `%LOCALAPPDATA%\docker-secrets-engine.stale-20260915-162616`

Docker then started successfully, reporting Linux engine 29.3.1. No factory reset,
Docker settings edit, image/volume prune, WSL unregister, or disk-image change was used.
The pre-existing Darkroom Supabase containers remain present; Docker restarted them
according to their existing restart policies. Their unrelated vector restart loop and
exited edge-runtime container were not changed. Only our uniquely named test container
was removed. Docker is left running and the PostgreSQL image remains cached.

This repairs the observed stale runtime state; it is not an upstream Docker bug fix.
The same failure class is reported in Docker's
[issue tracker](https://github.com/docker/desktop-feedback/issues/448). If it recurs,
inspect the current error and stop Docker before quarantining only the affected
runtime directories. Do not delete Docker data or reset to factory defaults.

## Rollback

Run `Install-WorktreeSetup.ps1 -Uninstall` to remove the managed Git hooks. Tools commits
can be reverted locally independently. Existing rendered env files and worktree
dependencies are not removed by uninstall. Quarantined Docker socket directories remain
available for inspection; restoring stale sockets is unnecessary and would reintroduce
the startup failure.
