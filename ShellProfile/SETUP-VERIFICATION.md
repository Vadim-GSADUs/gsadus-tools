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
Only our uniquely named test container was removed.

### Root cause (found 2026-09-25, after recurrences on 9/17, 9/24 and 9/25)

The sockets were never corrupt. The failures came from where Docker was started from.

- **Agents launched Docker inside the Claude and Codex desktop apps.** They ran
  `Start-Process "…\Docker Desktop.exe"` from their tool shells. Both apps are MSIX
  packages, and a program started from inside one joins that app's job and file-system
  container, even without package identity.
- **Inside that container, AF_UNIX socket files under `%LOCALAPPDATA%` cannot be opened,
  removed or renamed (Windows error 1920).** `fsutil reparsepoint query` on a live,
  healthy Docker socket fails from an agent shell. The same query from an
  Explorer-launched process returns `IO_REPARSE_TAG_AF_UNIX` (`0x80000023`), including on
  the quarantined sockets. So Docker could not remove its sockets, either at stop or at
  the next start. A graceful `docker desktop stop` of a container-launched Docker left all
  three sockets behind.
- **The host app also kills Docker's whole process tree when the app restarts or
  updates:**
  - The Codex app restarted at 11:20 on 9/17 (a new app-process log starts at 11:20:33).
  - The Claude app auto-updated at 14:29:21 on 9/24 (AppX `TerminateApplications
    successful`).
  - Docker's backend and VM logs stop mid-stream at those moments. Windows logged no
    shutdown, sleep or crash.
  - The backend's uptime counter (`m=+…` in `GET /time`) dates each killed instance to
    the agent `Start-Process` that launched it: 9/15 23:27:39Z from Codex, 9/24 20:34:14Z
    from Claude.
- **New AppData folders created inside the container land in the app's private package
  store.** The Secrets Engine directory lived under
  `%LOCALAPPDATA%\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Local\docker-secrets-engine`.
  A Codex copy holds the 9/15 quarantine. This is why WSL could not see it. WSL `rm` and
  renaming the parent directory worked because neither opens the socket through the
  container.

Outside the container, Docker handles its own sockets. Docker 4.67 was launched through
Explorer with the leftover sockets still in `Docker\run`. It replaced them (the backend
log says `removed stale socket file`) and started normally. A later graceful stop left
`Docker\run` and `docker-secrets-engine` empty, and the next start succeeded with no
manual step.

Upstream reports of this error remain open (docker/desktop-feedback
[#448](https://github.com/docker/desktop-feedback/issues/448),
[#532](https://github.com/docker/desktop-feedback/issues/532),
[#536](https://github.com/docker/desktop-feedback/issues/536),
[#631](https://github.com/docker/desktop-feedback/issues/631) and others). #631's native
reproducer shows sockets working under `%TEMP%` but failing elsewhere under
`%LOCALAPPDATA%`. That matches the container behaviour here, because MSIX AppData
redirection leaves `%TEMP%` alone. A Docker maintainer reports mitigations in 4.92.

### Fix

- **Start Docker with the helper:**
  `pwsh -NoProfile -File C:\GSADUs\Tools\ShellProfile\Start-DockerDesktop.ps1`.
  - It launches Docker through Explorer, outside every app container, and waits for
    the engine.
  - It reports a backend start failure from the log instead of retrying.
  - It warns when the running Docker was started from an agent shell. It detects this
    from a fresh Secrets Engine socket in a package `LocalCache`.
- **Stop Docker with `docker desktop stop`.** Never force-kill Docker processes, and
  never `wsl --shutdown` while Docker runs. Starting Docker yourself from the Start menu
  or taskbar is also outside the container.
- Both harnesses carry this rule in their machine-level instructions:
  `~\.claude\CLAUDE.md` and `~\.codex\AGENTS.md`.
- No Docker setting changed. Docker AI stays enabled; with the root cause removed,
  disabling it would change nothing. Start at sign-in stays off (owner decision
  2026-09-25).

If Docker fails to start again, first read the backend error. If the running Docker was
started from an agent shell, `docker desktop stop` it and start it with the helper;
Docker clears any leftover sockets itself. Parent-directory quarantine is only a last
resort, if a helper start still fails with 1920. Never delete Docker data or reset to
factory defaults.

### Darkroom stack removed (owner decision 2026-09-25)

The retired Darkroom local Supabase stack restarted with every Docker start. Removed:

- 12 `supabase_*_Darkroom` containers
- the named volumes `supabase_db_Darkroom` and `supabase_storage_Darkroom`
- `supabase_network_Darkroom`

Its 12 images were removed too. Kept:

- `postgres:17.11`, the test image
- three `public.ecr.aws/supabase/*` images the Darkroom containers did not use:
  `imgproxy:v3.8.0`, `realtime:v2.80.12` and `storage-api:v1.48.21`
- the 14 anonymous volumes (942 MB) that no Darkroom container used

Nothing was pruned.

### Verification (2026-09-25, Docker Desktop 4.67.0)

- **Explorer-launched Docker:** `docker version` reported server 29.3.1. The Secrets
  Engine socket was created in the real `%LOCALAPPDATA%` and is visible from WSL.
- **Clean stop/start:** a second full quit and relaunch through the helper also
  succeeded.
- **Postgres test:** `Test-DisposablePostgres.ps1` passed (transaction, rollback,
  graceful exit 0, test container removed).
- **Container detector:** it stays silent for the Explorer-launched Docker. Replayed
  against the 09:05 agent-launched instance, it names the Claude package.

### Update to 4.92.0 (owner, 2026-09-25)

The owner updated Docker Desktop in place from 4.67.0 to **4.92.0 (240144)**, engine
**29.8.0**, using the Update button.

- The updater's relaunch ran outside the app: the Secrets Engine socket is in the real
  `%LOCALAPPDATA%`, and the helper stayed silent.
- A full `docker desktop stop` and helper relaunch succeeded.
- `Test-DisposablePostgres.ps1` passed on 29.8.0.

When 4.92 finds a leftover socket at startup, it renames it with a `.stale` suffix
instead of failing. It even renames an existing `.stale` file again, so one became
`.stale.stale` on the next start. The only such leftover here came from the update's own
stop of 4.67. A clean stop of an Explorer-launched 4.92 left nothing. Any `.stale` file
in `Docker\run` is harmless, and can be removed from WSL.

### Leftovers removed (owner-approved, 2026-09-25)

Socket files return 1920 inside an agent shell, so the leftovers were deleted from WSL
after checking that each folder held only 0-byte files:

- **Claude copies:** `docker-secrets-engine` and
  `docker-secrets-engine.stale-20260925-090510` in
  `%LOCALAPPDATA%\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Local\`.
- **Codex copies:** `docker-secrets-engine` and
  `docker-secrets-engine.stale-20260915-162616` in
  `%LOCALAPPDATA%\Packages\OpenAI.Codex_2p2nqsd0c76g0\LocalCache\Local\`. The second is
  the 9/15 quarantine; its `engine.sock` dated from 2026-04-06.
- **Stale socket:** `Docker\run\userAnalyticsOtlpHttp.sock.stale.stale`.

The three `%LOCALAPPDATA%\Docker\run.stale-*` quarantine folders were already gone. They
disappeared during the 4.92 update: the `Docker` folder was modified at 09:32 and
everything else in it is intact. No log line records their removal.

## Rollback

Run `Install-WorktreeSetup.ps1 -Uninstall` to remove the managed Git hooks. Tools commits
can be reverted locally independently. Existing rendered env files and worktree
dependencies are not removed by uninstall. `Start-DockerDesktop.ps1` changes no Docker
state and can be deleted. Removing the Docker rule from the two machine-level instruction
files restores the old agent behaviour.
