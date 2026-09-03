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
| `sentry-probe` | Read-only Sentry issues/events for any project in the org — see `..\Sentry\README.md`. |

## Sync model

Real work lives on `main` (pushed/pulled normally). Each machine snapshots its
working tree to its own `wip/<hostname>` branch — only that machine writes that
ref, so a force-push can never clobber the other machine. `unwip` fast-forwards
`main` then adopts the newest *other* machine's `wip/*` branch as uncommitted
changes. The autostash is only dropped after its content is confirmed already
safe on this machine's own `wip/<host>` ref, so local work is never lost.
