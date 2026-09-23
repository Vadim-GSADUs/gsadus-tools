# Direct Vercel deployment waiting

**Automatic since 2026-09-23:** a user-level Claude Code `PostToolUse` hook
(`C:/GSADUs/.claude/hooks/vercel-deploy/watch.mjs`, `asyncRewake`) runs this waiter after any
`git push` of `main` in WebApp or PM (worktrees included) and wakes the agent with the result,
so nobody has to remember it. One watcher per commit (a lock in the temp dir). Install or
repair on a machine: `node C:/GSADUs/.claude/hooks/vercel-deploy/install.mjs --apply`. The
manual form below stays for other harnesses and previews.

For both harnesses, after an **authorized** Git push, run from the checkout that
was pushed (including a worktree):

```powershell
$deployCommit = git rev-parse HEAD
pwsh -NoProfile -File C:/GSADUs/Tools/Vercel/Wait-Deployment.ps1 -Project WebApp -Commit $deployCommit
# PM: use -Project PM. Previews: add -Environment preview.
```

This GET-only helper uses the existing Vercel CLI login. It never deploys, promotes,
changes settings, or reads application secrets. Project/team names are explicit;
it works without a `.vercel/project.json` and does not need one in each worktree.
On a new machine install the Vercel CLI and run `vercel login`, then `vercel whoami`.
CLI 54.18.7 was verified on 2026-09-17; it already has all required flags.

The command discovers the **full pushed SHA**, newest matching deployment first,
every 3 seconds (plus CLI/API latency), for up to 120 seconds. Once found, it pins
that immutable URL and runs `vercel inspect --wait --timeout 600s --format json`.
An already Ready deployment returns immediately after its API requests. Readiness
polling after discovery is managed by Vercel CLI, not an agent's sleep loop.
Timeouts can be adjusted with `-DiscoveryTimeoutSeconds` / `-BuildTimeoutSeconds`.
Network requests may add time beyond these polling/build timeout windows.

Exit 0 means the requested deployment is READY; stdout includes its SHA, ID and URL.
Missing deployment, auth/API errors, ERROR, CANCELED, timeout or mismatched identity
exit 1. A canceled deployment is not silently replaced by another deployment; if a
redeploy is intentional, rerun the command to select it. A deployment skipped by
Git/ignored-build rules will time out during discovery, not report the old release.

Agents should launch this as a yielding foreground process, retain its process ID,
and consume completion promptly (short tool waits, at most 10 seconds when idle).
Do useful independent work while it runs. Do not add a fixed 60–120 second sleep,
wait for GitHub Actions to infer Vercel readiness, or keep checking build logs after
READY. CI checks and application smoke tests are separate gates: READY is Vercel's
deployment status, not proof that tests passed or that a production alias points to
this release. Where needed, inspect the public alias and compare deployment IDs.

If the immutable deployment URL/ID is already known, use the native command directly:

```powershell
vercel inspect <deployment-url-or-id> --scope vadim-7430s-projects --wait --timeout 600s --format json
```

Check `readyState` explicitly: command success alone must not turn CANCELED into
success. For a failure, retrieve `vercel inspect <deployment-url-or-id> --logs`.
Avoid passing `gsadus.vercel.app` or `gsadus-pm.vercel.app` as the deployment to wait
on: during a new build those aliases can still resolve to the previous Ready build.

## Investigation, 2026-09-17

CLI auth and direct status queries worked for both `gsadus` (WebApp) and `gsadus-pm`
in team `vadim-7430s-projects`. PM had a local project link; WebApp did not. No shared
agent deployment-wait implementation was found in the inspected skills/scripts.
No dashboard notification configuration is needed for this direct CLI workflow.
The historical agent delays were not timed; the change removes the need for coarse
agent polling and indirect GitHub status checks without claiming a measured speedup.

References: [inspect / --wait](https://vercel.com/docs/cli/inspect),
[list / metadata filters](https://vercel.com/docs/cli/list).

Validation: `pwsh -NoProfile -File Vercel/Test-WaitDeployment.ps1` covers ten
success/failure/discovery cases with a CLI double. Read-only production checks for
WebApp `994e3ab9` and PM `fd163c80` each returned READY in 3.7 seconds on 2026-09-17.
These were already Ready deployments; a live BUILDING-to-READY transition was not
triggered solely for testing.
