# Direct Vercel deployment waiting

**Automatic since 2026-09-23, waking reliably since 2026-10-14:** a user-level Claude Code
`PostToolUse` hook runs this waiter after any `git push` of `main` in WebApp or PM (worktrees
included) and wakes the agent with the result, so nobody has to remember it. Its triggers,
install (`node C:/GSADUs/.claude/hooks/vercel-deploy/install.mjs --apply`, per machine) and
dry test: [the hook's README](../../.claude/hooks/vercel-deploy/README.md). The manual form
below stays for other harnesses and previews.

For both harnesses, after an **authorized** Git push, run from the checkout that
was pushed (including a worktree):

```powershell
$deployCommit = git rev-parse HEAD
pwsh -NoProfile -File C:/GSADUs/Tools/Vercel/Wait-Deployment.ps1 -Project WebApp -Commit $deployCommit
# PM: use -Project PM. Previews: add -Environment preview.
```

`-Project` names the repo, and the waiter follows every Vercel project that repo's push
deploys: WebApp is two, `gsadus` (the App, repo root) and `gsadus-site` (the Site,
`apps/site`), each with its own `ignoreCommand`; PM is `gsadus-pm`.

This GET-only helper uses the existing Vercel CLI login. It never deploys, promotes,
changes settings, or reads application secrets. Project/team names are explicit;
it works without a `.vercel/project.json` and does not need one in each worktree.
On a new machine install the Vercel CLI and run `vercel login`, then `vercel whoami`.
Run `vercel whoami` from an agent shell as well. Agent shells in the Claude or Codex app may
read a private copy of the login, which a login from outside never updates
([ShellProfile README](../ShellProfile/README.md) → Desktop app containers).
CLI 54.18.7 was verified on 2026-09-17; it already has all required flags.

The command discovers each project's deployment of the **full pushed SHA**, newest first,
every 3 seconds (plus CLI/API latency), for up to 120 seconds. Once found, it pins each
immutable URL and runs `vercel inspect --wait --timeout <left>s --format json`, all of them
within one 600-second build budget. An already finished deployment returns immediately after
its API requests. Readiness polling after discovery is managed by Vercel CLI, not an agent's
sleep loop. Timeouts can be adjusted with `-DiscoveryTimeoutSeconds` / `-BuildTimeoutSeconds`.
Network requests may add time beyond these polling/build timeout windows.

Stdout is one JSON object: the commit and, per project, its deployment ID, URL and end state.

- **READY:** live.
- **CANCELED:** not built. Vercel creates a deployment for every push in every project and
  cancels it when the project's ignore step finds nothing to build (it does not tell that
  apart from a manual cancel). Production keeps its previous build. Not a failure.
- **Anything else** (ERROR, still BUILDING at the timeout): failed.

Exit 0 means no project failed; read each `state` to know what is live. Exit 1 means a
project failed (stderr names it with its `vercel inspect <url> --logs` command), or there was
no result at all: a project with no deployment for the SHA, auth/API errors, or a mismatched
identity. A canceled deployment is not silently replaced by another deployment; if a redeploy
is intentional, rerun the command to select it. `vercel inspect --wait` itself exits 1 on
ERROR and CANCELED while still printing the deployment, so the waiter reads that end state
rather than treating the exit as a CLI failure.

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

Check `readyState` explicitly: the exit code is 1 for ERROR and CANCELED alike. For a
failure, retrieve `vercel inspect <deployment-url-or-id> --logs`. Avoid passing
`gsadus.vercel.app`, `gsadus-site.vercel.app` or `gsadus-pm.vercel.app` as the deployment
to wait on: during a new build those aliases can still resolve to the previous Ready build.

## Investigation, 2026-09-17

CLI auth and direct status queries worked for both `gsadus` (WebApp) and `gsadus-pm`
in team `vadim-7430s-projects`. PM had a local project link; WebApp did not. No shared
agent deployment-wait implementation was found in the inspected skills/scripts.
No dashboard notification configuration is needed for this direct CLI workflow.
The historical agent delays were not timed; the change removes the need for coarse
agent polling and indirect GitHub status checks without claiming a measured speedup.

References: [inspect / --wait](https://vercel.com/docs/cli/inspect),
[list / metadata filters](https://vercel.com/docs/cli/list).

Validation: `pwsh -NoProfile -File Vercel/Test-WaitDeployment.ps1` covers fourteen
success/skip/failure/discovery cases with a CLI double that exits like the real one.
Read-only production checks on 2026-10-14, each about 9 seconds: WebApp `d69ff794` (App and
Site READY, exit 0), `fb0b8dd5` (App READY, Site CANCELED by its ignore step, exit 0),
`b4190731` (Site ERROR, one of the 21 failures of 2026-10-07/08, exit 1 naming
`gsadus-site`), PM `7f0187bf` (READY, exit 0). These were finished deployments; a live
BUILDING-to-READY transition was not triggered solely for testing.
