# Sentry

Read-only probe for the `gsadus` Sentry org — every project (`webapp`, `pm`, …) from any
repo, any directory. GET-only by construction; status changes stay in the Sentry UI or
the claude.ai Sentry connector.

| File | Purpose |
|---|---|
| `sentry-probe.mjs` | The probe. Node ≥ 18, no dependencies. |

```powershell
sentry-probe projects                                   # shell-profile function (pwsh)
sentry-probe issues --project pm
sentry-probe issue PM-10                                # detail + latest event tags/stack
sentry-probe events --project webapp --query "transaction:/site-check" --period 14d
node C:/GSADUs/Tools/Sentry/sentry-probe.mjs issues     # same thing from Git Bash / any shell
```

Add `--json` for the raw API payload. Default `issues` query is `is:unresolved`; pass
`--query ""` to include resolved.

**Auth.** `SENTRY_AUTH_TOKEN` from the environment if set, otherwise read at call time
from Doppler `webapp/dev` — the single place the token lives (a personal token on the
owner's account, scopes `project:read event:read org:read project:releases`, created
2026-09-03). Never print it. Rotate in Doppler; nothing else changes.
