# Inbox

Read-only triage of vadim@gsadus.com's mail for the Orchestrator's morning brief (owner
2026-10-14). The workflow, the hard rules and the owner's standing decisions live in the
workspace skill `C:\GSADUs\.claude\skills\inbox\SKILL.md`; this folder holds the tooling.

| File | Purpose |
|---|---|
| `inbox-scan.ps1` | Lists mail since the last scan (headers only), classifies it and prints the report. pwsh 7 + the `gws` CLI. |
| `inbox-rules.json` | Ordered classification rules, first match wins. Vendor senders only: no client or staff personal details. |

```powershell
pwsh -NoProfile -File C:\GSADUs\Tools\Inbox\inbox-scan.ps1 -ForwardReceipts            # the morning brief's run
pwsh -NoProfile -File C:\GSADUs\Tools\Inbox\inbox-scan.ps1 -Since (Get-Date).AddDays(-7) -NoState   # look back, no state
pwsh -NoProfile -File C:\GSADUs\Tools\Inbox\inbox-scan.ps1 -Body <message-id>          # one body, sanitized
pwsh -NoProfile -File C:\GSADUs\Tools\Inbox\inbox-scan.ps1 -Json -NoState              # classified messages as JSON
```

- **Window.** Mail newer than the last successful scan on this machine
  (`$HOME\.gsadus\inbox-scan.json`; each PC keeps its own), else the last 2 days; never more
  than 14 days back. `-NoState` leaves the marker alone.
- **Classes.** `needs` (the owner acts), `watch` (verify before reporting), `receipt` (sent to
  accounting), `receipt?` (a possible receipt from an unlisted sender: confirm first), `review`
  (no rule matched: judge it), `info` (routine, counted) and `noise` (counted). Unmatched mail
  with a List-Unsubscribe header counts as noise and is named as an unsubscribe candidate.
- **Writes.** Only `-ForwardReceipts`: each `receipt` without the `Forwarded to accounting`
  label is forwarded to accounting@gsadus.com with its attachments (`gws gmail +forward`),
  then labelled so it never goes twice. Nothing else writes.
- **Redaction.** Subjects, snippets and bodies are printed with links removed and card
  digits, account numbers, codes and long numbers masked.
- **Auth.** Exit code 3 with `INBOX: gws needs re-login` when the gws token has lapsed (the
  owner re-authorizes about daily: `gws auth login`).
- **Rules.** Add a rule when a sender recurs in `review`, or a newsletter keeps arriving.
  Test with `-Since … -NoState` before committing.
