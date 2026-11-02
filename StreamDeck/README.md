# Stream Deck

The owner's 15-key Stream Deck MK.2 on gsadus-vadim: one profile, `GSADUs`, one page, in the
free Elgato Stream Deck app (no plugins). Icon ink says what a key does: purple opens a page or
folder, green runs a command in a pwsh window that stays open, orange stops something.

- **Top row:** Gmail and Chat (as vadim@gsadus.com), PM, WebApp, gsadus.com.
- **Middle row:** Pipedrive, the Proposals and Projects drives, the Helpdesk inbox, and the
  Board in Obsidian.
- **Bottom row:** Orchestrator (the Claude app, where the Orchestrator session is pinned),
  `gws-login`, `gcloud auth login`, `remote`, and Stop servers.

Stop servers (`stop-dev-servers.ps1`) lists the node servers on the dev port lanes with their
repo or worktree, and on a yes stops each with its node process tree. The shell or agent that
started one stays open. It asks first because an agent may be using a server.

## Change a key

`build-profile.py` holds the key list and writes `GSADUs.streamDeckProfile`; icons come from
`icons\make-icons.mjs` (its header says how to run it). After a rebuild, delete the old GSADUs
profile in the app (Preferences → Profiles → −), then import the new file (the ⌄ menu under
the profile list → Import). Commit the script, the icons and the rebuilt profile together.
