# Stream Deck

The owner's 15-key Stream Deck MK.2 on gsadus-vadim: one profile, `GSADUs`, with three pages
in the free Elgato Stream Deck app (no plugins). The bottom-right pair on each page jumps to
the other two pages.

- **Dev:** WebApp `:3009`, Site `:3019` and PM `:3200` on localhost; Claude app; start Docker
  (through `ShellProfile\Start-DockerDesktop.ps1`, the workspace's only way); `unwip-all` and
  `wip-all` in a pwsh window that stays open; Vercel and Supabase dashboards.
- **Revit:** start Revit 2026; VG, TL, EH, HI, HR, 3D, WT and RL (pyRevit Reload) typed into
  the focused Revit window; the GSADUs Projects drive. `3D` needs that shortcut assigned to
  Default 3D View in Revit's Keyboard Shortcuts.
- **Office:** Meet mute (Ctrl+D) and camera (Ctrl+E), which reach only a focused Meet tab;
  open Meet; Pipedrive; PM live; the GSADUs Projects and Business drives.

## Change a key

`build-profile.py` holds the key list and writes `GSADUs.streamDeckProfile`; icons come from
`icons\make-icons.mjs` (its header says how to run it). After a rebuild, delete the old GSADUs
profile in the app (Preferences → Profiles → −), then double-click the new file to import it.
Commit the script, the icons and the rebuilt profile together.
