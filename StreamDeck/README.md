# Stream Deck

The owner's 15-key Stream Deck MK.2 on gsadus-vadim: three pages (Dev, Revit, Office), free
Elgato Stream Deck app, no plugins. Icons are baked with their label, so leave each key's
title empty.

## Load it

Install the app (`winget install --id Elgato.StreamDeck --exact --source winget`), then
double-click `GSADUs.streamDeckProfile` to import the profile. After changing keys, export it
again over that file (Preferences → Profiles → ⋯ → Export) and commit.

## Keys

Action names are the app's (System → Open, Website, Text, Hotkey; Navigation → Go to Page).
Icons are in `icons\`, named `<page>-<key>.png`.

| Page | Key | Action | Setting |
|---|---|---|---|
| Dev | WebApp 3009 / Site 3019 / PM 3200 | Website | `http://localhost:3009`, `:3019`, `:3200` |
| Dev | Claude | Open | `claude.cmd` |
| Dev | Docker | Open | `start-docker.cmd` (the workspace's only way to start Docker) |
| Dev | unwip-all / wip-all | Open | `unwip-all.cmd`, `wip-all.cmd` (a pwsh window that stays open) |
| Dev | Vercel | Website | `https://vercel.com/vadim-7430s-projects` |
| Dev | Supabase | Website | `https://supabase.com/dashboard/project/rfguhdxdrzbqcoiulyuv` |
| Revit | Revit 2026 | Open | `C:\Program Files\Autodesk\Revit 2026\Revit.exe` |
| Revit | VG, Thin lines, Hide, Isolate, Reset, Tile, pyRevit | Text | `VG`, `TL`, `EH`, `HI`, `HR`, `WT`, `RL` |
| Revit | Default 3D | Text | `3D` (assign it first in Revit → Keyboard Shortcuts) |
| Office | Meet mute / Meet camera | Hotkey | Ctrl+D / Ctrl+E |
| Office | Open Meet | Website | `https://meet.google.com` |
| Office | Pipedrive | Website | `https://goldenstateadus.pipedrive.com` |
| Office | PM live | Website | `https://gsadus-pm.vercel.app` |
| Office | Projects / Business | Open | `G:\Shared drives\GSADUs Projects`, `…\GSADUs Business` |
| all | bottom-right pair | Go to Page | the other two pages |

The Text and Hotkey keys type into the focused window: Revit must be in front for the Revit
keys, and the Meet tab for mute and camera.

## Icons

144×144 PNGs drawn from Tabler outline icons, one ink color per page. To add or change one,
edit the list in `icons\make-icons.mjs` and run it as its header says.
