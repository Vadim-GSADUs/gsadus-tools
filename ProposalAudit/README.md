# ProposalAudit

Audits the proposal folders (PP, `G:\Shared drives\GSADUs Business\1 - Proposals`) and the
project folders (P, `G:\Shared drives\GSADUs Projects\<year> Projects`). It also keeps the
one tool that changes their subfolders. The findings below are dated snapshots, and git
history holds the full September and October write-ups.

## Open items

Waiting on the owner:

1. **Batch 9: a `Resource hub.md` in each proposal's `4 - Resources`.** Its shape was decided
   on 2026-10-28; it runs last, after the meeting with Sales and the owners.
   - The nightly CRM history job writes it, beside `CRM history.md`.
   - It holds this project's links, the folder guide, and company resources from one
     central list.
   - In Drive in the browser the `.md` opens as a Doc without making a copy, and agents
     read it on G:. No Google Doc twin is needed, and nothing goes into PP0.
2. **Batch 4: retire PP's `B - Questionnaire & Selection`.** Approved, but held until that
   meeting.
   - First, the WebApp document-templates thread moves the blank forms out of PP0.
   - Then trash the folder from PP0, and run
     `cleanup_branch.py --branch "3 - Supporting Documents/B - Questionnaire & Selection"`.
     It skips the 31 proposals with real files.
3. **Batch 2: change orders.** They stay in the PP, beside the contract, for now. The owner
   revisits this with Sales and the owners.
4. **The CRM history heartbeat (owner step).** Create a healthchecks.io check, then set
   `HEARTBEAT_URL_PROPOSAL_CRM_HISTORY` in Vercel. The job and its heartbeat are documented
   in WebApp `docs/proposals/CRM-HISTORY.md` ("A run", step 4).
5. **The walkthrough copies in P folders.** 6,435 files (43.7 GB) are byte copies of their
   PP's. PM's bridge now links each P to its PP's walkthrough instead.
   - Removing the copies needs the owner's yes and quarantine tooling (see Rules).
   - Telling staff to stop copying photos is the owner's.
6. **Misfiled items found by the plan-set audit (Batch 6).**
   - PP5's `D - Plan Set` holds a panel-upgrade subfolder (15 files) for P1's site.
   - P51's only design document, an HOA package, sits in PP292.
   - PP135's licensed survey is not in P28.
7. **The September worklist, re-checked against Drive on 2026-10-28.** These touch files, so
   they need the quarantine tooling.
   - **A2:** 26 byte-identical photos in PP231 (about 93 MB), and 4 same-name pairs with
     different content that need a ruling.
   - **A3:** 95 `Thumbs.db` and 2 `.tmp` files. Explorer re-creates `Thumbs.db`, so this
     recurs.
   - **A5:** 18 proposal folders whose names spell out the street suffix or hold a double
     space, as of 2026-09-10. Re-list them with `check_defects.py names`. They change only
     through `renameFolder()`.
   - **A6:** PP104 is still the only gap in PP1–PP389.
   - **A7:** PP118 and PP319 share one address. That is a CRM question.
   - **A8:** the same 10 proposals still hold no files: PP4, PP7, PP9, PP13, PP14, PP17,
     PP18, PP19, PP23 and PP54.
   - PP5's `EstimateReport (4).xls` is 13 months newer than P2's copy.
8. **The repo's visibility.** `Vadim-GSADUs/gsadus-tools` is public. The owner dropped this
   as a priority (addresses aren't secrets), but whether to make it private is the owner's
   call. A private repo can still be shared with staff.

Not waiting on the owner:

9. **Pipedrive's 54 won deals against 63 PP↔P pairs.** This was never reconciled.
10. **Notes for PM's P-folder housekeeping.** P0 belongs to the office, so these are not
    this folder's job.
    - P0's `C - Received from Client` still has its two subfolders.
    - 13 projects file at the DIV level, leaving `1 - Material` and `2 - Labor` unused.

## Use

```bash
# Drive API layer (gws, read-only): IDs, md5, Drive dates, the engine's tag. About 25 s.
python scan_proposals.py --drive --db data/census_pp.db
python scan_proposals.py --drive --db data/census_p.db --root "G:/Shared drives/GSADUs Projects/2026 Projects"   # and 2024, 2025
python analyze.py census --db data/census_pp.db -v        # per template folder: real use
python analyze.py crossdrive --db data/census_pp.db --p-db data/census_p.db -v
python analyze.py renames --before data/pre.db --db data/post.db --rename "OLD=>NEW"   # ID-by-ID check

# Mount layer (local walk, about 40 s) and the defects worklist
python scan_proposals.py
python analyze.py summary
python check_defects.py all

# The one tool that changes subfolders: a dry run unless --apply, logged step by step
python cleanup_branch.py --branch "<path>"                    # remove a branch where it is empty
python cleanup_branch.py --branch "<path>" --prune-empty      # keep it, drop empty folders beneath
python cleanup_branch.py --branch "<path>" --rename-to "<name>" [--move-files-to "<dest>"] [--create-missing]
python cleanup_branch.py --branch "<path>" --flatten          # lift files into it, drop its subfolders
python cleanup_branch.py --undo <log> [--apply]               # reverse a rename or flatten run
```

`data/` is local only. Every scan, log and manifest there names client folders, so its
`.gitignore` keeps all of it out of git.

## Rules

- **The mtime trap.** On the G: mount, mtime is sync and cache state, not use. The Drive
  API has its own version: a migration re-stamps `createdTime` and keeps `modifiedTime`.
  `analyze.py` finds bulk days from the data (10 or more folders created that day, or
  receiving mostly old files) and uses `modifiedTime` on them.
- **A real file** is not a shortcut, not OS litter, and not an untouched template copy.
  An untouched copy is matched by the engine's tag, by md5, or by the same bytes in 3 or
  more folders.
- **Proposal folder names change only through `renameFolder()`** in
  `WebApp/lib/proposals/drive.ts`, because the reconciler and the collision probe read
  them. Subfolders can change on the mount (owner, 2026-10-27), since no code reads their
  names. Drive for Desktop syncs a local rename or move as the same item, so IDs survive.
- **Never touch the `0 - P…` / `0 - PP…` shortcuts.** PM's bridge makes them, and they
  are the registry of signed projects. `cleanup_branch.py` refuses to touch any branch
  that holds a `0 - P*.lnk`.
- **PP0 is live.** The engine copies it at every mint (`walkTemplateTree`), so a change
  goes to PP0 first. Existing proposals don't pick up template changes, so they change
  through `cleanup_branch.py`.
- **P0 and the P folders belong to the office** (owner ruling, 2026-10-22). Normalizing
  P folders is PM's housekeeping.
- **Verify through the Drive API after sync,** ID by ID, against a scan taken before.
  The mount is not evidence.
- **Nothing here deletes a file.** `os.rmdir` refuses a non-empty folder, which is the
  safety net for folder work. Any change that removes files needs quarantine tooling
  first: move rather than delete, hashes before and after, and a refusal on any
  difference.
- **Client names and addresses stay out of committed files.** Name folders by PP# and P#.

## The proposal folder today (PP0, 2026-10-28)

```
1 - CRM                         Pipedrive history, written nightly: CRM history.md + Open deal in Pipedrive.url
2 - Proposal - Contract         everything the client signs; PM onboarding finds the contract here by name
    Change Order(s)             change orders, beside the contract (for now)
3 - Supporting Documents
    A - Site Walkthrough        walkthrough photos and videos (Images & Videos), filled site-visit forms
    B - Questionnaire & Selection   retiring (Batch 4)
    C - Received from Client    anything the client sends, no subfolders
    D - Plan Set                pre-contract references; approved and stamped sets go to the P
    Z - Other Misc
4 - Resources                   Resource hub.md (Batch 9)
0 - P… shortcut                 added by PM's bridge once the project exists
```

**On the P side**, PM's bridge adds two shortcuts: `0 - PP…` at the root, and
`PP### Site Walkthrough` in `2 - Supporting Documents/A - Project (Images & Videos)`.
Approved plans, permits, construction media, takeoffs, receipts and closeout papers live
in the P.

## What was changed

Each change was owner-authorised and verified through the Drive API unless noted. The undo
logs live in `~/.gsadus/proposal-folder-renames/`.

| Date | Change | Result | Undo |
|---|---|---|---|
| 2026-09-10 | Tiers 1–3: the empty DIV tree and dead placeholders removed from every PP | 33,584 → 4,692 folders, 0 files lost (mount-verified) | past the 30-day trash window |
| 2026-10-27 | Batch 1: `1 - Takeoff & Estimate` → `1 - CRM`, `4 - Administrative & Change-Issues` → `4 - Resources` | 763 renamed in place, 9 created, PP221's 2 files moved to the contract folder; 41,610 files unchanged | `cleanup_branch.py --undo …/2026-10-27-batch1.jsonl` |
| 2026-10-28 | Batch 5: `C - Received from Client` flattened | 779 subfolders gone, 40 files lifted with the same IDs | `cleanup_branch.py --undo …/2026-10-28-batch5-flatten.jsonl` |
| 2026-10-28 | Batch 6: 5 approved plan files moved from PPs into their P, beside the P's own copies | IDs, parents and md5 checked | `…/2026-10-28-batch6-moves.jsonl`: each line has `file_id`, `from`, `to`; reverse it with `files.update` (`addParents` = from, `removeParents` = to) |
| 2026-10-28 | Batch 3: the P → PP walkthrough shortcut, as PM's bridge's third leg | all 63 pairs, Vadim-GSADUs/gsadus-pm#1 | in PM |
| 2026-10-28 | Batch 7: the nightly CRM history into each PP's `1 - CRM` | WebApp `23eb889f`; contract in `docs/proposals/CRM-HISTORY.md` | in WebApp |

Batch 8 (client details in this repo) was dropped on 2026-10-28.

## Findings worth keeping (census of 2026-10-27)

- **The September counts included template copies.** 1,502 untouched copies sit in 251 PP
  folders (876 MB), and only 452 of them carry the engine's tag. Recycling them belongs to
  the WebApp document-templates thread. A recycle by tag alone would miss about 70%.
- **Use of each PP0 folder over the last 12 months (241 proposals, PP149 on):**

  | Folder | Use |
  |---|---|
  | `2 - Proposal - Contract` | 90% |
  | `A - Site Walkthrough` | 98% |
  | `Change Order(s)` | real files in 45 PPs, all signed projects, mostly filed after the P existed |
  | `B - Questionnaire & Selection` | real files in 31 PPs; 754 blank copies |
  | `D - Plan Set` | 30 PPs |
  | `Z - Other Misc` | 34 PPs |
- **The walkthrough was kept twice.** 72% of the real files in P folders were byte copies
  of their PP's, mostly walkthrough photos. Batch 3 stops this for new projects.
- **Plan files in PPs are pre-contract references** (Batch 6 audit of 139 files):
  - 44 are our own early drawings, 32 client documents and 10 downloaded pre-approved
    sets;
  - 24 of the 29 approved files had already been copied to their P.
- **Late-life P folders read low** (construction, closeout, `4 - Administrative`) because
  few projects have reached those stages yet.

Refresh any of these with `scan_proposals.py --drive` and `analyze.py census` or
`crossdrive`.
