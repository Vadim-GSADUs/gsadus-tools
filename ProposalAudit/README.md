# ProposalAudit

Read-only forensics on the proposal corpus at
`G:\Shared drives\GSADUs Business\1 - Proposals` (358 folders as of 2026-09-10).

Nothing in this folder writes to Drive. `scan_proposals.py` walks and records;
`analyze.py` only queries the resulting SQLite. Any future remediation tooling must
live in its own script with an explicit dry-run — never bolted onto these.

## Use

```bash
python scan_proposals.py                 # full corpus -> data/audit.db  (~40s)
python analyze.py summary                # corpus shape: what is used, what never is
python analyze.py occupancy --under "1 - Takeoff & Estimate"
python analyze.py never
python analyze.py outliers
python analyze.py dupes

python check_defects.py all              # the housekeeping worklist
python check_defects.py names --csv data/rename_manifest.csv
```

`analyze.py` answers *what shape is this corpus in*; `check_defects.py` answers
*what is specifically wrong with it*. Both are read-only. `check_defects.py dupes`
is the only thing that reads file bytes (to hash them) — still a read.

`--subtree` scopes a scan to one branch; `--resume` skips proposals already recorded
at that scope; `--only` takes a regex over proposal names.

## Why it is built this way

**Threads.** The Drive mount is latency-bound. A single-threaded `find -printf` over
the corpus was projected at ~9 hours; 24 threads over `os.scandir` finish in 39
seconds. `scandir` matters as much as the threads — on Windows the directory listing
already carries size and mtime, so `DirEntry.stat()` costs nothing, while `os.walk`
forces a stat syscall per file.

**SQLite, not a report.** The dataset is the deliverable. Every question is a query,
so re-asking never means re-deriving. 71,896 node rows, one per file and folder.

## The mtime trap

**`mtime` in this database is not a usage signal, and no analysis here treats it as
one.** On a Drive for Desktop mount it reflects sync and local cache state, and the
proposal engine's mint writes fresh timestamps onto every object it creates. Judging
"untouched" by mtime would mark the entire corpus as recently active and be wrong.

Occupancy (`n_deep > 0` — does a real file live anywhere beneath this folder?) is the
only defensible use signal available from the filesystem, and it is what every
command here reports.

True human-activity signals require the Drive API, not the mount:

- `lastModifyingUser` — the cheap win. An object whose only modifier is the engine's
  service account has never been touched by a person.
- `viewedByMeTime` — per-user, weak, free.
- Drive Activity API (`driveactivity.googleapis.com/v2/activity:query`) — real
  create/edit/move/comment events. The only true answer.

A `scan_drive_api.py` layer using the existing proposal-engine service account is the
intended next step; it also reads the `gsadus_template_item` / `gsadus_pp` /
`gsadus_intent` `appProperties` that `WebApp/lib/proposals/drive.ts` stamps on every
engine-minted object, which the filesystem cannot see at all.

## Findings — 2026-09-10, full corpus

| | |
|---|---|
| Proposals | 358 |
| Directories | 33,584 |
| **Empty directories** | **31,234 (93.0%)** |
| Files | 38,312 |
| Size | 229.3 GB |
| Distinct folder paths in the design | 269 |
| Paths that have **ever** held a file | **31** |
| Paths that have **never** held a file | **238 (88.5%)** |

**The DIV tree is entirely unused.** All 63 paths under
`1 - Takeoff & Estimate/3rd Party Coordination` (`DIV-01 … DIV-20`, each with
`1 - Material` / `2 - Labor`, plus `0 - Takeoff & Estimating/V1`) hold **zero files
across all 358 proposals** — ~22,400 real folders, never used once, not in the oldest
proposal and not in the newest.

Where work actually lands:

| Used in | Files | Path |
|---:|---:|---|
| 341/358 | 36,818 | `3 - Supporting Documents` |
| 338/358 | 1,354 | `2 - Proposal - Contract` |
| 328/358 | 34,957 | `3 - Supporting Documents/A - Site Walkthrough/Images & Videos` |
| 242/358 | 388 | `2 - Proposal - Contract/Change Order(s)` |
| 223/358 | 688 | `3 - Supporting Documents/B - Questionnaire & Selection` |
| 110/358 | 324 | `3 - Supporting Documents/Z - Other Misc` |
| 29/358 | 69 | `3 - Supporting Documents/D - Plan Set` |
| 6/357 | 64 | `1 - Takeoff & Estimate` (PP1–PP6 and PP110 only) |
| 1/350 | 2 | `4 - Administrative & Change-Issues` (PP221 only) |

Nine folders carry the business. The template ships 93.

**Structure is stable, not drifted.** 340 of 358 proposals share the identical
94-directory shape. The real outliers are ~11: PP358 (187 dirs), PP1 and PP221 (95),
PP2–PP8 (88–89, pre-standardisation), PP235 (47), PP234 (18).

**Humans create folders the template never offered.** Hand-made subfolders appear
under `A - Site Walkthrough` — `Existing House Photos/{Architectural, Structural}`,
`Elsegundo` — and under `Z - Other Misc` and `D - Plan Set`. The template supplies 63
folders nobody has ever wanted and none of the few people keep building by hand.

**`(1)` collision suffixes split into two distinct problems.** Exactly one proposal,
PP358, has duplicate *folders*: its entire four-branch structure exists twice, with
real files on both sides (13 files in the `(1)` copies) — a live mint bug, not
historical debt. Separately, 18 proposals hold suffixed *files*, which turn out to be
three unrelated situations; see A2 below for the resolved breakdown. `analyze.py
dupes` reports folders and files separately and they should never be conflated.

## Housekeeping worklist — findings 2026-09-10

Reported by `check_defects.py`. Nothing below has been executed; all of it is
read-only output awaiting a decision. **A1 (PP358's forked structure) is excluded**
— that is a live auto-mint bug under separate investigation, not historical debt.

### A2 · `(1)` collision-suffixed files — resolved, mostly benign

48 suffixed files, and they are three different problems wearing one costume:

- **26 byte-identical copies**, all in `PP231 3941 Pinell St` (`IMG_9136`–`IMG_9161`)
  — one duplicated photo import. sha256-confirmed redundant, ~93 MB.
- **4 same-name / different-content pairs** that are *not* duplicates and need a
  human to say which is authoritative:
  `Scanned Document (1).pdf` (PP21, 8.9 MB vs 20 MB) ·
  `PP216 … ADU Site Visit Form` (135 KB vs 21 KB) ·
  `PP248 - ADU Site Visit (1).pdf` (115 KB vs 159 KB) ·
  `ADU Site Visit - MASTER (1).pdf` (PP290, 1.2 MB vs 86 KB)
- **18 misnamed singletons** with no original in the folder — the `(1)` is simply
  part of the filename, acquired at download. Cosmetic at most.

This corrects an earlier reading of the data: `Change Order #1/#3 - signed (1).pdf`
in PP33 have **no** non-`(1)` counterpart. They are single files with an awkward
name, not duplicated signed contracts. There is no contract-integrity issue here.

### A3 · OS litter
95 × `Thumbs.db` (88 MB), 2 × `.tmp`, 1 × `desktop.ini`. Safe to delete — but
`Thumbs.db` regenerates whenever anyone browses the folder in Explorer, so this is a
recurring symptom rather than a one-time cleanup.

### A4 · PP↔P project links — **NOT debt. DO NOT DELETE.**

An earlier pass in this file mistook these for legacy litter. They are the opposite:
the owner mints a `0 - P<n> <address>.lnk` shortcut into a proposal's root **when the
project is signed and onboarded**, binding the PP# proposal to its P# project folder.

That makes them the single most valuable piece of lifecycle data in the corpus — the
only won/lost signal that exists anywhere in the filesystem:

- **59 of 357 proposals carry a P-link → 16.5% reach signed + onboarded.**
- **P1–P59 with no gaps**, so the shortcuts constitute a complete PP↔P registry
  derivable from the mount alone (`check_defects.py projects` prints it).
- **8 of the 59 carry `- Cancelled` in the shortcut name** (P4, P8, P15, P20, P31,
  P34, P36, P53) — project status encoded in the filename. 51 live.
- Two shortcuts spell the address differently from their folder — `807 Stoneridge
  Cir` vs `…Circle`, `7443 Henrietta Dr` vs `…Drive`. The shortcut is the newer
  artifact and already uses the USPS form, which independently supports A5.

Anything that sweeps `.lnk` files, or "normalises" a proposal folder by removing
non-template contents, would destroy this. Any future remediation script must treat
`0 - P*.lnk` as protected.

### A5 · Folder-name hygiene — manifest ready
18 renames: 16 unabbreviated street suffixes and 2 double-spaces. USPS abbreviation
is already the canonical form (Amendment A5, 2026-08-27, enforced on new mints by
`@gsadus/pipedrive` 0.1.3), so this is backfill to a decided standard, not a new rule.
Manifest: `data/rename_manifest.csv`.

Only the **final** token is treated as a street suffix. `PP54 165 Terrace St` is
Terrace *Street* — abbreviating mid-name words would corrupt the street name, so any
name with a suffix word away from the end is surfaced for a human instead of guessed.

Renames must go through `renameFolder()` in `WebApp/lib/proposals/drive.ts` — the one
sanctioned rename path. Drive URLs are ID-based so links survive, but the reconciler
observes names and `candidateNumberTaken()` matches on them.

### A6 · PP104 is the only gap in PP1–PP358
The filesystem cannot say whether it was never minted, trashed, or is a burned
number. The engine's claim registry can.

### A7 · PP118 and PP319 are both "212 Imad Ct"
112 files / 370 MB versus 14 files / 457 MB, and **zero filenames in common** — which
argues for two genuinely distinct jobs at one address rather than a duplicate mint.
A CRM question, not a filesystem one.

### A8 · 10 proposals hold zero files
PP4, 7, 9, 13, 14, 17, 18, 19, 23, 54 — **928 folders between them, not one file.**
Minted, then abandoned. Conforming their structure would be work spent on folders
nobody will open; the useful move is lifecycle state, so they stop counting toward
"our 358 projects."

### Still open before anything is executed
A remediation script (separate file, `--apply` off by default, `0 - P*.lnk` treated as
protected, manifest logged before/after) is the next build — but it should wait on:
sign-off to delete the 26 confirmed-identical photos; a ruling on the 4 divergent
pairs; a decision on the 10 empty shells; and the registry lookups for PP104 and
PP118/PP319. A4 is closed: the project links stay.

## Cross-drive finding — why the DIV tree is empty (2026-09-10)

Scanned the Projects drive with the same tool
(`--root "G:/Shared drives/GSADUs Projects/<year> Projects" --db data/projects.db`)
— 62 projects, 6,170 dirs, 91.6% empty, 8,036 files, 55.2 GB.

**The DIV tree is a PROJECT-phase workspace that was copied into the PROPOSAL
template by mistake.** On the P side it is genuinely used: P1 holds subcontractor
bids in `DIV-06 = Electrical`, `DIV-07 = Plumbing`, `DIV-14 = Insulation`; P2 in
`DIV-06`, `DIV-11 = Openings`, `DIV-12 = Wall Coverings`. On the PP side it holds
zero files in 358 proposals — because subcontractor bids do not exist at proposal
stage, by definition.

The P side has since renamed it `3rd Party Coordination & Bids` (57 projects); the
5 oldest projects and the entire proposal template still carry the original
`3rd Party Coordination`. Two forks of one origin, drifted apart.

That reframes the recommendation: this is not "delete 63 unused folders", it is
"remove the construction-phase workspace from the proposal template, where it can
never apply, and leave it on the project side, where it works."

### History that explains PP1–PP6

All 64 PP-side files under `1 - Takeoff & Estimate` belong to PP1, PP2, PP3, PP5,
PP6 (plus one stray HEIC in PP110), and every one sits in
`Approved Project Estimate/V1` — never in the DIV tree. These predate the Projects
drive: project work happened in proposal folders because there was nowhere else.
Once the Projects drive existed the work moved, and the proposal template kept the
vestigial branch for 352 more mints.

### Migrating PP -> P would be a no-op, and is already done

Of those five, three carry a P-link: PP2→P1, PP5→P2, PP6→P3. Their estimate content
**is already on the P side**, hash-verified: 30 of 32 files byte-identical.

Four exceptions, worth a human ruling before anything is deleted:

| File | PP side | P side | Note |
|---|---|---|---|
| `EstimateReport (2).xls` (PP2/P1) | 36,864 B · 2024-10-01 | same size, same mtime | differs only in bytes — `.xls` internal metadata; effectively identical |
| `EstimateReport (5).xls` (PP6/P3) | 32,256 B · 2024-10-01 | same size, same mtime | same |
| **`EstimateReport (4).xls` (PP5/P2)** | 33,280 B · **2025-11-25** | 33,280 B · 2024-10-01 | **PP copy is 13 months NEWER — deleting it would lose the later edit** |
| `#3.jpg` (PP5/P2) | 235,656 B · 2024-10-01 | 237,124 B · 2024-12-18 | P side newer and larger; P wins |

**PP1, PP3 and PP110 have no P-link** — they never converted, so their estimate
content exists only on the proposal side. It is not redundant and has nowhere to
migrate to.

Also note the `EstimateReport (1)…(5).xls` numbering, which A2 flagged as "misnamed
singletons": the `(n)` is a browser download counter across one session, one report
per proposal. Distinct files, not duplicates — A2's reading holds.

## EXECUTED — Tier 1, 2026-09-10 (owner-authorised)

Removed `1 - Takeoff & Estimate/3rd Party Coordination` and everything beneath it
from the proposal template and every proposal that had it.
Script: `cleanup_tier1.py` (dry run default; `--apply` to execute).
Manifest: `data/tier1_manifest.jsonl`, one JSON line per proposal.

**26,014 folders removed across 357 proposals + the template. 0 failures. 52.6s.**

| | Before | After |
|---|---:|---:|
| Directories | 33,584 | **7,665** |
| Empty directories | 31,234 (93.0%) | 5,306 (69.2%) |
| Files | 38,312 | **38,380** |
| Size | 229.3 GB | 229.6 GB |
| Distinct folder paths | 269 | 176 |
| Template | 94 dirs | **21 dirs**, 13 files intact |

**Zero data loss, independently verified by re-scan:** the file count went UP, not
down — two new proposals (PP359, PP360) were minted by the live engine *while the
cleanup was running*, which is also the clearest possible illustration of why the
template had to be fixed first. 349 proposals now sit at the clean 21-directory shape.

Verified intact after the pass: all **59 `0 - P*.lnk` project links**; the takeoff
content in all six proposals that had it (PP1 11 files, PP2/PP3/PP5/PP6 13 each,
PP110 1); `2 - Proposal - Contract` (360 proposals, 1,367 files) and
`3 - Supporting Documents` (360 proposals, 36,886 files) untouched.

### Why it was safe

`os.rmdir` bottom-up, never `shutil.rmtree`. The OS refuses to remove a non-empty
directory, so no file could be destroyed even if the emptiness check had been wrong —
the guarantee did not depend on the analysis being correct. Every subtree was also
re-walked live immediately before deletion rather than trusted from the audit db;
**0 of 357 were skipped**, confirming the zero-files finding at execution time.

Deletions land in the shared drive trash (30-day recovery) and are attributed to the
signed-in owner. Running locally rather than through the Drive API was a deliberate
owner decision: the OS safety net was judged worth more than collapsing 73 syscalls
per proposal into one server-side call.

### Remnants — resolved same day

Both hold-outs were cleared by the owner by hand on 2026-09-10, and a re-scan
confirms **zero `3rd Party Coordination*` folders remain anywhere in the corpus**:

- **PP358** — the mint-bug investigation closed. Its DIV tree AND the forked `(1)`
  structure are gone; it now sits at exactly the clean 21-directory template shape
  with 89 files (the 13 removed were the duplicate template payload, not client work).
  It is no longer excluded by `cleanup_tier1.py`.
- **PP1** — its `3rd Party Coordination (Bids)` variant is gone. 22 dirs, 489 files
  intact, including the hand-made `Existing House Photos/{Architectural, Structural}`
  subfolders that the template never offered.

Final corpus state: **360 proposals, 7,519 directories, 5,160 empty (68.6%), 38,380
files, 229.6 GB.** 350 proposals sit at the 21-directory shape.

### Scanner bug found and fixed during verification

`scan_proposals.py` used `INSERT OR REPLACE` and never removed rows for nodes that
had disappeared, so a re-scan after any deletion left **ghost rows** behind — the
first post-cleanup scan reported 146 DIV nodes that did not exist on disk. Any
occupancy analysis run after a cleanup would have silently over-reported.

Fixed: each proposal's rows are cleared before its fresh walk is inserted, scoped to
`--subtree` when one is given so a partial scan never discards the wider corpus.
Verified by rebuild — `nodes` dir rows now equal the `proposals` total exactly.

The lesson generalises: **a scan database is a cache, not a record.** Only the
filesystem is authoritative, which is why `cleanup_tier1.py` re-walks live rather
than trusting the db it was planned from.

### Note on the databases

`data/audit.db` is the pre-cleanup snapshot (358 proposals) and is now historical.
`data/audit_after.db` is the post-cleanup state (360 proposals). Re-run
`scan_proposals.py` to refresh; the two are kept apart so the before/after comparison
above stays reproducible.

## Archive of the pre-cleanup template — 2026-09-10

`G:\Shared drives\GSADUs Business\2 - Golden State ADU's\_Archive\A - PP0 0000 Address Ln. - Template`

The owner copy-pasted the template into `_Archive` after the Tier 1 pass, which
captured the **post**-cleanup shape (21 dirs) — the template was the first thing
cleaned that afternoon, so a live copy could no longer show the original. The DIV
tree was rebuilt into the archive copy from `data/audit.db`, the pre-cleanup
snapshot, and verified path-for-path: **93 directories, 13 files, exact match to the
snapshot, nothing missing and nothing extra.** The live template was re-checked
afterwards and remains clean (21 dirs, 0 DIV folders).

The restore script guards on `_archive` appearing in the target path and refuses any
target under `1 - Proposals` — recreating the DIV tree in the live template would
undo the cleanup and every future mint would carry it again.

A plain-text rendering of the same structure is at `data/template_pre_cleanup.txt`.
It is the more durable reference of the two: it cannot be accidentally re-minted
from, does not add 93 empty folders to Drive, and stays greppable.

## EXECUTED — Tier 2, 2026-09-10 (owner-authorised)

Removed the remainder of `1 - Takeoff & Estimate` — `Approved Project Estimate/V1`
and the branch itself — wherever it was empty.

**1,059 folders removed across 353 proposals + the template. 0 failures. 3.0s.**

| | After Tier 1 | After Tier 2 |
|---|---:|---:|
| Directories | 7,519 | **6,460** |
| Empty directories | 5,160 (68.6%) | 4,101 (63.5%) |
| Live template | 21 dirs | **18 dirs**, 13 files |

353 present-and-empty + 6 skipped + 1 absent (PP234) = 360, and 1,059 = 353 × 3.
350 proposals now sit at 18 directories.

**Six proposals kept their branch**, holding real estimate content from before the
Projects drive existed: PP1 (11 files), PP2/PP3/PP5/PP6 (13 each), PP110 (1). PP5's
`EstimateReport (4).xls` — the copy 13 months newer than P2's — is inside a kept
branch, so that open ruling is preserved untouched.

Verified after the pass: 59 `.lnk` project links intact; `2 - Proposal - Contract`
(1,367 files) and `3 - Supporting Documents` (36,887 files) untouched; no ghost rows.

### The script was generalised, not copied

`cleanup_tier1.py` became **`cleanup_branch.py --branch "<path>"`**. A second
near-identical script would have been duplicate code of exactly the kind workspace
rule 6 forbids, and would have drifted from the original the first time either was
fixed.

The refactor also removed the need for a protected-proposals list. **The emptiness
rule IS the keep-list**: a branch holding even one file is skipped and reported, so
the six proposals survived by the same mechanism that made Tier 1 safe — not by a
name someone remembered to hardcode. The dry run named all six correctly before a
single folder was touched.

Manifest is now `data/cleanup_manifest.jsonl` (carries the branch per row);
`data/tier1_manifest.jsonl` remains as the Tier 1 record.

## The failed restore — 2026-09-10 (recorded because it cost a day's sync)

Tier 2 removed `1 - Takeoff & Estimate` itself. The owner wanted that folder kept as
a **named placeholder** — removing the empty `Approved Project Estimate` was fine,
removing the parent was not — and asked for it back via the Drive API rather than by
re-creating folders.

`restore_branch.py` was written to un-trash the 353 folders by ID. **It brought back
26,753 folders**, the entire DIV tree included, and they had to be deleted a second
time. No files were affected — that subtree held none, and `os.rmdir` could not have
touched one — but it cost an extra delete pass and two more sync cycles on every
machine.

### Two mistakes, and only one of them is interesting

**The wrong belief:** that because the cleanup deleted bottom-up with `os.rmdir`,
Drive had recorded a separate trash operation per folder, so un-trashing a parent
would return it empty. Drive does not work that way. It collapses the deletion into
the highest folder removed, and restoring that folder restores its subtree.

**The wrong check** — this is the one worth remembering. The script queried
`explicitlyTrashed` and printed `NOT explicitly trashed: 0`, captioned "must be 0 —
else children would return too". That number was real and it was irrelevant: the
query matched only folders **named** `1 - Takeoff & Estimate`, which were of course
explicitly trashed. It never looked at a single descendant, which was the actual
question. The check appeared to validate the assumption while structurally being
incapable of testing it.

> A check that reports success on the wrong subject is worse than no check. It
> manufactures confidence, and it survives review because the output looks like
> evidence.

**The judgment error underneath both:** re-creating 353 empty folders would have
taken seconds with no sync churn. The argument for un-trashing was preserving
`gsadus_template_item` identity — which applied to **16 of 353 folders**, all of
which are now empty placeholders anyway. A real-but-narrow concern was allowed to
select the complicated path.

`restore_branch.py` is kept, working and correctly documented: it un-trashes a branch
**and its subtree** by identity. Its header now leads with that, and the
`explicitlyTrashed` line is explicitly marked as not a safety check.

### Corrective pass

Re-ran the two child branches, leaving the parent standing:

```
cleanup_branch.py --branch "1 - Takeoff & Estimate/3rd Party Coordination" --apply
cleanup_branch.py --branch "1 - Takeoff & Estimate/Approved Project Estimate" --apply
```

**26,430 folders removed, 0 failures.** Five proposals skipped for holding files.
End state: `1 - Takeoff & Estimate` present in all 359 proposals, 353 of them bare
placeholders, 64 takeoff files preserved, 59 `.lnk` links intact.

## EXECUTED — Tier 3, 2026-09-10 (owner-authorised)

`4 - Administrative & Change-Issues`, pruned rather than removed: the branch folder
stays as a named placeholder, only the empty directories beneath it go.

**1,759 folders removed across 352 proposals + the template. 0 failures. 3.1s.**

| | Before | After |
|---|---:|---:|
| Corpus directories | 6,451 | **4,692** |
| Corpus files | 38,484 | **38,484** — unchanged |
| Bare placeholders | 0 | **351** |
| `0 - P*.lnk` links | 59 | **59** |

Delta reconciles exactly: 6,451 − 4,692 = 1,759 = 351 × 5 + PP221's 4.

### This one is not the DIV story, and the P drive proved it

Occupancy across all 352 proposals carrying the branch:

| Sub-folder | Files corpus-wide |
|---|---:|
| `1 - Change Order Request` | 0 |
| `2 - Construction Issues` | 2 — PP221 only |
| `3 - Warranty` | 0 |
| `4 - Notice Of Completion` | 0 |
| `5 - Certificate of Occupancy` | 0 |

The DIV tree was removed because it *could never apply* at proposal stage. That
reasoning was checked here before being reused, and **it does not hold**: the P drive
carries the same `4 - Administrative` skeleton with identical children, and uses it —
6 files across 62 projects, including `Change Order Request` and `Certificate of
Occupancy`, both dead on the proposal side.

So these are genuine post-signing artifacts filed on the project side. Not a category
error — a redundant skeleton. The template was stripped to a bare placeholder on that
basis (owner decision): the P folder already carries the structure where the documents
actually land.

**PP221 keeps `2 - Construction Issues` and its 2 files**, having lost only its 4 dead
siblings. The old whole-branch mode would have skipped PP221 entirely and left them.

### `--prune-empty`

`cleanup_branch.py` gained a mode that keeps the branch folder and removes only empty
directories beneath it — derived per proposal, not from a hand-typed child list that
is wrong the moment one proposal differs. It is also the mode that would have avoided
the takeoff mistake above.

The `os.rmdir` guarantee is unchanged. A parent that qualifies as empty implies its
children do, they are all listed, and removal is still deepest-first — so the OS sees
a genuinely empty directory at every call and cannot delete a file even if the
emptiness logic were wrong.

## Day close — 2026-09-10

| | Start of day | End of day |
|---|---:|---:|
| Directories | 33,584 | **4,692** |
| Files | 38,482 | 38,484 *(2 added by live use)* |
| Files lost | — | **0** |
| Proposals | 358 | 360 *(PP359/PP360 minted live during the session)* |

350 proposals now share an identical 13-directory shape. Every deletion this day was
an empty directory, protected by `os.rmdir` refusing non-empty targets — which is why
the one wrong call was recoverable at the cost of sync time and nothing else.

### Still open

- **A2** — 26 byte-identical photos in PP231 (~93 MB); 4 divergent pairs needing a human ruling
- **A3** — 95 `Thumbs.db` (88 MB), 2 `.tmp`, 1 `desktop.ini`
- **A5** — 18 folder renames (`data/rename_manifest.csv`); must go through `renameFolder()` in `drive.ts`, never Explorer. PP54 flagged for review
- **A6** — PP104 gap; **A7** — PP118/PP319 both "212 Imad Ct" (CRM call)
- **PP5's `EstimateReport (4).xls`** — 13 months newer than P2's copy

**These all touch files, and the `os.rmdir` safety net does not extend to them.**
`os.remove` deletes whatever it is pointed at. Before any of it runs, the tooling
needs: move-to-quarantine instead of delete, per-file hashes recorded before and
after, and a hard refusal on any divergence. Do not reuse the folder tooling here.
