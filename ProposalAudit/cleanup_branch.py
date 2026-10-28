#!/usr/bin/env python3
"""
cleanup_branch.py -- remove or rename dead folders in every proposal that has them.

Four modes:

  WHOLE BRANCH (default) -- remove the branch itself, but only where it is
  completely empty. A branch holding even one file is skipped and reported.

    python cleanup_branch.py --branch "1 - Takeoff & Estimate/3rd Party Coordination"

  PRUNE (--prune-empty) -- keep the branch folder as a named placeholder and
  remove only the empty directories BENEATH it. A proposal that holds real
  content keeps exactly the sub-folders that hold it and loses the dead ones.

    python cleanup_branch.py --branch "4 - Administrative & Change-Issues" --prune-empty

  RENAME (--rename-to NAME) -- give the branch folder a new name, in the template and
  every proposal, keeping its Drive ID. A branch holding a file is skipped, unless
  --move-files-to DEST is given: then every file beneath it moves into the proposal's
  DEST folder, and a second run, once Drive has synced, drops the emptied subfolders
  and renames the folder, which starts empty.
  --create-missing also makes the folder, under its new name, where neither name
  exists. --undo LOG reverses a run from its log.

    python cleanup_branch.py --branch "1 - Takeoff & Estimate" --rename-to "1 - CRM"
    python cleanup_branch.py --branch "4 - Administrative & Change-Issues" \\
        --rename-to "4 - Resources" --move-files-to "2 - Proposal - Contract"

  FLATTEN (--flatten) -- keep the branch folder and drop every subfolder beneath it:
  files in a subfolder move up into the branch folder, keeping their Drive IDs, and a
  second run, once Drive has synced, drops the emptied subfolders.
  A proposal where a lifted file would meet one of the same name is skipped. Logged
  step by step like a rename, so --undo reverses it too.

    python cleanup_branch.py --branch "3 - Supporting Documents/C - Received from Client" --flatten

WHY RENAME RUNS ON THE MOUNT (owner, 2026-10-27): the rule that renames go through
WebApp's renameFolder() protects the proposal folders' own names, which the engine's
reconciler and collision probe read. Nothing reads the subfolder names this mode
renames, and Drive for Desktop syncs a local rename as a rename of the same Drive
item, so IDs, links and the engine's gsadus_template_item tags survive. os.rename
refuses an existing target, the same kind of guarantee os.rmdir gives the removals.
The result is verified through the Drive API after sync (scan_proposals.py --drive),
ID by ID, not trusted from the mount.

WHY TWO RUNS (2026-10-28): a run that moves files never also removes the folders they
came from. On the P-side flatten, 5 of the 6 files lifted out of a nested folder (P30
C/Documents/HOA Documents) reached their new folder in Drive and were then trashed with
it: Drive for Desktop synced the folder removal before it had settled the moves. The
ID-by-ID check after sync caught it, and the files came back from the trash. So when
files must move, a run moves only them; once the drive-wide poll shows the moves synced,
the same command again finds empty subfolders and removes them (and, in rename mode,
renames). Always verify ID by ID after sync.

WHY PRUNE EXISTS: on 2026-09-10 the whole-branch mode removed "1 - Takeoff &
Estimate" itself, and the owner wanted that folder kept as a placeholder. It had
to be recovered by re-running the two child branches separately. Enumerating
children by hand only works when you already know every child every proposal has
-- prune derives that per proposal instead of trusting a list someone typed.

The emptiness rule IS the keep-list in both modes. Nothing is hardcoded as
protected because nothing needs to be: the six proposals with real estimate
content (PP1, PP2, PP3, PP5, PP6, PP110) survived Tier 2 by the same rule that
made Tier 1 safe -- not by a name I remembered to type.

Executed so far (README "What was changed"; the September evidence is in git history):
  2026-09-10  Tier 1  "1 - Takeoff & Estimate/3rd Party Coordination"  26,014 folders
  2026-09-10  Tier 2  "1 - Takeoff & Estimate"                          1,059 folders
  2026-09-10  fix     re-ran both children after the parent was restored 26,430 folders
  2026-09-10  Tier 3  "4 - Administrative & Change-Issues" --prune-empty
  2026-10-27  Batch 1 --rename-to "1 - CRM" and "4 - Resources"           763 renames
  2026-10-28  Batch 5 "3 - Supporting Documents/C - Received from Client" --flatten
Kept for Batch 4 (retire "B - Questionnaire & Selection", held) and for --undo of the
Batch 1 and Batch 5 logs in ~/.gsadus/proposal-folder-renames/.

SAFETY -- three independent layers, deliberately redundant:

 1. os.rmdir bottom-up, NEVER shutil.rmtree. The operating system refuses to remove
    a directory that still contains anything, so a file cannot be destroyed by this
    script even if layers 2 and 3 were both wrong. This is the layer that matters:
    it does not depend on my reasoning being correct.
 2. Every subtree is re-walked live immediately before deletion and skipped if it
    holds even one file. The audit db is evidence, not authority -- the corpus may
    have changed since it was scanned.
 3. An explicit exclusion list, and a refusal on any '0 - P*.lnk' project link
    (those encode signed/onboarded status and must never be touched).

Deletions land in the shared drive's trash via the Drive for Desktop mount, so they
are recoverable for 30 days, and they are attributed to the signed-in owner -- a
human decided this, and the audit log should say so.

Run the TEMPLATE first (it is included here). walkTemplateTree() reads the template
live at mint time, so if the engine ever re-runs hydration to repair a proposal it
re-creates missing items FROM THE TEMPLATE. Clean proposals before the template and
a single repair puts the DIV tree straight back.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from datetime import datetime, timezone

DEFAULT_ROOT = r"G:\Shared drives\GSADUs Business\1 - Proposals"
DATA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
# Every log names proposal folders, i.e. client addresses: data/ is gitignored.
RENAME_LOG = os.path.join(DATA, "folder_rename_log.jsonl")
FLATTEN_LOG = os.path.join(DATA, "folder_flatten_log.jsonl")
NUM_RE = re.compile(r"^(PP|P)(\d+)\b")   # proposal PP221 or project P28

# Proposals this pass must not touch. PP358 was held out on 2026-09-10 while its
# forked structure was under mint-bug investigation; that closed the same day and the
# owner resolved PP1 and PP358 by hand, so the set is empty. Add a name here rather
# than skipping by memory -- an exclusion that lives only in someone's head is not one.
EXCLUDE = set()


def survey(path):
    """Walk a subtree. Returns (dirs, files, protected) without modifying anything."""
    dirs, files, protected = [], 0, []
    for dirpath, dirnames, filenames in os.walk(path):
        dirs.append(dirpath)
        files += len(filenames)
        for f in filenames:
            if f.lower().startswith("0 - p") and f.lower().endswith(".lnk"):
                protected.append(os.path.join(dirpath, f))
    return dirs, files, protected


def survey_prune(path):
    """Post-order walk for --prune-empty. Returns (removable, files, protected).

    `removable` is every directory strictly BELOW `path` whose subtree holds zero
    files. The branch root itself is never included -- keeping it is the whole
    point of this mode. A parent that qualifies implies its children qualify too;
    they are all listed and remove_bottom_up() takes them deepest-first, so the
    OS still sees an empty directory at every single rmdir.
    """
    removable, protected, total = [], [], 0

    def rec(p, is_root):
        nonlocal total
        deep = 0
        with os.scandir(p) as it:
            for e in it:
                if e.is_dir(follow_symlinks=False):
                    deep += rec(e.path, False)
                else:
                    deep += 1
                    total += 1
                    lo = e.name.lower()
                    if lo.startswith("0 - p") and lo.endswith(".lnk"):
                        protected.append(e.path)
        if deep == 0 and not is_root:
            removable.append(p)
        return deep

    rec(path, True)
    return removable, total, protected


def remove_bottom_up(dirs):
    """rmdir deepest-first. Returns (removed, failures).

    rmdir on a non-empty directory raises OSError -- that is the point. We never
    force, never recurse-delete, and a failure here is a signal to stop and look.
    """
    removed, failures = 0, []
    for d in sorted(dirs, key=lambda p: p.count(os.sep), reverse=True):
        try:
            os.rmdir(d)
            removed += 1
        except OSError as ex:
            failures.append((d, str(ex)))
    return removed, failures


# --- rename mode ------------------------------------------------------------------

def label(name):
    """'PP221' or 'P28' for a proposal or project folder: the report never needs the
    street address."""
    m = NUM_RE.match(name)
    return m.group(1) + m.group(2) if m else ("template" if "template" in name.lower() else name)


def ranges(names):
    """['PP1 …', 'PP2 …', 'PP3 …', 'PP7 …'] -> 'PP1-PP3, PP7' (numbered folders only)"""
    out = []
    for prefix in ("PP", "P"):
        nums = sorted(int(m.group(2)) for m in map(NUM_RE.match, names) if m and m.group(1) == prefix)
        i = 0
        while i < len(nums):
            j = i
            while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
                j += 1
            out.append("%s%d" % (prefix, nums[i]) if i == j
                       else "%s%d-%s%d" % (prefix, nums[i], prefix, nums[j]))
            i = j + 1
    return ", ".join(out)


def walk_branch(path):
    """Every file and every subdirectory strictly beneath `path`, plus protected links."""
    files, subdirs, protected = [], [], []
    for dirpath, dirnames, filenames in os.walk(path):
        subdirs.extend(os.path.join(dirpath, d) for d in dirnames)
        for f in filenames:
            files.append(os.path.join(dirpath, f))
            if f.lower().startswith("0 - p") and f.lower().endswith(".lnk"):
                protected.append(os.path.join(dirpath, f))
    return files, subdirs, protected


def plan_rename(base, branch, new_name, move_to, create_missing):
    """What renaming `branch` to `new_name` means for one proposal folder. Reads only.

    Returns (kind, detail, steps): kind is rename | create | already | absent | skip;
    steps is the ordered list of (action, src, dst) to carry it out.
    """
    old = os.path.join(base, branch)
    new = os.path.join(os.path.dirname(old), new_name)
    has_old, has_new = os.path.isdir(old), os.path.isdir(new)
    if has_old and has_new:
        return "skip", "both names present", []
    if not has_old:
        if has_new:
            return "already", "", []
        if create_missing and os.path.isdir(os.path.dirname(new)):
            return "create", "", [("mkdir", None, new)]
        return "absent", "", []
    files, subdirs, protected = walk_branch(old)
    if protected:
        return "skip", "holds %d protected project link(s)" % len(protected), []
    steps = []
    if files:
        if not move_to:
            return "skip", "holds %d file(s)" % len(files), []
        dest = os.path.join(base, move_to)
        if not os.path.isdir(dest):
            return "skip", "no '%s' to move its files into" % move_to, []
        targets = [os.path.join(dest, os.path.basename(f)) for f in files]
        clash = [t for t in targets if os.path.exists(t)]
        if clash or len(set(t.lower() for t in targets)) < len(targets):
            return "skip", "a file of the same name is already in '%s'" % move_to, []
        # Moves only, this run: see WHY TWO RUNS.
        return ("rename", "%d file(s) to '%s' first; run again once Drive has synced"
                % (len(files), move_to), [("move", f, t) for f, t in zip(files, targets)])
    if subdirs and move_to:
        steps += [("rmdir", d, None) for d in sorted(subdirs, key=lambda p: p.count(os.sep),
                                                     reverse=True)]
    steps.append(("rename", old, new))
    detail = "carries %d empty subfolder(s)" % len(subdirs) if subdirs and not move_to else ""
    return "rename", detail, steps


def run_steps(steps, name, log):
    """Carry out one proposal's steps in order; stop at the first failure.

    os.rename refuses an existing target and os.rmdir a non-empty directory, so a
    stale plan fails loudly instead of overwriting or deleting anything."""
    for action, src, dst in steps:
        try:
            if action == "move" or action == "rename":
                if os.path.exists(dst):
                    raise FileExistsError("target exists: " + dst)
                os.rename(src, dst)
            elif action == "rmdir":
                os.rmdir(src)
            elif action == "mkdir":
                os.mkdir(dst)
            ok, err = True, ""
        except OSError as ex:
            ok, err = False, str(ex)
        log.write(json.dumps({
            "ts": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "proposal": name, "action": action, "src": src, "dst": dst,
            "ok": ok, "error": err}) + "\n")
        log.flush()
        if not ok:
            return False, "%s failed: %s" % (action, err)
    return True, ""


def plan_flatten(base, branch):
    """What flattening `branch` means for one proposal folder. Reads only.

    Returns (kind, detail, steps) like plan_rename; kind is flatten | already | absent |
    skip. Files already directly in the branch folder stay where they are.
    """
    root = os.path.join(base, branch)
    if not os.path.isdir(root):
        return "absent", "", []
    files, subdirs, protected = walk_branch(root)
    if protected:
        return "skip", "holds %d protected project link(s)" % len(protected), []
    if not subdirs:
        return "already", "", []
    lifted = [f for f in files if os.path.dirname(f) != root]
    targets = [os.path.join(root, os.path.basename(f)) for f in lifted]
    clash = [t for t in targets if os.path.exists(t)]
    if clash or len(set(t.lower() for t in targets)) < len(targets):
        return "skip", "a lifted file would meet one of the same name", []
    if lifted:  # moves only, this run: see WHY TWO RUNS
        return ("flatten", "%d file(s) up first; run again once Drive has synced" % len(lifted),
                [("move", f, t) for f, t in zip(lifted, targets)])
    return "flatten", "", [("rmdir", d, None) for d in sorted(subdirs, key=lambda p: p.count(os.sep),
                                                              reverse=True)]


def rename_main(args, branch, names):
    new_name = args.rename_to.strip()
    if not new_name or os.sep in new_name or "/" in new_name:
        print("ERROR: --rename-to takes a folder name, not a path", file=sys.stderr)
        return 2
    title = ["rename '%s'  ->  '%s'" % (branch, new_name)]
    if args.move_files_to:
        title.append("files beneath it move to '%s' first" % args.move_files_to)
    labels = [("rename", "would rename", "renamed")]
    if args.create_missing:
        labels.append(("create", "would create", "created"))
    labels += [("already", "already renamed", "already renamed"),
               ("skip", "SKIPPED (kept as is)", "SKIPPED (kept as is)"),
               ("absent", "neither name present", "neither name present")]
    return steps_main(args, names, title, labels,
                      lambda base: plan_rename(base, branch, new_name, args.move_files_to,
                                               args.create_missing))


def flatten_main(args, branch, names):
    title = ["flatten '%s': files move up, subfolders go" % branch]
    labels = [("flatten", "would flatten", "flattened"),
              ("already", "no subfolders", "no subfolders"),
              ("skip", "SKIPPED (kept as is)", "SKIPPED (kept as is)"),
              ("absent", "branch absent", "branch absent")]
    return steps_main(args, names, title, labels, lambda base: plan_flatten(base, branch))


def steps_main(args, names, title, labels, plan):
    """Run a step-planned mode (rename, flatten) over the template and every proposal.

    `plan(base)` gives one proposal folder's (kind, detail, steps); `labels` lists each
    kind's (kind, dry-run line, applied line) in report order."""
    print("=" * 74)
    for t in title:
        print(t)
    print("APPLY -- changing Drive" if args.apply else "DRY RUN -- nothing will change")
    print("=" * 74)

    # The template first: a mint during the pass then already gets the new shape.
    names = sorted(names, key=lambda n: (NUM_RE.match(n) is not None, n))
    log = open(args.manifest, "a", encoding="utf-8") if args.apply else None
    groups = {k: [] for k in [kind for kind, _, _ in labels] + ["failed"]}
    notes = []
    t0 = time.perf_counter()

    def one_pass(batch):
        for name in batch:
            if name in EXCLUDE:
                groups["skip"].append(name)
                notes.append((name, "excluded"))
                continue
            try:
                kind, detail, steps = plan(os.path.join(args.root, name))
            except OSError as ex:
                kind, detail, steps = "skip", "unreadable: %s" % str(ex)[:60], []
            if args.apply and steps:
                ok, err = run_steps(steps, name, log)
                if not ok:
                    kind, detail = "failed", err
            groups[kind].append(name)
            if detail or kind in ("skip", "failed") or not NUM_RE.match(name):
                notes.append((name, "%s%s" % (kind, ": " + detail if detail else "")))
            if any(action == "move" for action, _, _ in steps):  # files move: step by step
                for action, src, dst in steps:
                    shown = os.path.relpath(src or dst, os.path.join(args.root, name))
                    print("  %-10s %-6s %s%s" % (label(name), action, shown,
                                                 "  ->  " + os.path.relpath(
                                                     dst, os.path.join(args.root, name))
                                                 if action in ("move", "rename") else ""))

    one_pass(names)
    if args.apply:
        # A proposal minted while the pass ran is not in `names`: pick it up now.
        with os.scandir(args.root) as it:
            fresh = sorted(e.name for e in it if e.is_dir(follow_symlinks=False)
                           and e.name not in set(names))
        if fresh:
            print("  minted during the pass: %s" % ", ".join(label(n) for n in fresh))
            one_pass(fresh)
        log.close()

    def line(title, group):
        """Folders by PP/P number; anything else (the PP template) named after them."""
        numbered = [n for n in groups[group] if NUM_RE.match(n)]
        other = [label(n) for n in groups[group] if not NUM_RE.match(n)]
        print("%-24s %4d folders  %s%s" % (title, len(numbered), ranges(numbered),
                                           "  + " + ", ".join(other) if other else ""))

    print("\n" + "-" * 74)
    for name, note in notes:
        print("  %-10s %s" % (label(name), note))
    print("-" * 74)
    for kind, dry, applied in labels:
        line(applied if args.apply else dry, kind)
    if args.apply:
        line("FAILED", "failed")
        print("log: %s" % args.manifest)
    print("elapsed %.1fs" % (time.perf_counter() - t0))
    if not args.apply:
        print("\nDry run only. Re-run with --apply to execute.")
    return 1 if groups["failed"] else 0


def undo_main(args):
    """Reverse a rename run from its log, newest step first. Dry run unless --apply.

    A removed subfolder comes back as a NEW empty folder (new Drive ID); the original
    sits in the shared drive's trash for 30 days if its identity matters."""
    with open(args.undo, encoding="utf-8") as fh:
        entries = [json.loads(line) for line in fh if line.strip()]
    print("undo %d logged step(s) from %s -- %s"
          % (len(entries), args.undo, "APPLY" if args.apply else "DRY RUN"))
    bad = 0
    for e in reversed([e for e in entries if e.get("ok")]):
        act, src, dst = e["action"], e["src"], e["dst"]
        if act in ("rename", "move"):
            what, ready = ("rename back", os.path.exists(dst) and not os.path.exists(src))
        elif act == "rmdir":
            what, ready = ("recreate", not os.path.exists(src))
        else:
            what, ready = ("remove", os.path.isdir(dst) and not os.listdir(dst))
        print("  %-11s %-10s %s%s" % (what, label(e["proposal"]),
                                      os.path.relpath(dst or src, args.root),
                                      "" if ready else "   !! state changed, skipped"))
        if not ready:
            bad += 1
            continue
        if args.apply:
            if act in ("rename", "move"):
                os.rename(dst, src)
            elif act == "rmdir":
                os.mkdir(src)
            else:
                os.rmdir(dst)
    print("%d step(s) could not be undone as logged" % bad if bad else "all steps reversible")
    return 1 if bad else 0


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--branch",
                    help="branch to remove or rename, relative to each proposal, e.g. "
                         "'1 - Takeoff & Estimate'. Forward slashes are fine.")
    ap.add_argument("--prune-empty", action="store_true",
                    help="keep the branch folder itself and remove only the empty "
                         "directories beneath it")
    ap.add_argument("--rename-to", default="",
                    help="rename the branch folder to this name instead of removing it")
    ap.add_argument("--move-files-to", default="",
                    help="rename mode: first move every file beneath the branch into "
                         "this folder of the same proposal, and drop the emptied subfolders")
    ap.add_argument("--create-missing", action="store_true",
                    help="rename mode: make the folder under its new name where neither "
                         "name exists")
    ap.add_argument("--flatten", action="store_true",
                    help="keep the branch folder, lift every file beneath it into it, and "
                         "drop its subfolders")
    ap.add_argument("--undo", default="", help="reverse a rename or flatten run from its log")
    ap.add_argument("--root", default=DEFAULT_ROOT)
    ap.add_argument("--apply", action="store_true",
                    help="actually change Drive. Without this the script only reports.")
    ap.add_argument("--limit", type=int, default=0, help="process only the first N")
    ap.add_argument("--only", default="", help="substring match on proposal name")
    ap.add_argument("--manifest", default="",
                    help="log file (default: data/cleanup_manifest.jsonl; rename mode "
                         "data/folder_rename_log.jsonl)")
    args = ap.parse_args()

    if not os.path.isdir(args.root):
        print("ERROR: root not reachable: " + args.root, file=sys.stderr)
        return 2
    if args.undo:
        return undo_main(args)
    if not args.branch:
        ap.error("--branch is required")
    if (args.move_files_to or args.create_missing) and not args.rename_to:
        ap.error("--move-files-to and --create-missing need --rename-to")
    if sum(map(bool, (args.rename_to, args.prune_empty, args.flatten))) > 1:
        ap.error("--rename-to, --prune-empty and --flatten are separate passes")
    args.manifest = args.manifest or (RENAME_LOG if args.rename_to else
                                      FLATTEN_LOG if args.flatten else
                                      os.path.join(DATA, "cleanup_manifest.jsonl"))
    os.makedirs(os.path.dirname(args.manifest), exist_ok=True)

    branch = args.branch.replace("/", os.sep).strip(os.sep)

    with os.scandir(args.root) as it:
        names = sorted(e.name for e in it if e.is_dir(follow_symlinks=False))
    if args.only:
        names = [n for n in names if args.only.lower() in n.lower()]
    if args.limit:
        names = names[:args.limit]
    if args.rename_to:
        return rename_main(args, branch, names)
    if args.flatten:
        return flatten_main(args, branch, names)

    mode = "APPLY -- deleting" if args.apply else "DRY RUN -- nothing will change"
    print("=" * 74)
    print("remove branch: '%s'" % branch)
    print(mode)
    print("=" * 74)

    log = open(args.manifest, "a", encoding="utf-8") if args.apply else None

    stats = {"targeted": 0, "absent": 0, "excluded": 0, "skipped_files": 0,
             "dirs": 0, "removed": 0, "failed": 0, "nothing_to_prune": 0,
             "unreadable": 0}
    blocked, failed_detail, partial = [], [], []
    t0 = time.perf_counter()

    for name in names:
        if name in EXCLUDE:
            stats["excluded"] += 1
            print("  EXCLUDED  %s  (live mint bug, separate investigation)" % name)
            continue
        path = os.path.join(args.root, name, branch)
        if not os.path.isdir(path):
            stats["absent"] += 1
            continue

        try:
            if args.prune_empty:
                dirs, files, protected = survey_prune(path)
            else:
                dirs, files, protected = survey(path)
        except OSError as ex:
            # An unreadable subtree is a reason to stop looking at this proposal,
            # never a reason to delete from a partial picture.
            stats["unreadable"] += 1
            blocked.append((name, -1, 0))
            print("  !! SKIP    %-44s unreadable: %s" % (name[:44], str(ex)[:60]))
            continue

        if protected:
            stats["skipped_files"] += 1
            blocked.append((name, files, len(protected)))
            print("  !! SKIP    %-44s holds %d PROTECTED link(s)" % (name[:44], len(protected)))
            continue

        if args.prune_empty:
            if not dirs:
                stats["nothing_to_prune"] += 1
                continue
            if files:
                # Partial prune: this proposal keeps real content and loses only the
                # dead folders around it. Worth naming in the report -- it is the one
                # case where a deletion happens inside a populated subtree.
                partial.append((name, files, len(dirs)))
        elif files:
            stats["skipped_files"] += 1
            blocked.append((name, files, 0))
            print("  !! SKIP    %-44s holds %d file(s) -- NOT EMPTY" % (name[:44], files))
            continue

        stats["targeted"] += 1
        stats["dirs"] += len(dirs)

        if not args.apply:
            continue

        removed, failures = remove_bottom_up(dirs)
        stats["removed"] += removed
        if failures:
            stats["failed"] += 1
            failed_detail.extend(failures)
        log.write(json.dumps({
            "ts": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "proposal": name, "branch": branch,
            "mode": "prune-empty" if args.prune_empty else "whole-branch",
            "dirs_found": len(dirs), "dirs_removed": removed,
            "files_kept": files, "failures": len(failures), "ok": not failures,
        }) + "\n")
        log.flush()
        if stats["targeted"] % 50 == 0:
            print("  ... %d proposals, %d folders removed" % (stats["targeted"], stats["removed"]))

    if log:
        log.close()

    el = time.perf_counter() - t0
    print("\n" + "-" * 74)
    print("mode                              : %s"
          % ("PRUNE (branch folder kept)" if args.prune_empty else "whole branch"))
    print("proposals with folders to remove  : %d" % stats["targeted"])
    print("branch absent                     : %d" % stats["absent"])
    print("excluded                          : %d" % stats["excluded"])
    if args.prune_empty:
        print("nothing to prune (already clean)  : %d" % stats["nothing_to_prune"])
        print("partial prunes (kept real content): %d" % len(partial))
    else:
        print("SKIPPED because not empty         : %d" % stats["skipped_files"])
    print("SKIPPED, protected link present   : %d" % sum(1 for b in blocked if b[2]))
    print("SKIPPED, unreadable               : %d" % stats["unreadable"])
    print("folders in scope                  : %d" % stats["dirs"])
    if args.apply:
        print("folders REMOVED                   : %d" % stats["removed"])
        print("proposals with failures           : %d" % stats["failed"])
        print("manifest                          : %s" % args.manifest)
    print("elapsed                           : %.1fs" % el)

    if partial:
        print("\nPARTIAL prunes -- these kept real content, dead folders removed around it:")
        for n, f, d in partial:
            print("   %-46s kept %d file(s), removed %d empty dir(s)" % (n[:46], f, d))
    if blocked:
        print("\nNOT TOUCHED (investigate before any further pass):")
        for n, f, p in blocked:
            print("   %-46s %s%s" % (n[:46], "unreadable" if f < 0 else "%d file(s)" % f,
                                     ", %d PROTECTED link(s)" % p if p else ""))
    if failed_detail:
        print("\nrmdir failures (first 15) -- these directories were NOT empty:")
        for d, e in failed_detail[:15]:
            print("   %s\n      %s" % (d, e))

    if not args.apply:
        print("\nDry run only. Re-run with --apply to execute.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
