#!/usr/bin/env python3
"""
cleanup_branch.py -- remove dead folders from every proposal that has them.

Two modes:

  WHOLE BRANCH (default) -- remove the branch itself, but only where it is
  completely empty. A branch holding even one file is skipped and reported.

    python cleanup_branch.py --branch "1 - Takeoff & Estimate/3rd Party Coordination"

  PRUNE (--prune-empty) -- keep the branch folder as a named placeholder and
  remove only the empty directories BENEATH it. A proposal that holds real
  content keeps exactly the sub-folders that hold it and loses the dead ones.

    python cleanup_branch.py --branch "4 - Administrative & Change-Issues" --prune-empty

WHY PRUNE EXISTS: on 2026-09-10 the whole-branch mode removed "1 - Takeoff &
Estimate" itself, and the owner wanted that folder kept as a placeholder. It had
to be recovered by re-running the two child branches separately. Enumerating
children by hand only works when you already know every child every proposal has
-- prune derives that per proposal instead of trusting a list someone typed.

The emptiness rule IS the keep-list in both modes. Nothing is hardcoded as
protected because nothing needs to be: the six proposals with real estimate
content (PP1, PP2, PP3, PP5, PP6, PP110) survived Tier 2 by the same rule that
made Tier 1 safe -- not by a name I remembered to type.

Executed so far (see README for the evidence behind each):
  2026-09-10  Tier 1  "1 - Takeoff & Estimate/3rd Party Coordination"  26,014 folders
  2026-09-10  Tier 2  "1 - Takeoff & Estimate"                          1,059 folders
  2026-09-10  fix     re-ran both children after the parent was restored 26,430 folders
  2026-09-10  Tier 3  "4 - Administrative & Change-Issues" --prune-empty

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
import sys
import time
from datetime import datetime, timezone

DEFAULT_ROOT = r"G:\Shared drives\GSADUs Business\1 - Proposals"

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


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--branch", required=True,
                    help="branch to remove, relative to each proposal, e.g. "
                         "'1 - Takeoff & Estimate'. Forward slashes are fine.")
    ap.add_argument("--prune-empty", action="store_true",
                    help="keep the branch folder itself and remove only the empty "
                         "directories beneath it")
    ap.add_argument("--root", default=DEFAULT_ROOT)
    ap.add_argument("--apply", action="store_true",
                    help="actually delete. Without this the script only reports.")
    ap.add_argument("--limit", type=int, default=0, help="process only the first N")
    ap.add_argument("--only", default="", help="substring match on proposal name")
    ap.add_argument("--manifest", default=os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "data", "cleanup_manifest.jsonl"))
    args = ap.parse_args()

    branch = args.branch.replace("/", os.sep).strip(os.sep)

    if not os.path.isdir(args.root):
        print("ERROR: root not reachable: " + args.root, file=sys.stderr)
        return 2

    with os.scandir(args.root) as it:
        names = sorted(e.name for e in it if e.is_dir(follow_symlinks=False))
    if args.only:
        names = [n for n in names if args.only.lower() in n.lower()]
    if args.limit:
        names = names[:args.limit]

    mode = "APPLY -- deleting" if args.apply else "DRY RUN -- nothing will change"
    print("=" * 74)
    print("remove branch: '%s'" % branch)
    print(mode)
    print("=" * 74)

    os.makedirs(os.path.dirname(args.manifest), exist_ok=True)
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
