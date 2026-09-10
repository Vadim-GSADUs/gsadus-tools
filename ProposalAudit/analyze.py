#!/usr/bin/env python3
"""
analyze.py -- questions we ask of the corpus scanned by scan_proposals.py.

Every command is a query over data/audit.db. Nothing here touches Drive, so it is
safe to run repeatedly and cheap to extend: a new question is a new subcommand,
not a new scan.

  python analyze.py summary                 corpus totals + the bloat ratio
  python analyze.py occupancy               every folder path, ranked by real use
  python analyze.py occupancy --under "1 - Takeoff & Estimate"
  python analyze.py never                   paths that have NEVER held a file
  python analyze.py outliers                proposals that break the mould
  python analyze.py dupes                   '(1)' folders = copy-paste accidents

"Occupied" means n_deep > 0: the folder, or something beneath it, holds at least
one real file. That is the only defensible definition of "used" available from the
filesystem -- mtime on a Drive mount is sync state, not human activity (see README).
"""
from __future__ import annotations

import argparse
import os
import sqlite3
import sys

DEFAULT_DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "audit.db")
DIV_ROOT = "1 - Takeoff & Estimate/3rd Party Coordination"


def connect(path):
    if not os.path.exists(path):
        print("ERROR: no db at %s -- run scan_proposals.py first" % path, file=sys.stderr)
        raise SystemExit(2)
    return sqlite3.connect(path)


def n_proposals(db):
    return db.execute("SELECT COUNT(*) FROM proposals").fetchone()[0]


def cmd_summary(db, args):
    n = n_proposals(db)
    d, e, f, b = db.execute(
        "SELECT SUM(n_dirs), SUM(n_empty_dirs), SUM(n_files), SUM(bytes) FROM proposals"
    ).fetchone()
    paths = db.execute(
        "SELECT COUNT(DISTINCT relpath) FROM nodes WHERE is_dir=1 AND relpath<>''").fetchone()[0]
    used = db.execute(
        "SELECT COUNT(*) FROM (SELECT relpath FROM nodes WHERE is_dir=1 AND relpath<>'' "
        "GROUP BY relpath HAVING SUM(CASE WHEN n_deep>0 THEN 1 ELSE 0 END) > 0)").fetchone()[0]
    print("proposals scanned      %d" % n)
    print("directories on Drive   %d" % d)
    print("  ...empty             %d  (%.1f%%)" % (e, 100.0 * e / d))
    print("files                  %d" % f)
    print("bytes                  %.1f GB" % ((b or 0) / 1e9))
    print()
    print("distinct folder paths  %d" % paths)
    print("  ...ever hold a file  %d" % used)
    print("  ...NEVER used        %d  (%.1f%% of the design)"
          % (paths - used, 100.0 * (paths - used) / paths))


def cmd_occupancy(db, args):
    n = n_proposals(db)
    where, params = "is_dir=1 AND relpath<>''", []
    if args.under:
        where += " AND (relpath = ? OR relpath LIKE ?)"
        params = [args.under, args.under + "/%"]
    rows = db.execute(
        "SELECT relpath, COUNT(*), SUM(CASE WHEN n_deep>0 THEN 1 ELSE 0 END), SUM(n_deep) "
        "FROM nodes WHERE " + where + " GROUP BY relpath ORDER BY 3 DESC, 1", params).fetchall()
    print("%-6s %-8s %8s  %s" % ("used", "present", "files", "path"))
    print("%-6s %-8s %8s  %s" % ("-" * 6, "-" * 7, "-" * 8, "-" * 60))
    for rp, present, occ, files in rows:
        pct = 100.0 * (occ or 0) / n
        print("%4d%-2s %-8d %8d  %s"
              % (occ or 0, "" if pct >= 1 else " !", present, files or 0, rp))
    print("\n%d paths;  '!' = used by under 1%% of the %d proposals" % (len(rows), n))


def cmd_never(db, args):
    rows = db.execute(
        "SELECT relpath, COUNT(*) FROM nodes WHERE is_dir=1 AND relpath<>'' "
        "GROUP BY relpath HAVING SUM(CASE WHEN n_deep>0 THEN 1 ELSE 0 END) = 0 "
        "ORDER BY 2 DESC, 1").fetchall()
    total = sum(r[1] for r in rows)
    print("%d folder paths have NEVER held a file in ANY proposal." % len(rows))
    print("They exist as %d real folders on Drive.\n" % total)
    for rp, cnt in rows:
        print("  %5d x  %s" % (cnt, rp))


def cmd_outliers(db, args):
    mode = db.execute(
        "SELECT n_dirs FROM proposals GROUP BY n_dirs ORDER BY COUNT(*) DESC LIMIT 1"
    ).fetchone()[0]
    print("modal shape: %d dirs. Proposals that differ:\n" % mode)
    print("%-42s %6s %6s %6s" % ("proposal", "dirs", "files", "empty"))
    for p, d, f, e in db.execute(
            "SELECT proposal, n_dirs, n_files, n_empty_dirs FROM proposals "
            "WHERE n_dirs <> ? ORDER BY ABS(n_dirs - ?) DESC", (mode, mode)):
        print("%-42s %6d %6d %6d" % (p[:42], d, f, e))
    print("\nmissing the DIV tree entirely:")
    for (p,) in db.execute(
            "SELECT proposal FROM proposals WHERE proposal NOT IN "
            "(SELECT proposal FROM nodes WHERE relpath = ?) ORDER BY pp", (DIV_ROOT,)):
        print("  " + p)


def cmd_dupes(db, args):
    """'(1)' is Drive's collision suffix -- something was pasted onto itself.

    Folders and files are separate problems and must not be conflated: a duplicate
    FOLDER forks the structure, while a duplicate FILE is a document-integrity
    question (a second copy of a signed change order is not cosmetic).
    """
    folders = db.execute(
        "SELECT proposal, COUNT(*) FROM nodes WHERE is_dir=1 AND relpath LIKE '%(1)%' "
        "GROUP BY proposal ORDER BY 2 DESC").fetchall()
    print("=== duplicate FOLDERS (structure forked) -- %d proposal(s) ===" % len(folders))
    for p, c in folders:
        print("  %-44s %3d dir(s)" % (p[:44], c))

    files = db.execute(
        "SELECT proposal, name FROM nodes WHERE is_dir=0 AND name LIKE '%(1)%' "
        "ORDER BY proposal, name").fetchall()
    seen = sorted({p for p, _ in files})
    print("\n=== duplicate FILES -- %d file(s) across %d proposal(s) ==="
          % (len(files), len(seen)))
    for p, n in files:
        print("  %-40s %s" % (p[:40], n))


COMMANDS = {
    "summary": cmd_summary, "occupancy": cmd_occupancy, "never": cmd_never,
    "outliers": cmd_outliers, "dupes": cmd_dupes,
}


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("command", choices=sorted(COMMANDS))
    ap.add_argument("--db", default=DEFAULT_DB)
    ap.add_argument("--under", default="", help="occupancy: restrict to this subtree")
    args = ap.parse_args()
    COMMANDS[args.command](connect(args.db), args)
    return 0


if __name__ == "__main__":
    sys.exit(main())
