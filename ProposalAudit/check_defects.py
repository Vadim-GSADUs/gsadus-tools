#!/usr/bin/env python3
"""
check_defects.py -- the Tier A housekeeping worklist, reported not executed.

READ-ONLY BY CONSTRUCTION. Nothing here renames, moves, or deletes anything. It
reads data/audit.db, and for duplicate detection it reads file BYTES off the Drive
mount to hash them -- that is the only Drive access, and it is a read.

Remediation belongs in a separate script with an explicit --apply and a manifest,
so that the tool which *finds* problems can never be the tool which acts on them.
That separation is the whole reason this file is safe to run at any time.

  python check_defects.py all
  python check_defects.py dupes       A2  '(1)' files: identical / divergent / misnamed
  python check_defects.py junk        A3  Thumbs.db, desktop.ini, .tmp
  python check_defects.py projects    A4  PP<->P links (owner-minted, DO NOT DELETE)
  python check_defects.py names       A5  USPS rename manifest (--csv to write one)
  python check_defects.py registry    A6/A7  number gaps, duplicate addresses
  python check_defects.py shells      A8  proposals holding zero files
"""
from __future__ import annotations

import argparse
import collections
import csv
import hashlib
import os
import re
import sqlite3
import sys
from concurrent.futures import ThreadPoolExecutor

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DB = os.path.join(HERE, "data", "audit.db")
DEFAULT_ROOT = r"G:\Shared drives\GSADUs Business\1 - Proposals"

# USPS Publication 28 suffix abbreviations. Amendment A5 (2026-08-27) made these the
# canonical folder form and @gsadus/pipedrive 0.1.3 enforces them on new mints; the
# proposals listed by `names` predate that and are backfill, not a new rule.
USPS = {
    "Street": "St", "Avenue": "Ave", "Drive": "Dr", "Circle": "Cir", "Court": "Ct",
    "Road": "Rd", "Lane": "Ln", "Boulevard": "Blvd", "Place": "Pl", "Terrace": "Ter",
    "Trail": "Trl", "Parkway": "Pkwy", "Square": "Sq", "Highway": "Hwy",
}
DUP_RE = re.compile(r"^(.*) \((\d+)\)(\.[^.]*)?$")


def connect(path):
    if not os.path.exists(path):
        print("ERROR: no db at %s -- run scan_proposals.py first" % path, file=sys.stderr)
        raise SystemExit(2)
    return sqlite3.connect(path)


def head(title):
    print("\n" + "=" * 78)
    print(title)
    print("=" * 78)


def sha256(path):
    h = hashlib.sha256()
    try:
        with open(path, "rb") as fh:
            for chunk in iter(lambda: fh.read(1 << 20), b""):
                h.update(chunk)
        return h.hexdigest()
    except OSError as ex:
        return "ERR:" + str(ex)


# --------------------------------------------------------------------------- A2

def cmd_dupes(db, args):
    head("A2 -- '(1)' collision-suffixed files")
    rows = db.execute(
        "SELECT proposal, relpath, name, size FROM nodes "
        "WHERE is_dir=0 AND name LIKE '%(1)%' ORDER BY proposal, relpath").fetchall()
    sizes = {}
    for p, rp, n, s in db.execute("SELECT proposal, relpath, name, size FROM nodes WHERE is_dir=0"):
        sizes[(p, os.path.dirname(rp), n)] = (s, rp)

    paired, orphans = [], []
    for p, rp, n, s in rows:
        m = DUP_RE.match(n)
        if not m:
            continue
        base = m.group(1) + (m.group(3) or "")
        hit = sizes.get((p, os.path.dirname(rp), base))
        (paired if hit else orphans).append((p, rp, n, s, hit))

    same = [x for x in paired if x[4][0] == x[3]]
    diff = [x for x in paired if x[4][0] != x[3]]

    print("%d suffixed files: %d have a same-folder original, %d do not.\n"
          % (len(paired) + len(orphans), len(paired), len(orphans)))

    print("-- SAME SIZE as original (%d) -- hashing to confirm byte-identity" % len(same))
    if same:
        def job(x):
            p, rp, n, s, hit = x
            a = os.path.join(args.root, p, rp.replace("/", os.sep))
            b = os.path.join(args.root, p, hit[1].replace("/", os.sep))
            return x, sha256(a), sha256(b)
        with ThreadPoolExecutor(max_workers=8) as pool:
            results = list(pool.map(job, same))
        ident = [r for r in results if r[1] == r[2] and not r[1].startswith("ERR:")]
        notid = [r for r in results if r[1] != r[2]]
        by_prop = collections.Counter(r[0][0] for r in ident)
        for p, c in by_prop.items():
            print("   %-42s %3d byte-identical copies" % (p[:42], c))
        if notid:
            print("   !! same size but DIFFERENT bytes:")
            for x, ha, hb in notid:
                print("      %-38s %s" % (x[2][:38], x[0][:34]))
        print("   -> %d confirmed redundant, safe to remove once a human signs off" % len(ident))

    print("\n-- DIFFERENT SIZE from original (%d) -- NOT duplicates, need human eyes" % len(diff))
    for p, rp, n, s, hit in diff:
        print("   %-40s %9d B   vs original %9d B" % (n[:40], s, hit[0]))
        print("      %s / %s" % (p[:38], os.path.dirname(rp)))

    print("\n-- NO original in the folder (%d) -- misnamed singletons, not duplicates" % len(orphans))
    print("   The '(1)' is part of the filename (downloaded that way). Cosmetic at most.")
    for p, rp, n, s, _ in orphans:
        print("   %-46s %s" % (n[:46], p[:30]))


# --------------------------------------------------------------------------- A3

def cmd_junk(db, args):
    head("A3 -- OS litter")
    for n, c, b in db.execute(
            "SELECT name, COUNT(*), SUM(size) FROM nodes WHERE is_dir=0 AND "
            "(name IN ('Thumbs.db','desktop.ini','.DS_Store') OR ext='.tmp') "
            "GROUP BY name ORDER BY 2 DESC"):
        print("  %-24s %4d files  %8.1f KB" % (n, c, (b or 0) / 1024.0))
    print("\n  Safe to delete, but Thumbs.db regenerates whenever someone browses the")
    print("  folder in Explorer -- it is a symptom, not a one-time mess.")


# --------------------------------------------------------------------------- A4

LNK_RE = re.compile(r"^0 - P(\d+) (.*?)( - Cancelled)?\.lnk$", re.I)


def cmd_projects(db, args):
    """A4 -- the '0 - P<n>' shortcuts. NOT debt. DO NOT DELETE.

    These are deliberately minted by the owner when a proposal is signed and
    onboarded, linking the PP# proposal folder to its P# project folder. Their mere
    presence is the only won/lost signal that exists in the filesystem, and their
    names carry project status. Treating them as litter would destroy the single
    most useful piece of lifecycle data in the corpus.
    """
    head("A4 -- PP <-> P project links  (owner-minted; DO NOT DELETE)")
    rows = db.execute(
        "SELECT proposal, name FROM nodes WHERE is_dir=0 AND ext='.lnk' "
        "ORDER BY proposal").fetchall()

    total = db.execute("SELECT COUNT(*) FROM proposals WHERE pp IS NOT NULL").fetchone()[0]
    recs = []
    for p, n in rows:
        m = LNK_RE.match(n)
        pp = re.match(r"PP(\d+)", p)
        recs.append((int(m.group(1)) if m else None, int(pp.group(1)) if pp else None,
                     p, m.group(2) if m else n, bool(m and m.group(3))))
    recs.sort(key=lambda r: r[0] or 10 ** 6)

    live = [r for r in recs if not r[4]]
    dead = [r for r in recs if r[4]]
    print("  %d proposals of %d carry a P-link  ->  %.1f%% reach signed + onboarded"
          % (len(recs), total, 100.0 * len(recs) / total))
    print("  of those, %d cancelled after signing (%.0f%%), %d live\n"
          % (len(dead), 100.0 * len(dead) / len(recs), len(live)))

    nums = sorted(r[0] for r in recs if r[0])
    gaps = [n for n in range(min(nums), max(nums) + 1) if n not in nums]
    print("  P%d..P%d, gaps: %s" % (min(nums), max(nums), gaps or "none"))

    print("\n  %-5s %-7s %-36s %s" % ("P#", "PP#", "proposal folder", "status"))
    for pn, pp, p, addr, cancelled in recs[:args.limit]:
        print("  P%-4d PP%-5d %-36s %s" % (pn, pp, p[:36], "CANCELLED" if cancelled else ""))
    if len(recs) > args.limit:
        print("  ... %d more (--limit 0 for all)" % (len(recs) - args.limit))

    # The shortcut name is the owner's own spelling of the address, minted later than
    # the folder -- so where the two disagree, the shortcut is the newer intent.
    mismatch = [(p, addr) for _, _, p, addr, _ in recs
                if re.sub(r"^PP\d+ ", "", p).lower() != addr.lower()]
    if mismatch:
        print("\n  address differs between folder and shortcut (%d):" % len(mismatch))
        for p, addr in mismatch:
            print("     folder: %-34s shortcut: %s" % (re.sub(r"^PP\d+ ", "", p)[:34], addr))
        print("     Where this is only a street suffix, the shortcut already uses the")
        print("     USPS form -- independent support for the A5 backfill.")


# --------------------------------------------------------------------------- A5

def cmd_names(db, args):
    head("A5 -- folder-name hygiene (USPS backfill)")
    plan, review = [], []
    for (p,) in db.execute("SELECT proposal FROM proposals WHERE pp IS NOT NULL ORDER BY pp"):
        squeezed = re.sub(r"\s+", " ", p).strip()
        tokens = squeezed.split()
        why = ["whitespace"] if squeezed != p else []

        # ONLY the final token is a street suffix. "165 Terrace St" is Terrace *Street* --
        # blind word replacement would rewrite the street NAME and corrupt the address.
        if tokens and tokens[-1] in USPS:
            tokens[-1] = USPS[tokens[-1]]
            why.append("USPS suffix")
        elif any(t in USPS for t in tokens[1:]):
            # a suffix word appears mid-name (street name, or a trailing "Unit A").
            # Never guess -- surface it for a human instead.
            review.append(p)
        new = " ".join(tokens)
        if why:
            plan.append((p, new, "+".join(why)))

    if review:
        print("  NEEDS A HUMAN (%d) -- a suffix word appears mid-name, so it is probably"
              % len(review))
        print("  part of the street name, not the suffix. Not auto-renamed:")
        for p in review:
            print("     %s" % p)
        print()

    print("  %d proposals would be renamed:\n" % len(plan))
    print("  %-44s -> %-40s %s" % ("current", "proposed", "reason"))
    for old, new, why in plan:
        print("  %-44s -> %-40s %s" % (old[:44], new[:40], why))

    print("\n  Renaming is the riskiest item on the list. Drive URLs are ID-based so links")
    print("  survive, but the reconciler observes NAMES and candidateNumberTaken() matches")
    print("  on them. Use renameFolder() in WebApp/lib/proposals/drive.ts -- documented as")
    print("  the one sanctioned rename path -- not Explorer.")
    if args.csv:
        with open(args.csv, "w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["current_name", "proposed_name", "reason"])
            w.writerows(plan)
        print("\n  manifest written: %s" % args.csv)


# --------------------------------------------------------------------------- A6/A7

def cmd_registry(db, args):
    head("A6/A7 -- number gaps and duplicate addresses")
    pps = sorted(r[0] for r in db.execute("SELECT pp FROM proposals WHERE pp IS NOT NULL"))
    missing = [n for n in range(pps[0], pps[-1] + 1) if n not in pps]
    print("  PP%d..PP%d, %d folders" % (pps[0], pps[-1], len(pps)))
    print("  MISSING: %s" % (", ".join("PP%d" % n for n in missing) or "none"))
    print("  -> the filesystem cannot say whether PP104 was never minted, was trashed,")
    print("     or is a burned number. The engine's claim registry can. One lookup.\n")

    norm = lambda s: re.sub(r"\s+", " ", re.sub(r"^PP\d+\s*", "", s)).strip().lower()
    by = collections.defaultdict(list)
    for (p,) in db.execute("SELECT proposal FROM proposals WHERE pp IS NOT NULL"):
        by[norm(p)].append(p)
    for k, v in sorted(by.items()):
        if len(v) < 2:
            continue
        print("  SAME ADDRESS: %s" % "  ||  ".join(sorted(v)))
        names = {}
        for p in v:
            names[p] = {n for (n,) in db.execute(
                "SELECT name FROM nodes WHERE proposal=? AND is_dir=0", (p,))}
            d, f, b = db.execute(
                "SELECT n_dirs,n_files,bytes FROM proposals WHERE proposal=?", (p,)).fetchone()
            print("     %-26s %3d dirs %4d files %7.1f MB" % (p[:26], d, f, (b or 0) / 1e6))
        a, c = sorted(v)[0], sorted(v)[1]
        overlap = names[a] & names[c]
        print("     shared filenames: %d  (%s)" % (
            len(overlap), "likely the same job" if len(overlap) > 5 else "likely distinct jobs"))
        print("     -> CRM question, not a filesystem question.\n")


# --------------------------------------------------------------------------- A8

def cmd_shells(db, args):
    head("A8 -- proposals holding zero files")
    rows = db.execute(
        "SELECT proposal, n_dirs FROM proposals WHERE n_files=0 ORDER BY pp").fetchall()
    print("  %d proposals, %d folders between them, not one file.\n"
          % (len(rows), sum(r[1] for r in rows)))
    for p, d in rows:
        print("   %-46s %3d dirs" % (p[:46], d))
    print("\n  These are minted-then-abandoned deals. Repairing their STRUCTURE would be")
    print("  work spent on folders nobody will ever open. The useful move is lifecycle:")
    print("  mark them dead so they stop counting toward 'our 358 projects'.")


COMMANDS = {"dupes": cmd_dupes, "junk": cmd_junk, "projects": cmd_projects,
            "names": cmd_names, "registry": cmd_registry, "shells": cmd_shells}


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("command", choices=sorted(COMMANDS) + ["all"])
    ap.add_argument("--db", default=DEFAULT_DB)
    ap.add_argument("--root", default=DEFAULT_ROOT)
    ap.add_argument("--csv", default="", help="names: write the rename manifest here")
    ap.add_argument("--limit", type=int, default=12, help="projects: rows to show (0 = all)")
    args = ap.parse_args()
    if args.limit == 0:
        args.limit = 10 ** 6
    db = connect(args.db)
    for name in (sorted(COMMANDS) if args.command == "all" else [args.command]):
        COMMANDS[name](db, args)
    return 0


if __name__ == "__main__":
    sys.exit(main())
