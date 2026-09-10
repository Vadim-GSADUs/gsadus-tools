#!/usr/bin/env python3
"""
scan_proposals.py -- parallel metadata walk of the GSADUs proposal corpus into SQLite.

WHY THREADS: the proposals live on the Google Drive for Desktop mount (G:), which is
latency-bound, not CPU-bound. A single-threaded `find` over the corpus clocked ~9 hours;
one thread per proposal folder collapses that to minutes. Nothing here is CPU work.

WHY SQLITE: the dataset is the deliverable, not a report. Every question we ask later
(occupancy, drift, bloat, outliers) is a query over these two tables, so a rescan never
means re-deriving an answer by hand.

WHY scandir AND NOT os.walk: on Windows the directory listing already carries size and
mtime, so DirEntry.stat() is free. os.walk throws that away and forces a stat syscall
per file -- which on a network-backed mount is the entire cost of the scan.

CAUTION -- mtime here is NOT a usage signal. On a Drive mount it reflects sync and cache
state, and the engine's mint writes fresh timestamps on every object it creates. Use it
for forensics only. Real "was this ever touched by a human" lives in the Drive API
(lastModifyingUser, viewedByMeTime) or the Drive Activity API. See README.md.
"""
from __future__ import annotations

import argparse
import os
import re
import sqlite3
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone

DEFAULT_ROOT = r"G:\Shared drives\GSADUs Business\1 - Proposals"
DEFAULT_DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "audit.db")

SCHEMA = """
CREATE TABLE IF NOT EXISTS proposals (
    proposal     TEXT PRIMARY KEY,  -- folder name as it sits on disk
    pp           INTEGER,           -- parsed PP number; NULL for the template / oddities
    scope        TEXT,              -- '' = whole proposal, else the subtree relpath scanned
    complete     INTEGER,           -- 1 only when scope = '' (a full walk)
    n_dirs       INTEGER,
    n_files      INTEGER,
    n_empty_dirs INTEGER,
    bytes        INTEGER,
    walk_secs    REAL,
    errors       INTEGER,
    scanned_at   TEXT
);
CREATE TABLE IF NOT EXISTS nodes (
    proposal  TEXT NOT NULL,
    relpath   TEXT NOT NULL,   -- '/'-separated, relative to the PROPOSAL root ('' = itself)
    name      TEXT,
    depth     INTEGER,
    is_dir    INTEGER,
    size      INTEGER,         -- files: bytes; dirs: deep byte total
    mtime     TEXT,            -- ISO date. NOT a usage signal -- see module docstring.
    ext       TEXT,
    n_direct  INTEGER,         -- dirs: files sitting directly inside
    n_deep    INTEGER,         -- dirs: files anywhere beneath
    n_subdirs INTEGER,
    PRIMARY KEY (proposal, relpath)
);
CREATE INDEX IF NOT EXISTS idx_nodes_relpath ON nodes(relpath);
CREATE INDEX IF NOT EXISTS idx_nodes_dir     ON nodes(is_dir, relpath);
CREATE TABLE IF NOT EXISTS scan_errors (
    proposal TEXT, path TEXT, error TEXT
);
"""

PP_RE = re.compile(r"^PP(\d+)\b")


def parse_pp(name):
    m = PP_RE.match(name)
    return int(m.group(1)) if m else None


def iso(ts):
    if not ts:
        return ""
    try:
        return datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d")
    except (OSError, OverflowError, ValueError):
        return ""


def scan_dir(base, path, depth, out, errors):
    """Post-order scandir recursion. Appends node rows to `out`.

    Returns (deep_file_count, deep_bytes) so each directory row can record what
    actually lives beneath it -- that is the number the occupancy question needs.
    """
    dirs, files = [], []
    try:
        with os.scandir(path) as it:
            for e in it:
                try:
                    if e.is_dir(follow_symlinks=False):
                        dirs.append(e)
                    else:
                        files.append(e)
                except OSError as ex:
                    errors.append((e.path, str(ex)))
    except OSError as ex:
        errors.append((path, str(ex)))
        return 0, 0

    deep_files, deep_bytes = len(files), 0
    for e in files:
        try:
            st = e.stat()
            size, mt = st.st_size, st.st_mtime
        except OSError as ex:
            errors.append((e.path, str(ex)))
            size, mt = -1, 0.0
        deep_bytes += max(size, 0)
        rel = os.path.relpath(e.path, base).replace("\\", "/")
        out.append((rel, e.name, depth + 1, 0, size, iso(mt),
                    os.path.splitext(e.name)[1].lower(), None, None, None))

    for e in dirs:
        cf, cb = scan_dir(base, e.path, depth + 1, out, errors)
        deep_files += cf
        deep_bytes += cb

    rel = os.path.relpath(path, base).replace("\\", "/")
    if rel == ".":
        rel = ""
    try:
        mt = iso(os.stat(path).st_mtime)
    except OSError:
        mt = ""
    out.append((rel, os.path.basename(path), depth, 1, deep_bytes, mt, "",
                len(files), deep_files, len(dirs)))
    return deep_files, deep_bytes


def walk_proposal(root, proposal, subtree):
    base = os.path.join(root, proposal)
    start = os.path.join(base, subtree.replace("/", os.sep)) if subtree else base
    out, errors = [], []
    t0 = time.perf_counter()
    if not os.path.isdir(start):
        return proposal, [], [("", "subtree not present")], 0.0
    depth0 = 0 if not subtree else subtree.count("/") + 1
    scan_dir(base, start, depth0, out, errors)
    return proposal, out, errors, time.perf_counter() - t0


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--root", default=DEFAULT_ROOT)
    ap.add_argument("--db", default=DEFAULT_DB)
    ap.add_argument("--workers", type=int, default=16,
                    help="parallel proposal walks (latency-bound; 16 is a sane default)")
    ap.add_argument("--subtree", default="",
                    help="limit each walk to this relpath inside the proposal, e.g. "
                         "'1 - Takeoff & Estimate/3rd Party Coordination'")
    ap.add_argument("--limit", type=int, default=0, help="scan only the first N proposals")
    ap.add_argument("--only", default="", help="regex: scan only matching proposal names")
    ap.add_argument("--resume", action="store_true",
                    help="skip proposals already recorded at this scope")
    args = ap.parse_args()

    if not os.path.isdir(args.root):
        print("ERROR: root not reachable: " + args.root, file=sys.stderr)
        return 2

    os.makedirs(os.path.dirname(args.db), exist_ok=True)
    db = sqlite3.connect(args.db)
    db.executescript(SCHEMA)

    with os.scandir(args.root) as it:
        names = sorted(e.name for e in it if e.is_dir(follow_symlinks=False))
    if args.only:
        rx = re.compile(args.only)
        names = [n for n in names if rx.search(n)]
    if args.resume:
        done = {r[0] for r in db.execute(
            "SELECT proposal FROM proposals WHERE scope = ?", (args.subtree,))}
        names = [n for n in names if n not in done]
    if args.limit:
        names = names[:args.limit]

    if not names:
        print("nothing to scan (already done? try without --resume)")
        return 0

    scope_note = ("  subtree=" + args.subtree) if args.subtree else ""
    print("scanning %d proposals x %d workers%s" % (len(names), args.workers, scope_note),
          flush=True)

    t0 = time.perf_counter()
    done_n = 0
    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        futs = [pool.submit(walk_proposal, args.root, n, args.subtree) for n in names]
        for fut in as_completed(futs):
            proposal, rows, errors, secs = fut.result()
            n_dirs = sum(1 for r in rows if r[3])
            n_files = sum(1 for r in rows if not r[3])
            n_empty = sum(1 for r in rows if r[3] and r[8] == 0)
            nbytes = next((r[4] for r in rows if r[3] and r[0] == ""), 0)

            # Clear this proposal's previous rows before inserting the fresh walk.
            # INSERT OR REPLACE alone would leave GHOST rows behind for anything that
            # has since been deleted on disk -- a re-scan after a cleanup would then
            # over-report folders that no longer exist. Scope the wipe to the subtree
            # when one was given, so a subtree scan never discards the wider corpus.
            if args.subtree:
                db.execute("DELETE FROM nodes WHERE proposal=? AND (relpath=? OR relpath LIKE ?)",
                           (proposal, args.subtree, args.subtree + "/%"))
            else:
                db.execute("DELETE FROM nodes WHERE proposal=?", (proposal,))
            db.execute("DELETE FROM scan_errors WHERE proposal=?", (proposal,))

            db.executemany(
                "INSERT OR REPLACE INTO nodes (proposal,relpath,name,depth,is_dir,size,"
                "mtime,ext,n_direct,n_deep,n_subdirs) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                [(proposal,) + tuple(r) for r in rows])
            db.executemany(
                "INSERT INTO scan_errors (proposal,path,error) VALUES (?,?,?)",
                [(proposal, p, e) for p, e in errors])
            db.execute(
                "INSERT OR REPLACE INTO proposals (proposal,pp,scope,complete,n_dirs,"
                "n_files,n_empty_dirs,bytes,walk_secs,errors,scanned_at) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                (proposal, parse_pp(proposal), args.subtree, 0 if args.subtree else 1,
                 n_dirs, n_files, n_empty, nbytes, round(secs, 2), len(errors),
                 datetime.now(timezone.utc).isoformat(timespec="seconds")))
            db.commit()

            done_n += 1
            if done_n % 25 == 0 or done_n == len(names):
                rate = done_n / (time.perf_counter() - t0)
                eta = (len(names) - done_n) / rate if rate else 0
                print("  %d/%d  %.1f/s  eta %5.0fs" % (done_n, len(names), rate, eta),
                      flush=True)

    elapsed = time.perf_counter() - t0
    tot = db.execute("SELECT COUNT(*) FROM nodes").fetchone()[0]
    err = db.execute("SELECT COUNT(*) FROM scan_errors").fetchone()[0]
    print("\ndone in %.1fs -- %d node rows in db, %d scan errors" % (elapsed, tot, err))
    print("db: " + args.db)
    db.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
