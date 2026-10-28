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

--drive: the Drive API layer. Instead of walking the mount, list the same tree through
the `gws` CLI (signed in as the owner) into the `drive_nodes` table: Drive IDs, md5,
createdTime/modifiedTime, lastModifyingUser, the engine's `gsadus_template_item` tag,
and shortcut targets -- everything the mount cannot show. It calls exactly two gws
verbs, `drives list` and `files list`; both are reads. Folders are expanded a level at
a time, 40 parents per query, so the cost is proportional to the tree, not the drive.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sqlite3
import subprocess
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
CREATE TABLE IF NOT EXISTS drive_nodes (
    id            TEXT PRIMARY KEY, -- Drive ID; relpaths can repeat (Drive allows same-named siblings)
    proposal      TEXT NOT NULL,
    relpath       TEXT NOT NULL,    -- as in nodes: relative to the PROPOSAL root ('' = itself)
    name          TEXT,
    depth         INTEGER,
    is_dir        INTEGER,
    mime          TEXT,
    size          INTEGER,          -- NULL for Google-native files
    md5           TEXT,             -- NULL for Google-native files
    created       TEXT,             -- Drive createdTime (ISO): when the object entered Drive
    modified      TEXT,             -- Drive modifiedTime (ISO)
    modifier      TEXT,             -- lastModifyingUser email (or display name)
    template_item TEXT,             -- appProperties.gsadus_template_item: an engine copy
    shortcut_to   TEXT,             -- shortcutDetails.targetId
    scanned_at    TEXT
);
CREATE INDEX IF NOT EXISTS idx_drive_proposal ON drive_nodes(proposal, relpath);
CREATE INDEX IF NOT EXISTS idx_drive_md5      ON drive_nodes(md5);
"""

PP_RE = re.compile(r"^PP(\d+)\b")

FOLDER_MIME = "application/vnd.google-apps.folder"
DRIVE_FIELDS = ("nextPageToken,files(id,name,mimeType,parents,size,md5Checksum,createdTime,"
                "modifiedTime,lastModifyingUser(emailAddress,displayName),appProperties,"
                "shortcutDetails(targetId))")
PARENTS_PER_QUERY = 40  # keeps the q string well under the GET URL limit


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


# --- Drive API layer (--drive) ------------------------------------------------------

def gws_exe():
    """The real gws binary. The npm shim is a .cmd, and cmd.exe mangles the JSON in
    --params; the binary beside it takes argv as given. GWS_EXE overrides."""
    if os.environ.get("GWS_EXE"):
        return os.environ["GWS_EXE"]
    shim = shutil.which("gws")
    if shim:
        exe = os.path.join(os.path.dirname(shim), "node_modules", "@googleworkspace",
                           "cli", "bin", "gws.exe")
        if os.path.isfile(exe):
            return exe
        return shim
    raise SystemExit("ERROR: gws not found (npm i -g @googleworkspace/cli)")


def gws_read(resource, params):
    """One read-only gws list call, every page. Returns the items of all pages."""
    cmd = [gws_exe(), "drive", resource, "list", "--params", json.dumps(params),
           "--page-all", "--page-limit", "1000", "--page-delay", "0"]
    for attempt in range(3):
        p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
        if p.returncode == 0:
            break
        time.sleep(2 * (attempt + 1))
    else:
        raise SystemExit("ERROR: gws failed (re-login? run: gws auth login)\n"
                         + (p.stderr or p.stdout)[-800:])
    key = "drives" if resource == "drives" else "files"
    items = []
    for line in p.stdout.splitlines():
        line = line.strip()
        if line.startswith("{"):
            items.extend(json.loads(line).get(key, []))
    return items


def qlit(value):
    return value.replace("\\", "\\\\").replace("'", "\\'")


def drive_children(drive_id, parent_ids, folders_only=False):
    q = "(" + " or ".join("'%s' in parents" % qlit(p) for p in parent_ids) + ") and trashed=false"
    if folders_only:
        q += " and mimeType='%s'" % FOLDER_MIME
    return gws_read("files", {
        "corpora": "drive", "driveId": drive_id, "includeItemsFromAllDrives": True,
        "supportsAllDrives": True, "pageSize": 1000, "q": q, "fields": DRIVE_FIELDS})


def resolve_root(root):
    """'G:\\Shared drives\\<drive>\\a\\b' -> (drive_id, folder_id), by name, read-only."""
    parts = [s for s in re.split(r"[\\/]+", root) if s]
    try:
        i = [s.lower() for s in parts].index("shared drives")
    except ValueError:
        raise SystemExit("ERROR: --drive needs a root under 'Shared drives': " + root)
    drive_name, segs = parts[i + 1], parts[i + 2:]
    drives = gws_read("drives", {"pageSize": 100, "fields": "nextPageToken,drives(id,name)"})
    match = [d["id"] for d in drives if d.get("name") == drive_name]
    if len(match) != 1:
        raise SystemExit("ERROR: shared drive %r not found (or ambiguous)" % drive_name)
    drive_id = folder_id = match[0]
    for seg in segs:
        hits = [f for f in drive_children(drive_id, [folder_id], folders_only=True)
                if f.get("name") == seg]
        if len(hits) != 1:
            raise SystemExit("ERROR: %d folders named %r under %s" % (len(hits), seg, folder_id))
        folder_id = hits[0]["id"]
    return drive_id, folder_id


def drive_row(f, proposal, relpath, depth, stamp):
    user = f.get("lastModifyingUser") or {}
    return (f["id"], proposal, relpath, f.get("name", ""), depth,
            1 if f.get("mimeType") == FOLDER_MIME else 0, f.get("mimeType", ""),
            int(f["size"]) if f.get("size") is not None else None, f.get("md5Checksum"),
            f.get("createdTime", ""), f.get("modifiedTime", ""),
            user.get("emailAddress") or user.get("displayName") or "",
            (f.get("appProperties") or {}).get("gsadus_template_item"),
            (f.get("shortcutDetails") or {}).get("targetId"), stamp)


def drive_scan(args, db):
    drive_id, root_id = resolve_root(args.root)
    stamp = datetime.now(timezone.utc).isoformat(timespec="seconds")
    top = [f for f in drive_children(drive_id, [root_id]) if f.get("mimeType") == FOLDER_MIME]
    top.sort(key=lambda f: f.get("name", ""))
    if args.only:
        rx = re.compile(args.only)
        top = [f for f in top if rx.search(f.get("name", ""))]
    if args.limit:
        top = top[:args.limit]
    if not top:
        print("nothing to scan")
        return 0
    print("drive scan: %d folders under %s x %d workers" % (len(top), args.root, args.workers),
          flush=True)

    t0 = time.perf_counter()
    rows = [drive_row(f, f["name"], "", 0, stamp) for f in top]
    where = {f["id"]: (f["name"], "") for f in top}  # folder id -> (proposal, relpath)
    frontier, depth = [f["id"] for f in top], 0
    while frontier:
        depth += 1
        batches = [frontier[i:i + PARENTS_PER_QUERY]
                   for i in range(0, len(frontier), PARENTS_PER_QUERY)]
        frontier = []
        with ThreadPoolExecutor(max_workers=args.workers) as pool:
            for fut in as_completed([pool.submit(drive_children, drive_id, b) for b in batches]):
                for f in fut.result():
                    proposal, prel = where[f["parents"][0]]
                    rel = prel + "/" + f["name"] if prel else f["name"]
                    rows.append(drive_row(f, proposal, rel, depth, stamp))
                    if f.get("mimeType") == FOLDER_MIME:
                        where[f["id"]] = (proposal, rel)
                        frontier.append(f["id"])
        print("  depth %d: %d rows so far" % (depth, len(rows)), flush=True)

    # Same ghost-row rule as the mount walk: a re-scan replaces each folder's rows.
    names = [f["name"] for f in top]
    db.executemany("DELETE FROM drive_nodes WHERE proposal=?", [(n,) for n in names])
    db.executemany("INSERT OR REPLACE INTO drive_nodes VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                   rows)
    db.commit()
    n_files = sum(1 for r in rows if not r[5])
    print("\ndone in %.1fs -- %d folders, %d dirs, %d files recorded"
          % (time.perf_counter() - t0, len(top), len(rows) - n_files, n_files))
    print("db: " + args.db)
    return 0


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
    ap.add_argument("--drive", action="store_true",
                    help="list the root through the Drive API (gws, read-only) into "
                         "drive_nodes instead of walking the mount")
    args = ap.parse_args()

    if args.drive:
        os.makedirs(os.path.dirname(args.db), exist_ok=True)
        db = sqlite3.connect(args.db)
        db.executescript(SCHEMA)
        return drive_scan(args, db)

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
