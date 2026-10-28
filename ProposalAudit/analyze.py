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

  python analyze.py census --db data/census_pp.db [--json out.json]
                                            per template folder: real use, recency, content
  python analyze.py crossdrive --db data/census_pp.db --p-db data/census_p.db
                                            where linked PP/P pairs actually file things
  python analyze.py renames --db after.db --before before.db --rename "OLD=>NEW"
                                            did a rename pass land in Drive, ID by ID

"Occupied" means n_deep > 0: the folder, or something beneath it, holds at least
one real file. That is the only defensible definition of "used" available from the
filesystem -- mtime on a Drive mount is sync state, not human activity (see README).

census and crossdrive read the drive_nodes table (scan_proposals.py --drive), where
"real" is stricter than "occupied": a file counts only if it is not a shortcut, not OS
litter, and not a template copy. A template copy is recognised by content (its md5
matches a template file), by the engine's gsadus_template_item tag when nobody edited
it since, or by the same bytes sitting in STOCK_MIN or more folders (an older template
version, or a stock handout). Recency comes from the folder NUMBER and from Drive
dates outside migration bursts, never from mount mtime.
"""
from __future__ import annotations

import argparse
import collections
import json
import os
import re
import sqlite3
import sys
from datetime import date, timedelta

DEFAULT_DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "audit.db")
DIV_ROOT = "1 - Takeoff & Estimate/3rd Party Coordination"

SHORTCUT_MIME = "application/vnd.google-apps.shortcut"
LITTER = {"desktop.ini", "thumbs.db", ".ds_store"}
STOCK_MIN = 3        # the same bytes in this many folders is a template copy or a handout
BURST_MIN = 10       # this many folders created on one day is a bulk migration, not sales
UNDERUSED = 0.10     # under this share of the recent cohort = underused
FOLDER_NUM_RE = re.compile(r"^P{1,2}(\d+)\b")

KINDS = [  # (label, extensions); Google-native files are matched by mime instead
    ("photos", {".jpg", ".jpeg", ".heic", ".png", ".gif", ".webp", ".tif", ".tiff", ".bmp"}),
    ("videos", {".mov", ".mp4", ".m4v", ".avi", ".3gp", ".hevc"}),
    ("pdf", {".pdf"}),
    ("word", {".doc", ".docx", ".rtf", ".pages", ".txt"}),
    ("sheets", {".xls", ".xlsx", ".xlsm", ".csv", ".numbers"}),
    ("cad", {".dwg", ".dxf", ".rvt", ".skp", ".pln"}),
    ("email", {".eml", ".msg"}),
    ("archive", {".zip", ".rar", ".7z"}),
]
TOPICS = [  # (label, regex over the lower-cased file name) -- generic words only
    ("contract", r"contract|agreement|addend"),
    ("change order", r"change.?order|\bco ?#?\d"),
    ("signed", r"signed|executed|docusign"),
    ("proposal", r"proposal|quote|bid\b|bids\b"),
    ("estimate", r"estimate|takeoff|budget|pricing|price"),
    ("walkthrough video", r"\.(mov|mp4)$"),
    ("site visit", r"site.?visit|walk.?through|walkthrough"),
    ("liability release", r"liability|release"),
    ("questionnaire/selection", r"questionnaire|selection|bundle|finish|color"),
    ("plans", r"plan|elevation|floor|site.?map|layout|a\d\.\d|sheet"),
    ("permit/city", r"permit|plan.?check|correction|approval|approved|inspection|\bcity of"),
    ("survey/title/parcel", r"survey|title|parcel|apn|deed|easement|soils|geotech|perc"),
    ("utility", r"pg&?e|utility|utilities|water bill|meter|panel|solar"),
    ("invoice/receipt", r"invoice|receipt|payment|deposit|paid|check"),
    ("lien", r"lien|waiver"),
    ("closeout", r"warranty|notice.?of.?completion|occupancy|final"),
    ("notes", r"note|request|email|message|screenshot|text"),
    ("camera roll", r"^(img|dsc|pxl|vid|mvimg|photo|image)[_ -]?\d"),
]
VOCAB = {w for t in TOPICS for w in re.findall(r"[a-z]+", t[1]) if len(w) > 2} | {
    "adu", "jadu", "pp", "form", "template", "master", "standard", "central", "mini", "split",
    "client", "owner", "existing", "new", "house", "home", "garage", "conversion", "detached",
    "attached", "copy", "draft", "rev", "scan", "scanned", "document", "documents", "doc",
    "photo", "photos", "image", "images", "video", "videos", "img", "dsc", "pxl", "vid",
    "report", "info", "information", "selection", "design", "interior", "exterior", "kitchen",
    "bath", "bathroom", "bedroom", "roof", "foundation", "electrical", "plumbing", "hvac",
    "window", "windows", "door", "doors", "cabinet", "cabinets", "appliance", "appliances",
    "materials", "options", "upgrade", "upgrades", "spec", "specs", "catalog", "brochure",
    "list", "and", "the", "for", "with", "of", "to", "v", "set", "full", "page", "pages",
    "architectural", "structural", "calcs", "energy", "title24", "submittal", "fees", "fee",
    "comments", "response", "issues", "construction", "completion", "certificate", "notice",
    "order", "change", "site", "visit", "signed", "final", "plans", "permit", "invoice",
}


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


# --- census / crossdrive: over drive_nodes (scan_proposals.py --drive) ----------------

def folder_num(name):
    m = FOLDER_NUM_RE.match(name)
    return int(m.group(1)) if m else None


def kind_of(name, mime):
    if mime.startswith("application/vnd.google-apps."):
        return "google " + mime.rsplit(".", 1)[1]
    ext = os.path.splitext(name)[1].lower()
    return next((label for label, exts in KINDS if ext in exts), ext or "no ext")


def topics_of(name):
    low = name.lower()
    return [label for label, rx in TOPICS if re.search(rx, low)]


NOT_VOCAB = {"grant", "price", "water", "check", "panel", "city", "county", "color", "text"}


def shape_of(name):
    """The file name with every word outside a generic vocabulary blanked to '…', so
    client names and street addresses never leave this function."""
    stem, ext = os.path.splitext(name)
    out = []
    for tok in re.split(r"([^A-Za-z0-9]+)", stem):
        if not tok:
            continue
        if not tok[0].isalnum():
            tok = re.sub(r"[^ -~]", " ", tok)  # non-ASCII separators become spaces
            out.append(tok if len(tok) <= 3 else " ")
        elif FOLDER_NUM_RE.match(tok.upper()):
            out.append(re.sub(r"\d+", "#", tok.upper()))
        elif tok.isdigit():
            out.append("#")
        elif tok.lower() in VOCAB and tok.lower() not in NOT_VOCAB:
            out.append(tok)
        else:
            m = re.match(r"(?i)^(img|dsc|pxl|vid|mvimg)(\d+)$", tok)
            out.append(m.group(1) + "#" if m else "…")
    s = re.sub(r"…(?:[\s_.,&'-]*…)+", "…", "".join(out)).strip()
    return (s or "…") + ext.lower()


def human(modifier):
    return "@" in (modifier or "") and "gserviceaccount" not in modifier


def find_template(db, name):
    if name:
        return name
    hits = [p for (p,) in db.execute(
        "SELECT proposal FROM drive_nodes WHERE relpath='' AND proposal LIKE '%Template%'")]
    if len(hits) != 1:
        raise SystemExit("ERROR: pass --template (found %d candidates: %s)" % (len(hits), hits))
    return hits[0]


def load_census(db, template, recent_days, stale_days):
    """Classify every file outside the template and attribute it to its template folder.

    Returns (folders, files, info): folders = template dir relpaths; files = dicts;
    info = cohort facts. Each file's `home` is the deepest TEMPLATE folder above it, so
    hand-made subfolders (and renamed legacy folders) roll up to where the template put
    them; `off` marks files that sit in such a non-template folder.
    """
    if not db.execute("SELECT 1 FROM drive_nodes LIMIT 1").fetchone():
        raise SystemExit("ERROR: no drive_nodes -- run scan_proposals.py --drive first")
    tpl_dirs = {r for (r,) in db.execute(
        "SELECT relpath FROM drive_nodes WHERE proposal=? AND is_dir=1", (template,))}
    tpl_files = db.execute(
        "SELECT name, md5 FROM drive_nodes WHERE proposal=? AND is_dir=0", (template,)).fetchall()
    tpl_md5 = {m for _, m in tpl_files if m}
    # Names the template carries now, plus every name an engine copy ever carried:
    # the second set remembers files the template has since dropped or renamed.
    tpl_names = {n.lower() for n, _ in tpl_files} | {n.lower() for (n,) in db.execute(
        "SELECT DISTINCT name FROM drive_nodes WHERE is_dir=0 AND template_item IS NOT NULL")}

    roots = {p: (folder_num(p), c or "") for p, c in db.execute(
        "SELECT proposal, created FROM drive_nodes WHERE relpath='' AND proposal<>?", (template,))}
    present = collections.defaultdict(set)
    for p, r in db.execute(
            "SELECT proposal, relpath FROM drive_nodes WHERE is_dir=1 AND proposal<>?", (template,)):
        present[r].add(p)

    rows = db.execute(
        "SELECT proposal, relpath, name, mime, size, md5, created, modified, modifier, "
        "template_item FROM drive_nodes WHERE is_dir=0 AND proposal<>?", (template,)).fetchall()
    holders = collections.defaultdict(set)
    day_folders = collections.defaultdict(set)
    day_stale = collections.defaultdict(lambda: [0, 0])
    for p, rel, name, mime, _, md5, created, modified, _, _ in rows:
        if md5:
            holders[md5].add(p)
        if mime == SHORTCUT_MIME or name.lower().endswith(".lnk"):
            continue
        c, m = (created or "")[:10], (modified or "")[:10]
        day_folders[c].add(p)
        day_stale[c][0] += 1
        if c and m and (date.fromisoformat(c) - date.fromisoformat(m)).days > 30:
            day_stale[c][1] += 1
    stock = {m for m, ps in holders.items() if len(ps) >= STOCK_MIN}
    # A migration re-stamps createdTime and keeps modifiedTime. Its days show as many
    # folders created at once, or many folders receiving files mostly last modified a
    # month or more before they arrived. On those days modifiedTime is the date to use.
    root_days = collections.Counter(c[:10] for _, c in roots.values())
    burst_days = {d for d, n in root_days.items() if n >= BURST_MIN}
    bulk_days = burst_days | {d for d, ps in day_folders.items() if len(ps) >= BURST_MIN
                              and day_stale[d][1] > day_stale[d][0] / 2}
    cutoff = (date.today() - timedelta(days=recent_days)).isoformat()
    recent = {p for p, (_, c) in roots.items() if c[:10] >= cutoff and c[:10] not in burst_days}

    files = []
    for p, rel, name, mime, size, md5, created, modified, modifier, item in rows:
        low = name.lower()
        if mime == SHORTCUT_MIME or low.endswith(".lnk"):
            cls = "link"
        elif low in LITTER or low.endswith(".tmp") or low.startswith("~$"):
            cls = "litter"
        elif md5 and md5 in tpl_md5:
            cls = "template"
        elif item and ((not md5 and (modified or "")[:16] == (created or "")[:16])
                       or not human(modifier)):
            cls = "template"
        elif md5 in stock:
            cls = "template" if low in tpl_names else "stock"
        elif item or low in tpl_names:
            cls = "filled"  # a template copy someone edited: real use
        else:
            cls = "real"
        d = rel.rsplit("/", 1)[0] if "/" in rel else ""
        home = d
        while home not in tpl_dirs:
            home = home.rsplit("/", 1)[0] if "/" in home else ""
        c, m = (created or "")[:10], (modified or "")[:10]
        bulk = c in bulk_days
        files.append({
            "proposal": p, "num": roots.get(p, (None,))[0], "relpath": rel, "name": name,
            "mime": mime, "size": size or 0, "md5": md5, "cls": cls, "home": home,
            "off": d != home, "subdir": d, "recent": p in recent,
            "filed": m if bulk else c,            # when it was put here
            "day": m if bulk else max(c, m),      # last activity on it
        })
    stale = (date.today() - timedelta(days=stale_days)).isoformat()
    info = {"template": template, "n_folders": len(roots), "n_recent": len(recent),
            "cutoff": cutoff, "stale": stale, "burst_days": sorted(burst_days),
            "bulk_days": sorted(bulk_days), "roots": roots, "present": present,
            "recent": recent}
    return sorted(tpl_dirs - {""}), files, info


def status_of(n_real, n_recent_used, n_recent, newest, stale):
    """never used: no real file anywhere. abandoned: nothing real filed or edited there
    since `stale`. underused: alive, but in under UNDERUSED of the recent cohort."""
    if n_real == 0:
        return "never used"
    if newest < stale:
        return "abandoned"
    if n_recent and n_recent_used / n_recent < UNDERUSED:
        return "underused"
    return "used"


def census_rows(folders, files, info):
    real = [f for f in files if f["cls"] in ("real", "filled")]
    out = []
    for path in [""] + folders:
        mine = [f for f in real if f["home"] == path]
        deep = [f for f in real if path == "" or f["home"] == path
                or f["home"].startswith(path + "/")]
        tpl = [f for f in files if f["home"] == path and f["cls"] == "template"]
        stk = [f for f in files if f["home"] == path and f["cls"] == "stock"]
        dp = {f["proposal"] for f in deep}
        drp = {f["proposal"] for f in deep if f["recent"]}
        newest = max((f["day"] for f in deep), default="")
        nums = [f["num"] for f in deep if f["num"] is not None]
        kinds = collections.Counter(kind_of(f["name"], f["mime"]) for f in mine)
        topics = collections.Counter(t for f in mine for t in topics_of(f["name"]))
        latest = sorted(mine, key=lambda f: (f["num"] or 0, f["day"]), reverse=True)
        samples, seen = [], set()
        for f in latest:  # newest folders first, one sample per folder, varied shapes
            shp = shape_of(f["name"])
            if f["proposal"] in seen or shp in {s["file"] for s in samples}:
                continue
            seen.add(f["proposal"])
            samples.append({"num": f["num"], "file": shp, "day": f["day"],
                            "kind": kind_of(f["name"], f["mime"]),
                            "where": "subfolder" if f["off"] else "", "cls": f["cls"]})
            if len(samples) == 4:
                break
        out.append({
            "path": path,
            "container": any(x.startswith(path + "/") for x in folders) if path else True,
            "present": len(info["present"].get(path, ())) if path else info["n_folders"],
            "props_direct": len({f["proposal"] for f in mine}), "files_direct": len(mine),
            "props_deep": len(dp), "files_deep": len(deep),
            "recent_props": len(drp), "recent_share": round(len(drp) / info["n_recent"], 3)
            if info["n_recent"] else 0,
            "newest": newest, "max_num": max(nums) if nums else None,
            "off_template_files": sum(1 for f in mine if f["off"]),
            "filled_copies": sum(1 for f in mine if f["cls"] == "filled"),
            "template_copies": len(tpl), "template_copy_props": len({f["proposal"] for f in tpl}),
            "stock_files": len(stk),
            "kinds": kinds.most_common(4), "topics": topics.most_common(4), "samples": samples,
            "status": status_of(len(deep), len(drp), info["n_recent"], newest, info["stale"]),
        })
    return out


def cmd_census(db, args):
    template = find_template(db, args.template)
    folders, files, info = load_census(db, template, args.recent_days, args.stale_days)
    rows = census_rows(folders, files, info)
    cls = collections.Counter(f["cls"] for f in files)
    print("template: %s" % template)
    print("%d folders; recent cohort = %d created since %s (bulk-migration days excluded: %s)"
          % (info["n_folders"], info["n_recent"], info["cutoff"], ", ".join(info["burst_days"]) or "-"))
    print("abandoned = no real activity since %s; dated by modifiedTime on bulk days: %s"
          % (info["stale"], ", ".join(info["bulk_days"]) or "-"))
    print("files: " + ", ".join("%s %d" % kv for kv in cls.most_common()) + "\n")
    print("%-11s %5s %5s %6s %5s %6s %10s %5s %5s  %s" % (
        "status", "here", "deep", "files", "rcnt", "share", "newest", "max#", "tpl", "path"))
    for r in rows:
        print("%-11s %5d %5d %6d %5d %5.0f%% %10s %5s %5d  %s%s" % (
            r["status"], r["props_direct"], r["props_deep"], r["files_deep"], r["recent_props"],
            100 * r["recent_share"], r["newest"] or "-", r["max_num"] or "-",
            r["template_copies"], r["path"] or "(folder root)", "/" if r["container"] else ""))
        if args.verbose:
            print("%38s kinds %s | topics %s" % ("", r["kinds"], r["topics"]))
            for s in r["samples"]:
                print("%38s #%-4s %s  %s %s" % ("", s["num"], s["day"], s["file"], s["where"]))
    print("\nhere = folders with a real file sitting in this folder (or a hand-made subfolder"
          " of it); deep = incl. template subfolders;\nrcnt/share = deep use in the recent"
          " cohort; tpl = untouched template copies sitting here.")
    if args.json:
        with open(args.json, "w", encoding="utf-8") as fh:
            json.dump({"template_kind": "PP" if template.startswith("A - PP") else "P",
                       "n_folders": info["n_folders"], "n_recent": info["n_recent"],
                       "cutoff": info["cutoff"], "classes": dict(cls), "rows": rows},
                      fh, indent=1)
        print("json: " + args.json)


def link_map(files_db):
    """Proposal/project folder -> the folder ID its '0 - P…' / '0 - PP…' shortcut targets."""
    return {p: t for p, t in files_db.execute(
        "SELECT proposal, shortcut_to FROM drive_nodes WHERE relpath NOT LIKE '%/%' "
        "AND is_dir=0 AND shortcut_to IS NOT NULL AND name LIKE '0 - P%'")}


def cmd_crossdrive(db, args):
    """For every PP<->P pair, where real files landed relative to the P folder's birth.

    A PP folder that keeps receiving files after its P folder exists shows which themes
    the project side does not yet hold in practice."""
    if not args.p_db:
        raise SystemExit("ERROR: crossdrive needs --p-db (the projects drive_nodes db)")
    pdb = connect(args.p_db)
    pp_tpl, p_tpl = find_template(db, args.template), find_template(pdb, "")
    _, pp_files, pp_info = load_census(db, pp_tpl, args.recent_days, args.stale_days)
    _, p_files, p_info = load_census(pdb, p_tpl, args.recent_days, args.stale_days)
    p_ids = {i: p for p, i in pdb.execute("SELECT proposal, id FROM drive_nodes WHERE relpath=''")}
    pp_ids = {i: p for p, i in db.execute("SELECT proposal, id FROM drive_nodes WHERE relpath=''")}
    pairs = {}
    for pp, target in link_map(db).items():
        if target in p_ids:
            pairs[pp] = p_ids[target]
    back = {pp_ids[t]: p for p, t in link_map(pdb).items() if t in pp_ids}
    for pp, p in back.items():
        pairs.setdefault(pp, p)
    print("PP->P pairs: %d (from PP-side links %d, P-side links %d)"
          % (len(pairs), len(link_map(db)), len(link_map(pdb))))

    p_born = {p: c[:10] for p, (_, c) in p_info["roots"].items()}
    after = collections.Counter()
    after_props = collections.defaultdict(set)
    before = collections.Counter()
    samples = collections.defaultdict(list)
    for f in pp_files:
        p = pairs.get(f["proposal"])
        if not p or f["cls"] not in ("real", "filled"):
            continue
        born = p_born.get(p, "")
        when = f["filed"]
        if born and when > born:
            after[f["home"]] += 1
            after_props[f["home"]].add(f["proposal"])
            if len(samples[f["home"]]) < 6:
                samples[f["home"]].append("#%s %s %s" % (f["num"], when, shape_of(f["name"])))
        else:
            before[f["home"]] += 1
    print("\nreal files filed in the PP folder AFTER its P folder existed:")
    print("%6s %6s %8s  %s" % ("files", "pairs", "before", "PP template folder"))
    for home, n in after.most_common():
        print("%6d %6d %8d  %s" % (n, len(after_props[home]), before[home], home or "(root)"))
        if args.verbose:
            for s in samples[home]:
                print("%24s%s" % ("", s))

    # Byte-copies: a P file whose md5 is a real file in ITS OWN paired PP folder was
    # copied across at or after onboarding -- one theme, two homes.
    p_of = {p: pp for pp, p in pairs.items()}
    pp_home = collections.defaultdict(dict)
    for f in pp_files:
        if f["proposal"] in pairs and f["md5"] and f["cls"] in ("real", "filled"):
            pp_home[f["proposal"]][f["md5"]] = f["home"]
    p_real = [f for f in p_files if f["cls"] in ("real", "filled")]
    flow, flow_pairs = collections.Counter(), collections.defaultdict(set)
    p_own, own_props = collections.Counter(), collections.defaultdict(set)
    for f in p_real:
        src = pp_home.get(p_of.get(f["proposal"]), {}).get(f["md5"])
        if src is not None:
            flow[(src, f["home"])] += 1
            flow_pairs[(src, f["home"])].add(f["proposal"])
        else:
            p_own[f["home"]] += 1
            own_props[f["home"]].add(f["proposal"])
    print("\nP-side real files that are byte-copies of their own PP's files: %d of %d"
          % (sum(flow.values()), len(p_real)))
    print("%6s %5s  %s  ->  %s" % ("files", "pairs", "PP folder", "P folder"))
    for (a, b), n in flow.most_common(12):
        print("%6d %5d  %s  ->  %s" % (n, len(flow_pairs[(a, b)]), a or "(root)", b or "(root)"))
    print("\nP-side real files with no PP original, by P template folder:")
    for home, n in sorted(p_own.items()):
        print("%6d %4d P  %s" % (n, len(own_props[home]), home or "(root)"))
    if args.json:
        with open(args.json, "w", encoding="utf-8") as fh:
            json.dump({"pairs": len(pairs), "p_real": len(p_real), "after": [
                {"home": h, "files": n, "pairs": len(after_props[h]), "before": before[h],
                 "samples": samples[h]} for h, n in after.most_common()],
                "copies": [{"pp": a, "p": b, "files": n, "pairs": len(flow_pairs[(a, b)])}
                           for (a, b), n in flow.most_common()],
                "p_own": dict(p_own)}, fh, indent=1)
        print("json: " + args.json)


def cmd_renames(db, args):
    """Did a rename pass land in Drive? Compares two drive_nodes scans ID by ID.

    The mount says what was asked; only the Drive API says what synced. For each
    OLD=>NEW, every folder that carried OLD before is looked up by its ID after: it
    must now carry NEW (or still OLD, for the folders the pass skipped on purpose).
    Files that sat beneath OLD are followed by ID to wherever they are now."""
    if not args.before or not args.rename:
        raise SystemExit("ERROR: renames needs --before <db> and --rename 'OLD=>NEW'")
    before = connect(args.before)
    after_by_id = {i: (p, r, n) for i, p, r, n in db.execute(
        "SELECT id, proposal, relpath, name FROM drive_nodes")}
    bad = 0
    for pair in args.rename:
        old, new = [s.strip() for s in pair.split("=>")]
        print("== %s  =>  %s" % (old, new))
        olds = before.execute("SELECT id, proposal FROM drive_nodes WHERE is_dir=1 AND relpath=?",
                              (old,)).fetchall()
        outcome = collections.defaultdict(list)
        for fid, p in olds:
            now = after_by_id.get(fid)
            key = ("gone" if now is None else "renamed" if now[2] == new
                   else "unchanged" if now[2] == old else "other name: " + now[2])
            outcome[key].append(p)
        for key in ("renamed", "unchanged"):
            ps = outcome.pop(key, [])
            print("  %-10s %4d  %s" % (key, len(ps), ", ".join(
                "PP%d" % n for n in sorted(folder_num(p) for p in ps if folder_num(p) is not None))
                if key == "unchanged" else ""))
        for key, ps in outcome.items():
            bad += len(ps)
            print("  !! %-7s %4d  %s" % (key, len(ps), ", ".join(sorted(ps))))
        olds_ids = {i for i, _ in olds}
        news = db.execute("SELECT id, proposal FROM drive_nodes WHERE is_dir=1 AND relpath=?",
                          (new,)).fetchall()
        created = sorted(folder_num(p) or 0 for i, p in news if i not in olds_ids)
        print("  %-10s %4d  %s" % ("new IDs", len(created), ", ".join("PP%d" % n for n in created)))
        per = collections.Counter(p for _, p in news)
        dup = [p for p, c in per.items() if c > 1]
        both = {p for _, p in news} & {p for (p,) in db.execute(
            "SELECT proposal FROM drive_nodes WHERE is_dir=1 AND relpath=?", (old,))}
        for title, ps in (("two folders named " + new, dup), ("both names present", both)):
            if ps:
                bad += len(ps)
                print("  !! %s: %s" % (title, ", ".join(sorted(ps))))
        stayed = 0
        for fid, p, rel in before.execute(
                "SELECT id, proposal, relpath FROM drive_nodes WHERE is_dir=0 AND relpath LIKE ?",
                (old + "/%",)):
            now = after_by_id.get(fid)  # (proposal, relpath, name)
            if now is not None and now[0] == p and os.path.dirname(now[1]) == os.path.dirname(rel):
                stayed += 1  # a skipped folder's content, where it was
                continue
            if now is None or now[0] != p:
                bad += 1
            where = ("!! GONE" if now is None else "!! other folder: " + now[0] if now[0] != p
                     else os.path.dirname(now[1]) or "(root)")
            print("  file moved PP%s  %s  ->  %s"
                  % (folder_num(p), shape_of(rel.rsplit("/", 1)[-1]), where))
        if stayed:
            print("  files      %4d  still where they were (inside folders the pass skipped)" % stayed)
    print("\n%s" % ("OK: every folder is where the pass meant it to be" if not bad
                    else "%d problem(s) above" % bad))


COMMANDS = {
    "summary": cmd_summary, "occupancy": cmd_occupancy, "never": cmd_never,
    "outliers": cmd_outliers, "dupes": cmd_dupes,
    "census": cmd_census, "crossdrive": cmd_crossdrive, "renames": cmd_renames,
}


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("command", choices=sorted(COMMANDS))
    ap.add_argument("--db", default=DEFAULT_DB)
    ap.add_argument("--under", default="", help="occupancy: restrict to this subtree")
    ap.add_argument("--template", default="",
                    help="census/crossdrive: the template folder's name (default: the one "
                         "folder whose name contains 'Template')")
    ap.add_argument("--p-db", default="", help="crossdrive: the projects drive_nodes db")
    ap.add_argument("--before", default="", help="renames: the drive_nodes db scanned before")
    ap.add_argument("--rename", action="append", default=[],
                    help="renames: 'OLD=>NEW' (repeatable)")
    ap.add_argument("--recent-days", type=int, default=365,
                    help="census: the recent cohort is folders created in this window")
    ap.add_argument("--stale-days", type=int, default=180,
                    help="census: no real activity in this window = abandoned")
    ap.add_argument("--json", default="", help="census/crossdrive: also write rows as JSON")
    ap.add_argument("-v", "--verbose", action="store_true",
                    help="census/crossdrive: content kinds, topics and sanitized samples")
    args = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    COMMANDS[args.command](connect(args.db), args)
    return 0


if __name__ == "__main__":
    sys.exit(main())
