#!/usr/bin/env python3
"""
restore_branch.py -- un-trash a branch folder via the Drive API, by IDENTITY.

  python restore_branch.py --branch "1 - Takeoff & Estimate"
  python restore_branch.py --branch "1 - Takeoff & Estimate" --apply

WHY UNTRASH RATHER THAN RE-CREATE: the original folders still exist in the shared
drive's trash with their original IDs, parents, and `appProperties`. Engine-minted
proposals (PP345+) carry `gsadus_template_item` on these folders -- the tag the
hydration completion check joins on. Re-creating them would mint new IDs with no
tags, and a later repair pass would see the template item as missing and duplicate
it. Restoring keeps identity intact; re-creating quietly breaks it.

Weigh that honestly, though: on 2026-09-10 it applied to 16 of 353 folders. The
other 337 predate the engine, carry no tags, and re-creating them would have cost
nothing. A real-but-narrow concern was allowed to pick the complicated path over
the simple one. Count the tagged folders first (the dry run prints it) and decide
with that number in hand.

!! THIS RESTORES THE WHOLE SUBTREE, NOT JUST THE FOLDER. !!

Read this before using it. On 2026-09-10 this script was written on the belief that
because the cleanup deleted bottom-up with os.rmdir, Drive had received a separate
trash operation per folder, so un-trashing a parent would return an EMPTY folder.

That was wrong, and it was never tested. Drive collapses the deletion into the
highest folder removed: un-trashing `1 - Takeoff & Estimate` brought back all 73
descendants per proposal -- 26,753 folders -- and they had to be deleted a second
time. No files were affected (that subtree held none), but it cost a full extra
delete pass and two more sync cycles across every machine.

The `explicitlyTrashed` check below did NOT catch it, and that is the more useful
lesson. It queried only folders NAMED `--branch`, which were of course explicitly
trashed, and reported "0 inherited" -- looking exactly like confirmation while never
examining the descendants, which were the actual question. A check that reports
success on the wrong subject is worse than no check: it manufactures confidence.

So: use this to restore a branch AND everything under it. If you want an empty
placeholder back, just re-create the folder -- it is seconds of work with no sync
churn. Only reach for un-trashing when folder IDENTITY genuinely matters (see
below), and confirm what is actually in the trash beneath the target first.

SCOPE: only folders whose parent is a proposal folder directly under the proposals
root are touched. Anything else that happens to share the name is ignored.

Credentials come from Doppler at call time and are never written to disk, logged,
or printed.
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from datetime import datetime, timezone

PROPOSALS_ROOT_FOLDER_ID = "15SYRPedu-ToLpnHQDu8TfSolPwYqYNc6"
FOLDER_MIME = "application/vnd.google-apps.folder"
SECRET = "PROPOSAL_ENGINE_DRIVE_CREDENTIALS"


def drive_client():
    """Service-account client. The key is read from Doppler into memory only."""
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
    except ImportError:
        sys.exit("missing deps: pip install google-api-python-client google-auth")

    proc = subprocess.run(
        ["doppler", "secrets", "get", SECRET, "--plain",
         "--project", "webapp", "--config", "dev"],
        capture_output=True, text=True)
    if proc.returncode != 0:
        # stderr only -- never echo stdout, it would be the key material.
        sys.exit("doppler failed: " + proc.stderr.strip()[:300])
    try:
        info = json.loads(proc.stdout.strip())
    except json.JSONDecodeError:
        sys.exit("secret %s is not valid JSON" % SECRET)

    creds = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive"])
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def list_all(drive, **kw):
    """files.list paginated to exhaustion -- a single page silently loses rows."""
    out, token = [], None
    while True:
        res = drive.files().list(
            pageSize=1000, pageToken=token,
            supportsAllDrives=True, includeItemsFromAllDrives=True, **kw).execute()
        out.extend(res.get("files", []))
        token = res.get("nextPageToken")
        if not token:
            return out


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--branch", required=True, help="folder name to restore")
    ap.add_argument("--apply", action="store_true", help="actually un-trash")
    ap.add_argument("--prefer-tagged", action="store_true",
                    help="when a proposal has several same-named trashed copies, keep "
                         "only the one carrying gsadus_template_item")
    ap.add_argument("--manifest", default=os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "data", "restore_manifest.jsonl"))
    args = ap.parse_args()

    drive = drive_client()

    root = drive.files().get(fileId=PROPOSALS_ROOT_FOLDER_ID,
                             fields="id,name,driveId", supportsAllDrives=True).execute()
    drive_id = root.get("driveId")
    print("proposals root : %s" % root.get("name"))

    proposals = {f["id"]: f["name"] for f in list_all(
        drive, q="'%s' in parents and mimeType='%s' and trashed=false"
                 % (PROPOSALS_ROOT_FOLDER_ID, FOLDER_MIME),
        fields="nextPageToken, files(id,name)")}
    print("proposal folders: %d" % len(proposals))

    # A drive-wide query (corpora='drive' + driveId) is refused: the service account
    # holds a FOLDER-scoped Content-manager grant on the proposals root, not shared
    # drive membership -- "teamDriveMembershipRequired". That containment is a feature
    # (this identity cannot reach anything outside the proposals tree), so search
    # parent-by-parent inside the grant instead. One query per proposal, threaded.
    esc = args.branch.replace("\\", "\\\\").replace("'", "\\'")

    # Batch parents into OR-groups: ~15 queries instead of 360, and SEQUENTIAL.
    # An earlier version threaded one shared googleapiclient across 12 workers and got
    # a storm of SSL WRONG_VERSION_NUMBER / BAD_RECORD_MAC errors -- that client is not
    # thread-safe, and concurrent use corrupts the underlying socket. Sharing it looks
    # like it works right up until the responses are garbage.
    ids = list(proposals)
    CHUNK = 25
    targets, foreign = [], 0
    for i in range(0, len(ids), CHUNK):
        group = ids[i:i + CHUNK]
        clause = " or ".join("'%s' in parents" % g for g in group)
        hits = list_all(
            drive,
            q="(%s) and name='%s' and mimeType='%s' and trashed=true" % (clause, esc, FOLDER_MIME),
            fields="nextPageToken, files(id,name,parents,explicitlyTrashed,appProperties)")
        for f in hits:
            parent = (f.get("parents") or [None])[0]
            if parent in proposals:
                targets.append((f, proposals[parent]))
            else:
                foreign += 1
        print("   scanned %d/%d proposals, %d found so far"
              % (min(i + CHUNK, len(ids)), len(ids), len(targets)))
    trashed = targets

    # A proposal with TWO trashed folders of the same name is a forked structure in the
    # trash. Google Drive permits same-name siblings -- the "(1)" seen on the Windows
    # mount is only Drive for Desktop disambiguating them for the filesystem. Restoring
    # both would rebuild the fork. Resolve by identity: keep the copy carrying
    # `gsadus_template_item`, which is the one the engine's completion check joins on;
    # an untagged twin is an orphan the engine cannot see.
    by_proposal = {}
    for f, prop in targets:
        by_proposal.setdefault(prop, []).append(f)
    dupes = {p: v for p, v in by_proposal.items() if len(v) > 1}
    if dupes:
        print("\n!! %d proposal(s) have MULTIPLE trashed '%s' folders:" % (len(dupes), args.branch))
        for p, v in dupes.items():
            for f in v:
                print("     %-40s %s  %s" % (p[:40], f["id"],
                                             "TAGGED" if f.get("appProperties") else "untagged"))
        if not args.prefer_tagged:
            sys.exit("REFUSED: restoring all of them would recreate a forked structure.\n"
                     "         Re-run with --prefer-tagged to keep only the tagged copy.")
        keep = []
        for f, prop in targets:
            v = by_proposal[prop]
            if len(v) == 1 or f.get("appProperties"):
                keep.append((f, prop))
        dropped = len(targets) - len(keep)
        print("   --prefer-tagged: keeping the tagged copy, skipping %d untagged twin(s)\n"
              % dropped)
        targets = keep

    tagged = sum(1 for f, _ in targets if f.get("appProperties"))
    inherited = sum(1 for f, _ in targets if not f.get("explicitlyTrashed"))
    print("trashed '%s' folders found : %d" % (args.branch, len(trashed)))
    print("  inside a proposal folder : %d" % len(targets))
    print("  elsewhere (ignored)      : %d" % foreign)
    print("  carrying appProperties   : %d  (identity that re-creation would lose)" % tagged)
    # NOTE: this only describes the NAMED folders, never their descendants. It does
    # NOT tell you whether children will come back -- they will, regardless of what
    # this prints. Kept because it is real information about the targets themselves;
    # do not read it as a safety check. See the header.
    print("  NOT explicitly trashed   : %d  (about these folders only, NOT their children)"
          % inherited)
    print("\n  !! Restoring these will also restore EVERYTHING beneath them.")
    print("     If you only need an empty placeholder, create the folder instead.")

    # Always dump the full candidate list. A summary count is not reviewable, and the
    # only way to know this set matches what a cleanup removed is to diff the names.
    cand = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data",
                        "restore_candidates.txt")
    os.makedirs(os.path.dirname(cand), exist_ok=True)
    with open(cand, "w", encoding="utf-8") as fh:
        for f, prop in sorted(targets, key=lambda t: t[1]):
            fh.write("%s\t%s\t%s\n" % (prop, f["id"], "tagged" if f.get("appProperties") else ""))
    print("candidate list: %s" % cand)

    if not args.apply:
        for f, prop in sorted(targets, key=lambda t: t[1])[:6]:
            print("   would restore: %-42s %s" % (prop[:42], f["id"]))
        print("   ... full list in the file above")
        print("\nDry run. Re-run with --apply to restore.")
        return 0

    os.makedirs(os.path.dirname(args.manifest), exist_ok=True)
    ok = fail = 0
    with open(args.manifest, "a", encoding="utf-8") as log:
        for f, prop in sorted(targets, key=lambda t: t[1]):
            try:
                drive.files().update(fileId=f["id"], body={"trashed": False},
                                     supportsAllDrives=True).execute()
                ok += 1
                status = "restored"
            except Exception as ex:              # noqa: BLE001 - report, never swallow
                fail += 1
                status = "FAILED: %s" % str(ex)[:200]
            log.write(json.dumps({
                "ts": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                "proposal": prop, "folder_id": f["id"], "branch": args.branch,
                "status": status}) + "\n")
            if ok and ok % 50 == 0:
                print("   ... %d restored" % ok)

    print("\nrestored: %d   failed: %d" % (ok, fail))
    print("manifest: %s" % args.manifest)
    return 1 if fail else 0


if __name__ == "__main__":
    sys.exit(main())
