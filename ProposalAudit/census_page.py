"""census_page.py -- build the owner's census page: today's PP0 and P0 side by side.

Every row is a folder that exists in today's templates, so the page shrinks as batches
land. The rows and counts come from a fresh Drive scan; THEMES and LEFT below are the
owner's decisions in words, so update them whenever a batch lands or a decision changes.
They must agree with the README's Open items.

    python scan_proposals.py --drive --db data/census_pp.db
    python scan_proposals.py --drive --db data/census_p.db --root "G:/Shared drives/GSADUs Projects/2024 Projects"
    (the same for 2025 Projects and 2026 Projects)
    python analyze.py census --db data/census_pp.db --json data/census_pp.json
    python analyze.py census --db data/census_p.db --json data/census_p.json
    python census_page.py          # writes data/folder-census.html

Then publish data/folder-census.html over https://claude.ai/artifact/5BhsxKE1jxpoXDbgveP8d9
(the Artifact tool with that url), so the owner's link keeps working.
"""
import html, json, os
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
OUT = os.path.join(DATA, "folder-census.html")
READ_ON = date.fromtimestamp(os.path.getmtime(os.path.join(DATA, "census_pp.json"))).isoformat()
pp = json.load(open(os.path.join(DATA, "census_pp.json"), encoding="utf-8"))
p = json.load(open(os.path.join(DATA, "census_p.json"), encoding="utf-8"))
PP = {r["path"]: r for r in pp["rows"]}
P = {r["path"]: r for r in p["rows"]}

SD, SD2 = "3 - Supporting Documents", "2 - Supporting Documents"
WT2 = SD2 + "/A - Project (Images & Videos)"
# state: done (the decided change has landed), open (something is still to do), settled
# (nothing planned). Themes follow the PP tree, then the P-only ones.
THEMES = [
    dict(name="CRM history", home="PP", state="done", label="Done · Batch 1, 7",
         pp=["1 - CRM"], p=[],
         p_note="Read through the <code>0 - PP…</code> shortcut.",
         why="Each deal's notes, calls and stage history as <code>CRM history.md</code>, rewritten nightly at 04:00. The backfill finishes by 2026-10-30."),
    dict(name="Proposal & contract", home="PP", state="settled", label="Settled",
         pp=["2 - Proposal - Contract"], p=[],
         p_note="No folder. The <code>0 - PP…</code> shortcut links back.",
         why="The signed contract sits here for 63 of 63 projects. PM onboarding finds it by this folder's name."),
    dict(name="Change orders", home="PP", state="open", label="Open · Batch 2, later",
         pp=["2 - Proposal - Contract/Change Order(s)"],
         p=["4 - Administrative", "4 - Administrative/1 - Change Order Request"],
         why="Stay in the PP for now, beside the contract and the other financial documents. Revisit with Sales and owners. Every PP with a real change order is a signed project."),
    dict(name="Site walkthrough", home="PP", state="open", label="Batch 3 done · old copies open",
         pp=[SD + "/A - Site Walkthrough", SD + "/A - Site Walkthrough/Images & Videos"],
         p=[WT2, WT2 + "/1 - Site Walk-Through (Images & Videos)"],
         p_note="All 63 projects now have a <code>PP### Site Walkthrough</code> shortcut in <code>A - Project</code>, made by PM's bridge.",
         why="One home, the PP. Still open: the 6,435 copies already in P folders (43.7 GB) stay until you decide."),
    dict(name="Construction photos", home="P", state="settled", label="Settled",
         pp=[], p=[WT2 + "/2 - Construction (Images & Videos)"],
         pp_note="None, correctly.",
         why="Projects that haven't started building read as low use."),
    dict(name="Questionnaire & selection", home="P", state="open", label="Open · Batch 4, held",
         pp=[SD + "/B - Questionnaire & Selection"], p=[SD2 + "/B - Questionnaire & Selection"],
         why="Retiring from the PP, after the Sales and owners meeting and once the blank forms leave PP0. Design selections live in the P."),
    dict(name="Received from client", home="both", state="done", label="Done · both sides",
         pp=[SD + "/C - Received from Client"], p=[SD2 + "/C - Received from Client"],
         pp_note="Flat since 2026-10-28 (Batch 5).",
         p_note="Flat since 2026-10-28, P0 included (your ruling in PM's notes).",
         why="Thin on both sides. Anything the client sends goes straight in."),
    dict(name="Plan set", home="both", state="done", label="Done · Batch 6",
         pp=[SD + "/D - Plan Set"], p=[SD2 + "/D - Plan Set", SD2 + "/D - Plan Set/Approved Plans & Docs"],
         why="PP: pre-contract references (our early drawings, client documents, downloaded sets). P: approved and permit sets. The 5 approved files left in PPs moved to their P."),
    dict(name="Other & misc", home="both", state="settled", label="Settled",
         pp=[SD + "/Z - Other Misc"], p=[SD2 + "/Z - Other Misc"],
         why="A catch-all on both sides. The PP one also holds 109 blank lien waivers."),
    dict(name="Resources", home="PP", state="open", label="Open · Batch 9, last",
         pp=["4 - Resources"], chips={"4 - Resources": "waits for Batch 9"}, p=[],
         p_note="None. P0 is the office's to change.",
         why="One <code>Resource hub.md</code> per proposal, written nightly: this project's links, the folder guide and company resources."),
    dict(name="Takeoff & estimate", home="P", state="settled", label="Settled",
         pp=[], p=["1 - Takeoff & Estimate", "1 - Takeoff & Estimate/Approved Project Estimate",
                   "1 - Takeoff & Estimate/Approved Project Estimate/V1", "@DIV"],
         pp_note="Not in PP0 since 2026-10-27. Six old proposals (PP1–PP3, PP5, PP6, PP110) keep a legacy <code>1 - Takeoff &amp; Estimate</code> with their files.",
         why="Takeoff drawings, sub bids and approved estimates."),
    dict(name="Receipts", home="P", state="settled", label="Settled",
         pp=[], p=["3 - Receipts"], pp_note="None, correctly.",
         why="PM's job costs file receipts here."),
    dict(name="Closeout", home="P", state="settled", label="Settled",
         pp=[], p=["4 - Administrative/2 - Construction Issues", "4 - Administrative/3 - Warranty",
                   "4 - Administrative/4 - Notice Of Completion", "4 - Administrative/5 - Certificate of Occupancy"],
         pp_note="None, correctly.",
         why="End-of-project folders. Few projects have reached closeout, so low use here means not yet, not never."),
]

LEFT = [
    ("", "Still to do · after the meeting with Sales and owners"),
    ("9", "One <code>Resource hub.md</code> in each proposal's <code>4 - Resources</code>, written nightly beside the CRM history. Shape decided 2026-10-28."),
    ("4", "Retire <code>B - Questionnaire &amp; Selection</code>: approved, held. First the document-templates thread takes the blank forms out of PP0."),
    ("2", "Change orders stay in the PP for now. Revisit with Sales and owners."),
    ("", "Still to do · yours, any time"),
    ("·", "CRM history heartbeat: create a healthchecks.io check, then set <code>HEARTBEAT_URL_PROPOSAL_CRM_HISTORY</code> in Vercel."),
    ("·", "The 6,435 walkthrough copies in P folders (43.7 GB): removing them needs your yes and quarantine tooling."),
    ("·", "Three misfiled items from the plan-set audit, and the September worklist (duplicates, litter, folder names, empty proposals): ProposalAudit README, Open items."),
    ("", "Done"),
    ("1", "<code>1 - CRM</code> and <code>4 - Resources</code> in every proposal and PP0, 2026-10-27."),
    ("5", "<code>C - Received from Client</code> flattened in every proposal and every project, both templates included, 2026-10-28."),
    ("6", "Plan sets audited; the 5 approved files left in PPs moved to their P, 2026-10-28."),
    ("3", "P→PP walkthrough shortcut in all 63 projects, from PM's bridge, 2026-10-28."),
    ("7", "Nightly CRM history in <code>1 - CRM</code>, live 2026-10-28."),
    ("8", "Dropped: client addresses in the Tools repo."),
]

STATUS_CLASS = {"used": "used", "underused": "under", "abandoned": "aband", "never used": "never"}
esc = html.escape


def folder_row(r, total, side, chip=None):
    parents, leaf = r["path"].split("/")[:-1], r["path"].split("/")[-1]
    s = STATUS_CLASS[r["status"]]
    share = round(100 * r["recent_share"])
    facts = ['<b>%d</b> of %d' % (r["props_deep"], total), '%d%% of last 12 months' % share]
    if r["newest"]:
        facts.append("newest " + r["newest"])
    tags = []
    if r["template_copies"]:
        tags.append('<span class="tag">%s blank template cop%s</span>'
                    % (format(r["template_copies"], ","), "y" if r["template_copies"] == 1 else "ies"))
    if r["container"] and r["path"]:
        tags.append('<span class="tag">incl. subfolders</span>')
    samples = ""
    if r["samples"]:
        lis = "".join(
            '<li><span class="num">%s%s</span><span class="d">%s</span><span class="f">%s</span>%s</li>'
            % (side, s_["num"], s_["day"], esc(s_["file"]),
               '<span class="w">subfolder</span>' if s_["where"] else "")
            for s_ in r["samples"])
        samples = '<details><summary>Sample files</summary><ul class="samples">%s</ul></details>' % lis
    parent = ('<span class="parent">%s /</span>' % esc(" / ".join(p_.split(" - ", 1)[0] for p_ in parents))
              if parents else "")
    chip_html = ('<span class="chip ghost">%s</span>' % esc(chip)) if chip else \
        '<span class="chip">%s</span>' % esc(r["status"])
    return ('<div class="frow s-%s" style="--depth:%d">'
            '<div class="fhead"><div class="fname">%s<span class="leaf">%s</span></div>%s</div>'
            '<div class="meter" role="img" aria-label="%d%% of recent folders"><i style="width:%d%%"></i></div>'
            '<div class="facts">%s</div>%s%s</div>'
            % (s, min(len(parents), 3), parent, esc(leaf), chip_html, share,
               max(share, 1 if r["props_deep"] else 0), " · ".join(facts), "".join(tags), samples))


def div_group():
    div = [r for r in p["rows"] if "3rd Party Coordination & Bids/" in r["path"]]
    counts = {k: sum(1 for r in div if r["status"] == k) for k in STATUS_CLASS}
    bar = "".join('<i class="seg s-%s" style="flex:%d" title="%s: %d"></i>'
                  % (STATUS_CLASS[k], v, k, v) for k, v in counts.items() if v)
    files = sum(r["files_direct"] for r in div)
    projects = P["1 - Takeoff & Estimate/3rd Party Coordination & Bids"]["props_deep"]
    rows = "".join(
        '<li class="s-%s"><span class="chip">%s</span><span class="dn">%s</span><span class="dc">%s</span></li>'
        % (STATUS_CLASS[r["status"]], esc(r["status"]), esc(r["path"].split("/")[2]),
           ("%d file%s" % (r["files_deep"], "" if r["files_deep"] == 1 else "s")) if r["files_deep"] else "")
        for r in div if r["path"].count("/") == 2)
    sub_never = sum(1 for r in div if r["path"].count("/") > 2 and r["status"] == "never used")
    return ('<div class="frow group" style="--depth:1">'
            '<div class="fhead"><div class="fname"><span class="parent">1 /</span>'
            '<span class="leaf">3rd Party Coordination &amp; Bids · DIV-01…DIV-20</span></div>'
            '<span class="chip ghost">%d folders</span></div>'
            '<div class="segbar">%s</div>'
            '<div class="facts"><b>%d</b> of %d projects, %d files in all · %s</div>'
            '<details><summary>Each division</summary><ul class="divs">%s</ul>'
            '<p class="fine">Their 50 <code>1 - Material</code> / <code>2 - Labor</code> subfolders: %d never used.</p></details></div>'
            % (len(div), bar, projects, p["n_folders"], files,
               ", ".join("%d %s" % (v, k) for k, v in counts.items() if v), rows, sub_never))


def cell(t, side):
    rows, src, total = (t["pp"], PP, pp["n_folders"]) if side == "PP" else (t["p"], P, p["n_folders"])
    note = t.get("pp_note" if side == "PP" else "p_note")
    chips = t.get("chips", {})
    out = [div_group() if path == "@DIV" else folder_row(src[path], total, side, chips.get(path))
           for path in rows]
    if note:
        out.append('<p class="note">%s</p>' % note)
    if not out:
        out.append('<p class="note">None.</p>')
    label = "Proposal · PP0" if side == "PP" else "Project · P0"
    return '<div class="cell %s"><span class="side-label">%s</span>%s</div>' % (side.lower(), label, "".join(out))


ARROW = {"PP": ("←", "Proposal"), "P": ("→", "Project"), "both": ("↔", "Both")}


def spine(t):
    a, word = ARROW[t["home"]]
    return ('<div class="spine home-%s"><h3>%s</h3>'
            '<div class="home"><span class="arrow" aria-hidden="true">%s</span>Home: <b>%s</b>'
            '<span class="state st-%s">%s</span></div>'
            '<p>%s</p></div>' % (t["home"].lower(), esc(t["name"]), a, word, t["state"], esc(t["label"]), t["why"]))


themes_html = "".join('<section class="theme %s">%s%s%s</section>'
                      % (t["state"], cell(t, "PP"), spine(t), cell(t, "P")) for t in THEMES)
left_html = "".join(('<li><span class="bn">%s</span><span>%s</span></li>' % b) if b[0]
                    else '<li class="ph"><span></span><span>%s</span></li>' % b[1] for b in LEFT)
n_open = sum(1 for t in THEMES if t["state"] == "open")

page = open(os.path.join(HERE, "census_page_template.html"), encoding="utf-8").read()
for k, v in {"{{THEMES}}": themes_html, "{{BATCHES}}": left_html, "{{N_PP}}": str(pp["n_folders"]),
             "{{N_P}}": str(p["n_folders"]), "{{R_PP}}": str(pp["n_recent"]), "{{R_P}}": str(p["n_recent"]),
             "{{N_OPEN}}": "%d of %d" % (n_open, len(THEMES)), "{{READ_ON}}": READ_ON}.items():
    assert k in page, k
    page = page.replace(k, v)
open(OUT, "w", encoding="utf-8").write(page)
print("wrote", OUT, len(page), "bytes;", n_open, "open themes")
