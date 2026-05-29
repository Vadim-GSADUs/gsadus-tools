"""
Revit Element Poison Finder
===========================

Find which element(s) in a large Revit selection prevent the whole selection
from moving, by LOGGING the outcome of many overlapping test selections and
letting set logic deduce the culprit.

Each test you log is one observation:
  - A selection that MOVED   (Clean) -> EVERY element in it is innocent.
  - A selection that LOCKED  (Dirty) -> AT LEAST ONE poison is inside it.

From the accumulated log the tool deduces:
  - Confirmed clean   = union of all Clean tests.
  - Remaining suspects= Collective - Confirmed clean.
  - Narrowing         = each Dirty test minus the clean elements; if that
                        reduces to one element, it is a CONFIRMED poison.
  - Overlap rule      = under a single culprit, the poison appears in EVERY
                        locked test, so candidates = intersection of all Dirty
                        tests minus clean. Each overlapping locked test shrinks
                        this set.
  - Contradiction     = a Dirty test whose elements are all proven clean is
                        impossible for a single poison -> likely MULTIPLE
                        poisons (or a mis-logged test).

Paste IDs in any format -- comma, space, vertical list, or raw Revit text like
'Elements of the current selection IDs are: 8744792, 8744643, 8745377'.
Every integer is extracted; everything else is ignored.

Run:  py element_filter.py
"""

import re
import tkinter as tk
from tkinter import ttk, messagebox

ID_RE = re.compile(r"\d+")


def parse_ids(text):
    """Extract integer IDs from arbitrary text, deduped, order preserved."""
    seen = set()
    out = []
    for tok in ID_RE.findall(text or ""):
        n = int(tok)
        if n not in seen:
            seen.add(n)
            out.append(n)
    return out


def join_ids(ids, sep_name):
    sep = ", " if sep_name == "comma" else "\n"
    return sep.join(str(i) for i in ids)


def deduce(collective, log):
    """Run the set deduction over the logged tests.

    collective : list[int]
    log        : list[(kind, list[int])] where kind is 'clean' or 'dirty'

    Returns a dict with the derived sets/flags (lists kept in collective order
    where it makes sense).
    """
    coll = list(dict.fromkeys(collective))  # dedup, keep order
    coll_set = set(coll)

    clean_set = set()
    dirty_tests = []  # list of (index, set(ids))
    n_moved = n_locked = 0
    for idx, (kind, ids) in enumerate(log, start=1):
        if kind == "clean":
            clean_set |= set(ids)
            n_moved += 1
        else:
            dirty_tests.append((idx, set(ids)))
            n_locked += 1

    # Remaining suspects in the pool.
    suspects = [i for i in coll if i not in clean_set]

    # Reduce each locked test by removing proven-clean elements.
    reduced = []          # list of (index, [ids]) with >=1 element
    contradictions = []   # indices of locked tests fully explained as clean
    confirmed = []        # poisons confirmed by a locked test of size 1
    for idx, dset in dirty_tests:
        r = [i for i in dset if i not in clean_set]
        if not r:
            contradictions.append(idx)
            continue
        reduced.append((idx, r))
        if len(r) == 1:
            confirmed.append(r[0])

    # Single-culprit candidates = intersection of all reduced locked tests.
    if reduced:
        inter = set(reduced[0][1])
        for _, r in reduced[1:]:
            inter &= set(r)
    else:
        inter = set()
    # Present intersection in collective order (fall back to sorted for any
    # ids that somehow are not in collective).
    narrowed = [i for i in coll if i in inter]
    narrowed += sorted(i for i in inter if i not in coll_set)

    confirmed = [i for i in coll if i in set(confirmed)] + \
                sorted(i for i in set(confirmed) if i not in coll_set)

    multi = bool(contradictions) or (len(reduced) >= 2 and not inter)

    return {
        "n_moved": n_moved,
        "n_locked": n_locked,
        "clean_count": len(clean_set & coll_set),
        "suspects": suspects,
        "narrowed": narrowed,
        "confirmed": confirmed,
        "contradictions": contradictions,
        "multi": multi,
        "reduced": reduced,
    }


class App:
    def __init__(self, root):
        self.root = root
        root.title("Revit Element Poison Finder")
        root.geometry("1120x760")
        root.minsize(960, 640)

        self.sep = tk.StringVar(value="comma")
        self.log = []            # list of (kind, [ids])
        self.cur_suspects = []
        self.cur_narrowed = []
        self.cur_confirmed = []

        self._build()
        self._recompute()

    # ---------- UI ----------
    def _build(self):
        root = self.root
        for c in (0, 1, 2):
            root.columnconfigure(c, weight=1)
        root.rowconfigure(1, weight=3)
        root.rowconfigure(5, weight=2)

        # Header
        hdr = ttk.Frame(root, padding=(10, 8))
        hdr.grid(row=0, column=0, columnspan=3, sticky="ew")
        ttk.Label(
            hdr,
            text="Paste the full pool once into COLLECTIVE. Each round, paste the "
            "selection you tested into CLEAN (it moved) or DIRTY (it locked), "
            "then click Test / Log.",
            font=("Segoe UI", 9), wraplength=820, justify="left",
        ).pack(side="left")
        ttk.Label(hdr, text="Output:").pack(side="left", padx=(16, 4))
        ttk.Radiobutton(hdr, text="comma", value="comma", variable=self.sep,
                        command=self._render_results).pack(side="left")
        ttk.Radiobutton(hdr, text="newline", value="newline", variable=self.sep,
                        command=self._render_results).pack(side="left")

        # Three paste buckets
        self.collective = self._box(root, 1, 0, "COLLECTIVE  (full pool — paste once)")
        self.clean = self._box(root, 1, 1, "CLEAN  (a selection that MOVED)")
        self.dirty = self._box(root, 1, 2, "DIRTY  (a selection that LOCKED)")
        self.collective.bind("<KeyRelease>", lambda e: self._recompute())

        self.c_count = self._count(root, 2, 0)
        self.cl_count = self._count(root, 2, 1)
        self.d_count = self._count(root, 2, 2)
        self.clean.bind("<KeyRelease>", lambda e: self._update_test_counts())
        self.dirty.bind("<KeyRelease>", lambda e: self._update_test_counts())

        # Action bar
        bar = ttk.Frame(root, padding=(10, 4))
        bar.grid(row=3, column=0, columnspan=3, sticky="ew")
        ttk.Button(bar, text="Test / Log",
                   command=self._log_test).pack(side="left")
        ttk.Button(bar, text="Undo last test",
                   command=self._undo_last).pack(side="left", padx=6)
        ttk.Button(bar, text="Clear log",
                   command=self._clear_log).pack(side="left", padx=6)
        ttk.Button(bar, text="Clear all",
                   command=self._clear_all).pack(side="left", padx=6)
        self.status = ttk.Label(bar, text="", foreground="#555")
        self.status.pack(side="right")

        ttk.Label(root, text="Tip: each test should overlap the previous ones — "
                  "shared elements are what let the logs triangulate the poison.",
                  foreground="#777", font=("Segoe UI", 8)).grid(
                      row=4, column=0, columnspan=3, sticky="w", padx=12)

        # Bottom: log (left) + results (right)
        bottom = ttk.Frame(root, padding=(10, 4))
        bottom.grid(row=5, column=0, columnspan=3, sticky="nsew")
        bottom.columnconfigure(0, weight=1)
        bottom.columnconfigure(1, weight=2)
        bottom.rowconfigure(0, weight=1)

        logf = ttk.LabelFrame(bottom, text="Test log", padding=4)
        logf.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        logf.rowconfigure(0, weight=1)
        logf.columnconfigure(0, weight=1)
        self.logbox = self._scrolled(logf, 0, 0, readonly=True)

        resf = ttk.LabelFrame(bottom, text="Deductions", padding=4)
        resf.grid(row=0, column=1, sticky="nsew")
        resf.rowconfigure(0, weight=1)
        resf.columnconfigure(0, weight=1)
        self.results = self._scrolled(resf, 0, 0, readonly=True)

        cb = ttk.Frame(resf)
        cb.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(cb, text="Copy suspects",
                   command=lambda: self._copy(self.cur_suspects)).pack(side="left")
        ttk.Button(cb, text="Copy candidates",
                   command=lambda: self._copy(self.cur_narrowed)).pack(side="left", padx=6)
        ttk.Button(cb, text="Copy confirmed poison",
                   command=lambda: self._copy(self.cur_confirmed)).pack(side="left")

    def _box(self, parent, r, c, title):
        frame = ttk.LabelFrame(parent, text=title, padding=4)
        frame.grid(row=r, column=c, sticky="nsew",
                   padx=(10 if c == 0 else 4, 10 if c == 2 else 4), pady=4)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return self._scrolled(frame, 0, 0)

    def _scrolled(self, parent, r, c, readonly=False, padx=0):
        wrap = ttk.Frame(parent)
        wrap.grid(row=r, column=c, sticky="nsew", padx=padx)
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        txt = tk.Text(wrap, wrap="word", font=("Consolas", 10), height=8, undo=True)
        txt.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=txt.yview)
        sb.grid(row=0, column=1, sticky="ns")
        txt.configure(yscrollcommand=sb.set)
        if readonly:
            txt.configure(state="disabled", background="#f4f4f4")
        return txt

    def _count(self, parent, r, c):
        lbl = ttk.Label(parent, text="0 ids", foreground="#555")
        lbl.grid(row=r, column=c, sticky="w", padx=14)
        return lbl

    # ---------- helpers ----------
    def _get(self, widget):
        return parse_ids(widget.get("1.0", "end"))

    def _set_ro(self, widget, text):
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        widget.configure(state="disabled")

    def _copy(self, ids):
        if not ids:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(join_ids(ids, self.sep.get()))

    def _update_test_counts(self):
        self.cl_count.config(text=f"{len(self._get(self.clean))} ids")
        self.d_count.config(text=f"{len(self._get(self.dirty))} ids")

    # ---------- actions ----------
    def _log_test(self):
        clean_ids = self._get(self.clean)
        dirty_ids = self._get(self.dirty)
        if not clean_ids and not dirty_ids:
            messagebox.showinfo(
                "Test / Log",
                "Paste the tested selection into CLEAN (it moved) or "
                "DIRTY (it locked) first.")
            return
        if clean_ids:
            self.log.append(("clean", clean_ids))
        if dirty_ids:
            self.log.append(("dirty", dirty_ids))
        self.clean.delete("1.0", "end")
        self.dirty.delete("1.0", "end")
        self._update_test_counts()
        self._recompute()

    def _undo_last(self):
        if not self.log:
            return
        self.log.pop()
        self._recompute()

    def _clear_log(self):
        if self.log and not messagebox.askyesno(
                "Clear log", "Discard all logged tests? (Collective is kept.)"):
            return
        self.log = []
        self._recompute()

    def _clear_all(self):
        if not messagebox.askyesno("Clear all", "Clear everything?"):
            return
        self.collective.delete("1.0", "end")
        self.clean.delete("1.0", "end")
        self.dirty.delete("1.0", "end")
        self.log = []
        self._update_test_counts()
        self._recompute()

    # ---------- compute & render ----------
    def _recompute(self, *_):
        coll = self._get(self.collective)
        self.c_count.config(text=f"{len(coll)} ids")
        self._update_test_counts()
        self.d = deduce(coll, self.log)
        self._render_log()
        self._render_results()

    def _render_log(self):
        lines = []
        for idx, (kind, ids) in enumerate(self.log, start=1):
            tag = "MOVED " if kind == "clean" else "LOCKED"
            preview = ", ".join(str(i) for i in ids[:8])
            if len(ids) > 8:
                preview += f", ... (+{len(ids) - 8})"
            lines.append(f"#{idx:<3} {tag}  ({len(ids):>3}):  {preview}")
        self._set_ro(self.logbox, "\n".join(lines))
        self.status.config(
            text=f"{self.d['n_moved']} moved / {self.d['n_locked']} locked logged")

    def _render_results(self):
        if not hasattr(self, "d"):
            return
        d = self.d
        sep = self.sep.get()
        self.cur_suspects = d["suspects"]
        self.cur_narrowed = d["narrowed"]
        self.cur_confirmed = d["confirmed"]

        lines = []
        lines.append(f"Tests logged:  {d['n_moved']} moved, {d['n_locked']} locked")
        lines.append(f"Confirmed CLEAN:  {d['clean_count']} elements")
        lines.append("")
        lines.append(f"Remaining suspects (Collective - Clean):  {len(d['suspects'])}")
        if d["suspects"]:
            lines.append("  " + join_ids(d["suspects"], "comma"))
        lines.append("")

        if d["n_locked"] == 0:
            lines.append("Log a LOCKED selection to start narrowing the poison.")
        else:
            lines.append(
                f"If a SINGLE element is the cause, it is one of these "
                f"{len(d['narrowed'])} (intersection of every locked test, "
                f"minus clean):")
            lines.append("  " + (join_ids(d["narrowed"], "comma")
                                 if d["narrowed"] else "(none — see warning below)"))

        if d["confirmed"]:
            lines.append("")
            lines.append(f">>> CONFIRMED poison(s): {join_ids(d['confirmed'], 'comma')}")
            lines.append("    (a locked test narrowed to exactly one element)")

        if d["contradictions"]:
            lines.append("")
            lines.append(
                "WARNING: locked test(s) #" +
                ", #".join(str(i) for i in d["contradictions"]) +
                " contain only proven-clean elements.")
            lines.append(
                "    Impossible for a single poison -> likely MULTIPLE poisons, "
                "or that test was mis-logged.")
        elif d["multi"]:
            lines.append("")
            lines.append(
                "WARNING: no single element appears in every locked test -> "
                "likely MULTIPLE poisons.")
            lines.append(
                "    Isolate one region at a time; the intersection only holds "
                "for a single culprit.")

        self._set_ro(self.results, "\n".join(lines))


def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
