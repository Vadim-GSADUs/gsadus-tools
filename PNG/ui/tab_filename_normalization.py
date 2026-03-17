import threading
import tkinter as tk
import tkinter.ttk as ttk

import ttkbootstrap as tb

import core.filename_normalization as fn
from ui.base_tab import BaseTab

_CASE_LABELS  = ["No change", "Title Case", "lowercase", "UPPERCASE"]
_CASE_VALUES  = ["none",      "title",      "lower",     "upper"]
_SPACE_LABELS = ["Keep spaces", "Underscore  _", "Hyphen  -", "Remove spaces"]
_SPACE_VALUES = ["keep",        "_",             "-",          "remove"]


class FilenameNormalizationTab(BaseTab):
    def __init__(self, parent, root):
        super().__init__(parent, root, config_key="filename_normalization",
                         run_label="Preview (Dry Run)")

        # ── Paths ──────────────────────────────────────────────────────────
        self._add_section("Paths")
        self._add_path_row("folder", "Folder")

        # ── Run Options ────────────────────────────────────────────────────
        self._add_section("Run Options")
        dry_var = self._add_checkbox(
            "dry_run", "Dry Run — preview only, no files renamed  (safe default)", default=True)
        self._add_checkbox("recursive", "Recursive (include subfolders)", default=True)

        # ── Clean-up ───────────────────────────────────────────────────────
        self._add_section("Clean-up")
        self._build_delimiter_row()
        self._build_strip_strings_row()
        self._build_find_replace_rows()
        self._add_checkbox("strip_digits",  "Remove leading / trailing digits from name", default=False)
        self._add_checkbox("strip_special", "Strip special characters  (keep letters, digits, spaces, - _)", default=False)

        # ── Formatting ─────────────────────────────────────────────────────
        self._add_section("Formatting")
        self._build_combobox_row("case_mode",      "Case:",          _CASE_LABELS,  default=0)
        self._build_combobox_row("replace_spaces", "Spaces:",        _SPACE_LABELS, default=0)

        # ── Structure ──────────────────────────────────────────────────────
        self._add_section("Structure")
        self._add_checkbox("truncate",   "Truncate to minimal unique word prefix", default=True)
        self._add_checkbox("add_folder", "Append parent folder name to filename",  default=True)
        self._build_prefix_suffix_row()

        # ── Numbering ──────────────────────────────────────────────────────
        self._add_section("Sequential Numbering")
        self._build_sequential_rows()

        dry_var.trace_add("write", lambda *_: self._update_run_label())
        self._finish_build()

    # ── Custom row builders ────────────────────────────────────────────────

    def _build_delimiter_row(self):
        """[toggle] Strip after delimiter: [entry]"""
        row = tb.Frame(self._ctrl_frame)
        row.grid(row=self._ctrl_row, column=0, columnspan=3, sticky="w", padx=4, pady=5)
        self._ctrl_row += 1
        var = tk.BooleanVar(value=True)
        self._vars["strip_after_delimiter"] = var
        tb.Checkbutton(row, text="Strip after delimiter:", variable=var,
                       bootstyle="round-toggle").pack(side="left", padx=(0, 10))
        dvar = tk.StringVar(value=" - ")
        self._vars["delimiter"] = dvar
        tb.Entry(row, textvariable=dvar, width=8,
                 font=("Consolas", 10)).pack(side="left")

    def _build_strip_strings_row(self):
        """Multi-line text area — each non-empty line is stripped from filenames."""
        outer = tb.Frame(self._ctrl_frame)
        outer.grid(row=self._ctrl_row, column=0, columnspan=3, sticky="ew", padx=4, pady=5)
        self._ctrl_row += 1
        outer.grid_columnconfigure(1, weight=1)

        tb.Label(outer, text="Strip string(s):", anchor="nw").grid(
            row=0, column=0, sticky="nw", padx=(0, 8), pady=(2, 0))

        text_frame = tb.Frame(outer)
        text_frame.grid(row=0, column=1, sticky="ew")
        text_frame.grid_columnconfigure(0, weight=1)

        self._strip_text = tk.Text(
            text_frame, height=3, font=("Consolas", 10),
            bg="#1e1e2e", fg="#d4d4d4", insertbackground="#d4d4d4",
            relief="flat", bd=1, wrap="none",
        )
        vsb = ttk.Scrollbar(text_frame, orient="vertical", command=self._strip_text.yview)
        self._strip_text.configure(yscrollcommand=vsb.set)
        self._strip_text.grid(row=0, column=0, sticky="ew")
        vsb.grid(row=0, column=1, sticky="ns")

        tb.Label(outer, text="one per line", foreground="gray", font=("TkDefaultFont", 8)).grid(
            row=1, column=1, sticky="w", pady=(2, 0))

        # Keep a StringVar in _vars so config persistence works automatically
        self._vars["strip_strings"] = tk.StringVar(value="")

        def _sync_to_var(event=None):
            self._vars["strip_strings"].set(self._strip_text.get("1.0", "end-1c"))
            self._schedule_rename_preview()

        self._strip_text.bind("<KeyRelease>", _sync_to_var)
        self._strip_text.bind("<<Paste>>", lambda e: self._win.after(1, _sync_to_var))

    def _build_find_replace_rows(self):
        """[toggle] Find & Replace  then  Find:[__] → Replace:[__]  [case-sensitive]"""
        fr_var = self._add_checkbox("find_replace", "Find & Replace", default=False)

        sub = tb.Frame(self._ctrl_frame)
        sub.grid(row=self._ctrl_row, column=0, columnspan=3, sticky="w", padx=(28, 4), pady=(0, 5))
        self._ctrl_row += 1
        self._fr_sub = sub

        tb.Label(sub, text="Find:").pack(side="left", padx=(0, 4))
        self._vars["find_str"] = tk.StringVar()
        tb.Entry(sub, textvariable=self._vars["find_str"], width=18).pack(side="left", padx=(0, 8))
        tb.Label(sub, text="→  Replace:").pack(side="left", padx=(0, 4))
        self._vars["replace_str"] = tk.StringVar()
        tb.Entry(sub, textvariable=self._vars["replace_str"], width=18).pack(side="left", padx=(0, 12))
        self._vars["find_case_sensitive"] = tk.BooleanVar(value=False)
        tb.Checkbutton(sub, text="Case-sensitive",
                       variable=self._vars["find_case_sensitive"],
                       bootstyle="round-toggle").pack(side="left")

        fr_var.trace_add("write", lambda *_: self._toggle_fr_sub())
        self._toggle_fr_sub()

    def _toggle_fr_sub(self):
        if self._vars["find_replace"].get():
            self._fr_sub.grid()
        else:
            self._fr_sub.grid_remove()

    def _build_combobox_row(self, key: str, label: str, display_values: list, default: int = 0):
        var = tk.StringVar(value=display_values[default])
        self._vars[key] = var
        row = tb.Frame(self._ctrl_frame)
        row.grid(row=self._ctrl_row, column=0, columnspan=3, sticky="w", padx=4, pady=5)
        self._ctrl_row += 1
        tb.Label(row, text=label, width=14, anchor="w").pack(side="left", padx=(0, 8))
        ttk.Combobox(row, textvariable=var, values=display_values,
                     state="readonly", width=22).pack(side="left")

    def _build_prefix_suffix_row(self):
        """Prefix:[___]  Suffix:[___] on one row."""
        row = tb.Frame(self._ctrl_frame)
        row.grid(row=self._ctrl_row, column=0, columnspan=3, sticky="w", padx=4, pady=5)
        self._ctrl_row += 1
        tb.Label(row, text="Prefix:").pack(side="left", padx=(0, 4))
        self._vars["prefix"] = tk.StringVar()
        tb.Entry(row, textvariable=self._vars["prefix"], width=16).pack(side="left", padx=(0, 20))
        tb.Label(row, text="Suffix:").pack(side="left", padx=(0, 4))
        self._vars["suffix"] = tk.StringVar()
        tb.Entry(row, textvariable=self._vars["suffix"], width=16).pack(side="left")

    def _build_sequential_rows(self):
        seq_var = self._add_checkbox("sequential", "Add sequential number to each filename", default=False)

        sub = tb.Frame(self._ctrl_frame)
        sub.grid(row=self._ctrl_row, column=0, columnspan=3, sticky="w", padx=(28, 4), pady=(0, 6))
        self._ctrl_row += 1
        self._seq_sub = sub

        # Start / Step / Padding (inline)
        for label, key, default in [("Start:", "seq_start", "1"),
                                     ("Step:",  "seq_step",  "1"),
                                     ("Padding:", "seq_padding", "3")]:
            tb.Label(sub, text=label).pack(side="left", padx=(0, 4))
            var = tk.StringVar(value=default)
            self._vars[key] = var
            tb.Entry(sub, textvariable=var, width=5).pack(side="left", padx=(0, 14))

        tb.Label(sub, text="Position:").pack(side="left", padx=(0, 8))
        self._vars["seq_position"] = tk.StringVar(value="Prefix")
        for txt in ("Prefix", "Suffix"):
            tb.Radiobutton(sub, text=txt, variable=self._vars["seq_position"], value=txt,
                           bootstyle="toolbutton-outline").pack(side="left", padx=(0, 4))

        seq_var.trace_add("write", lambda *_: self._toggle_seq_sub())
        self._toggle_seq_sub()

    def _toggle_seq_sub(self):
        if self._vars["sequential"].get():
            self._seq_sub.grid()
        else:
            self._seq_sub.grid_remove()

    # ── Config restore ─────────────────────────────────────────────────────

    def _on_config_loaded(self):
        # Restore strip strings text widget from the persisted StringVar
        content = self._vars["strip_strings"].get()
        self._strip_text.delete("1.0", "end")
        if content:
            self._strip_text.insert("1.0", content)
        self._toggle_fr_sub()
        self._toggle_seq_sub()
        self._update_run_label()

    def _update_run_label(self):
        label = "Preview (Dry Run)" if self._vars["dry_run"].get() else "Rename Files"
        self._run_label_default = label
        if hasattr(self, "_run_btn"):
            self._run_btn.configure(text=label)

    # ── Rename Preview ─────────────────────────────────────────────────────

    def _on_build_preview(self):
        outer = ttk.LabelFrame(self, text="Rename Preview")
        outer.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 4))
        outer.grid_columnconfigure(0, weight=1)

        tree_frame = tb.Frame(outer)
        tree_frame.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 2))
        tree_frame.grid_columnconfigure(0, weight=1)

        self._rename_tree = ttk.Treeview(
            tree_frame, columns=("old", "new"), show="headings", height=10)
        self._rename_tree.heading("old", text="Current Name")
        self._rename_tree.heading("new", text="New Name")
        self._rename_tree.column("old", width=390, minwidth=160)
        self._rename_tree.column("new", width=390, minwidth=160)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._rename_tree.yview)
        self._rename_tree.configure(yscrollcommand=vsb.set)
        self._rename_tree.grid(row=0, column=0, sticky="ew")
        vsb.grid(row=0, column=1, sticky="ns")

        self._rename_status_lbl = tb.Label(
            outer, text="Select a folder to preview renames", foreground="gray")
        self._rename_status_lbl.grid(row=1, column=0, sticky="w", padx=6, pady=(0, 6))

        self._rename_job: str | None = None

    def _on_preview_ready(self):
        for var in self._vars.values():
            var.trace_add("write", lambda *_: self._schedule_rename_preview())

    def _schedule_rename_preview(self, delay: int = 400):
        if not hasattr(self, "_rename_job"):
            return
        if self._rename_job is not None:
            self._win.after_cancel(self._rename_job)
        self._rename_job = self._win.after(delay, self._refresh_rename_preview)

    def _refresh_rename_preview(self):
        self._rename_job = None
        folder = self._vars["folder"].get().strip()
        if not folder:
            return
        kwargs = self._get_preview_kwargs()
        self._rename_status_lbl.configure(text="Computing…")

        def worker():
            try:
                pairs = fn.get_renames(**kwargs)
                self._win.after(0, lambda: self._set_rename_preview(pairs))
            except Exception as e:
                self._win.after(0, lambda msg=str(e): self._rename_status_lbl.configure(
                    text=f"Error: {msg}"))

        threading.Thread(target=worker, daemon=True).start()

    def _set_rename_preview(self, pairs: list[tuple[str, str]]):
        for row in self._rename_tree.get_children():
            self._rename_tree.delete(row)
        for old, new in pairs:
            self._rename_tree.insert("", "end", values=(old, new))
        n = len(pairs)
        self._rename_status_lbl.configure(
            text=f"{n} file(s) will be renamed" if n else "No renames needed")

    # ── Kwargs helpers ─────────────────────────────────────────────────────

    def _resolve_kwargs(self) -> dict:
        """Resolve display labels → internal values and validate integers."""
        def _int(key, fallback):
            try:
                return int(self._vars[key].get() or fallback)
            except ValueError:
                return fallback

        case_val  = _CASE_VALUES[_CASE_LABELS.index(self._vars["case_mode"].get())] \
                    if self._vars["case_mode"].get() in _CASE_LABELS else "none"
        space_val = _SPACE_VALUES[_SPACE_LABELS.index(self._vars["replace_spaces"].get())] \
                    if self._vars["replace_spaces"].get() in _SPACE_LABELS else "keep"

        return dict(
            folder=self._vars["folder"].get(),
            recursive=self._vars["recursive"].get(),
            add_folder=self._vars["add_folder"].get(),
            truncate=self._vars["truncate"].get(),
            strip_after_delimiter=self._vars["strip_after_delimiter"].get(),
            delimiter=self._vars["delimiter"].get(),
            strip_strings=[s for s in self._vars["strip_strings"].get().splitlines() if s.strip()],
            find_replace=self._vars["find_replace"].get(),
            find_str=self._vars["find_str"].get(),
            replace_str=self._vars["replace_str"].get(),
            find_case_sensitive=self._vars["find_case_sensitive"].get(),
            prefix=self._vars["prefix"].get(),
            suffix=self._vars["suffix"].get(),
            case_mode=case_val,
            replace_spaces=space_val,
            strip_special=self._vars["strip_special"].get(),
            strip_digits=self._vars["strip_digits"].get(),
            sequential=self._vars["sequential"].get(),
            seq_start=_int("seq_start", 1),
            seq_step=_int("seq_step", 1),
            seq_padding=_int("seq_padding", 3),
            seq_position="prefix" if self._vars["seq_position"].get() == "Prefix" else "suffix",
        )

    def _get_preview_kwargs(self) -> dict:
        return self._resolve_kwargs()

    def _get_run_kwargs(self) -> dict:
        kw = self._resolve_kwargs()
        kw["dry_run"] = self._vars["dry_run"].get()
        return kw

    def _run(self, **kwargs):
        fn.run(**kwargs)
