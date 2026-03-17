"""BaseTab — shared scaffold for all workflow tabs."""
import sys
import threading
import tkinter as tk
from queue import Empty, Queue
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText

import ttkbootstrap as tb
from ttkbootstrap.widgets.scrolled import ScrolledFrame

import config


class _QueueWriter:
    """Redirect sys.stdout to a Queue so background threads stream print() to the UI."""

    def __init__(self, q: Queue):
        self._q = q

    def write(self, s: str):
        if s.strip():
            self._q.put(s.rstrip("\n"))

    def flush(self):
        pass


class BaseTab(tb.Frame):
    """
    Base class for all workflow tabs.

    Subclass pattern:
        def __init__(self, parent, root):
            super().__init__(parent, root, config_key="my_tab")
            self._add_section("Paths")
            self._add_path_row("folder", "Input Folder")
            ...
            self._finish_build()          # must be the last call

        def _get_run_kwargs(self) -> dict: ...
        def _run(self, **kwargs): ...      # calls core module; may use print()
    """

    def __init__(self, parent, root: tb.Window, config_key: str, run_label: str = "Run"):
        super().__init__(parent)
        self._win = root               # kept as _win to avoid shadowing tkinter's _root()
        self._config_key = config_key
        self._run_label_default = run_label
        self._vars: dict[str, tk.Variable] = {}
        self._running = False

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)   # scrollable controls
        self.grid_rowconfigure(1, weight=0)   # preview panel (optional hook)
        self.grid_rowconfigure(2, weight=0)   # log label
        self.grid_rowconfigure(3, weight=0)   # log textbox
        self.grid_rowconfigure(4, weight=0)   # run button

        self._ctrl_frame = ScrolledFrame(self, autohide=True)
        self._ctrl_frame.grid(row=0, column=0, sticky="nsew", padx=12, pady=(12, 4))
        self._ctrl_frame.grid_columnconfigure(1, weight=1)
        self._ctrl_row = 0

    # ── Control builder helpers ────────────────────────────────────────────

    def _add_path_row(
        self, key: str, label: str, mode: str = "dir", save: bool = False
    ) -> tk.StringVar:
        """Add a labelled path-picker row. mode='dir' or 'file'."""
        var = tk.StringVar()
        self._vars[key] = var
        tb.Label(self._ctrl_frame, text=label, width=22, anchor="w").grid(
            row=self._ctrl_row, column=0, sticky="w", padx=(4, 8), pady=5
        )
        tb.Entry(self._ctrl_frame, textvariable=var).grid(
            row=self._ctrl_row, column=1, sticky="ew", padx=(0, 5), pady=5
        )
        tb.Button(
            self._ctrl_frame, text="Browse", bootstyle="secondary-outline", width=8,
            command=lambda v=var, m=mode, s=save: self._browse(v, m, s),
        ).grid(row=self._ctrl_row, column=2, padx=(0, 4), pady=5)
        self._ctrl_row += 1
        return var

    def _browse(self, var: tk.StringVar, mode: str, save: bool):
        if mode == "dir":
            result = filedialog.askdirectory(title="Select Folder")
        elif save:
            result = filedialog.asksaveasfilename(
                title="Save As", defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            )
        else:
            result = filedialog.askopenfilename(
                title="Select File",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            )
        if result:
            var.set(result)

    def _add_checkbox(self, key: str, label: str, default: bool = False) -> tk.BooleanVar:
        var = tk.BooleanVar(value=default)
        self._vars[key] = var
        tb.Checkbutton(self._ctrl_frame, text=label, variable=var, bootstyle="round-toggle").grid(
            row=self._ctrl_row, column=0, columnspan=3, sticky="w", padx=4, pady=5
        )
        self._ctrl_row += 1
        return var

    def _add_int_row(
        self, key: str, label: str, default: int = 0, width: int = 10
    ) -> tk.StringVar:
        var = tk.StringVar(value=str(default))
        self._vars[key] = var
        tb.Label(self._ctrl_frame, text=label, width=22, anchor="w").grid(
            row=self._ctrl_row, column=0, sticky="w", padx=(4, 8), pady=5
        )
        tb.Entry(self._ctrl_frame, textvariable=var, width=width).grid(
            row=self._ctrl_row, column=1, sticky="w", padx=(0, 5), pady=5
        )
        self._ctrl_row += 1
        return var

    def _add_text_row(
        self, key: str, label: str, default: str = "", width: int = 36
    ) -> tk.StringVar:
        var = tk.StringVar(value=default)
        self._vars[key] = var
        tb.Label(self._ctrl_frame, text=label, width=22, anchor="w").grid(
            row=self._ctrl_row, column=0, sticky="w", padx=(4, 8), pady=5
        )
        tb.Entry(self._ctrl_frame, textvariable=var, width=width).grid(
            row=self._ctrl_row, column=1, sticky="w", padx=(0, 5), pady=5
        )
        self._ctrl_row += 1
        return var

    def _add_section(self, title: str):
        tb.Label(
            self._ctrl_frame, text=title,
            font=("TkDefaultFont", 10, "bold"), anchor="w",
            bootstyle="inverse-secondary",
        ).grid(row=self._ctrl_row, column=0, columnspan=3, sticky="ew", padx=4, pady=(14, 4))
        self._ctrl_row += 1

    def _add_inline_int_group(self, items: list[tuple[str, str, int]]):
        """Render several int fields side-by-side. items = [(key, label, default), ...]."""
        frame = tb.Frame(self._ctrl_frame)
        frame.grid(row=self._ctrl_row, column=0, columnspan=3, sticky="w", padx=4, pady=5)
        for i, (key, label, default) in enumerate(items):
            var = tk.StringVar(value=str(default))
            self._vars[key] = var
            tb.Label(frame, text=label).grid(row=0, column=i * 2, padx=(16 if i else 0, 4))
            tb.Entry(frame, textvariable=var, width=7).grid(row=0, column=i * 2 + 1, padx=(0, 8))
        self._ctrl_row += 1

    # ── Log + Run area ─────────────────────────────────────────────────────

    def _finish_build(self):
        """Call last in subclass __init__ — adds the preview hook, log area, and Run button."""
        self._on_build_preview()  # hook: subclasses place a preview panel at row 1

        tb.Label(
            self, text="Output Log", anchor="w",
            font=("TkDefaultFont", 10, "bold"),
        ).grid(row=2, column=0, sticky="w", padx=16, pady=(6, 0))

        self._log_box = ScrolledText(
            self, height=7, state="disabled",
            font=("Consolas", 11),
            bg="#2b2b2b", fg="#d4d4d4",
            relief="flat", bd=1,
            wrap="word",
        )
        self._log_box.grid(row=3, column=0, sticky="ew", padx=12, pady=(2, 6))

        self._run_btn = tb.Button(
            self, text=self._run_label_default,
            command=self._on_run, bootstyle="success", width=20,
        )
        self._run_btn.grid(row=4, column=0, pady=(0, 12))

        self._load_config()
        self._on_preview_ready()  # hook: attach traces after config values are set

    def log(self, msg: str):
        """Thread-safe: append a line to the log widget."""
        def _append():
            self._log_box.configure(state="normal")
            self._log_box.insert("end", msg + "\n")
            self._log_box.see("end")
            self._log_box.configure(state="disabled")
        self._win.after(0, _append)

    def _clear_log(self):
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

    # ── Threading ──────────────────────────────────────────────────────────

    def _on_run(self):
        if self._running:
            return
        self._running = True
        self._run_btn.configure(state="disabled", text="Running…")
        self._clear_log()
        self._save_config()

        kwargs = self._get_run_kwargs()
        q: Queue[str | None] = Queue()

        def worker():
            old_stdout = sys.stdout
            sys.stdout = _QueueWriter(q)
            try:
                self._run(**kwargs)
            except Exception as e:
                q.put(f"ERROR: {e}")
            finally:
                sys.stdout = old_stdout
                q.put(None)  # sentinel

        threading.Thread(target=worker, daemon=True).start()
        self._win.after(50, lambda: self._drain(q))

    def _drain(self, q: Queue):
        while True:
            try:
                msg = q.get_nowait()
            except Empty:
                break
            if msg is None:
                self._running = False
                self._run_btn.configure(state="normal", text=self._run_label_default)
                return
            self.log(msg)
        self._win.after(50, lambda: self._drain(q))

    # ── Config persistence ─────────────────────────────────────────────────

    def _load_config(self):
        data = config.load().get(self._config_key, {})
        for k, v in data.items():
            if k in self._vars:
                try:
                    self._vars[k].set(v)
                except Exception:
                    pass
        self._on_config_loaded()

    def _on_config_loaded(self):
        """Override to perform UI updates after config values are applied."""

    def _on_build_preview(self):
        """Override to place a preview panel at row 1 of the tab grid."""

    def _on_preview_ready(self):
        """Override to attach variable traces after config is loaded."""

    def _save_config(self):
        data = {k: v.get() for k, v in self._vars.items()}
        cfg = config.load()
        cfg[self._config_key] = data
        config.save(cfg)

    # ── Abstract ───────────────────────────────────────────────────────────

    def _get_run_kwargs(self) -> dict:
        raise NotImplementedError

    def _run(self, **kwargs):
        raise NotImplementedError
