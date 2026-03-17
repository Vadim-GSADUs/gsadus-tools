"""Merged Whitespace tab — Crop / Pad / Audit modes with optional aspect ratio enforcement."""
import threading
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import colorchooser, filedialog
from pathlib import Path

import ttkbootstrap as tb
from PIL import Image as PILImage, ImageTk

import core.whitespace_audit as wsa
import core.whitespace_crop as wsc
import core.whitespace_pad as wsp
from ui.base_tab import BaseTab

_THUMB_W, _THUMB_H = 260, 190

_MODE_LABELS = {
    "crop":  "Crop Images",
    "pad":   "Pad Images",
    "audit": "Audit Margins",
}

# (display label, width/height ratio)
RATIO_PRESETS: list[tuple[str, float | None]] = [
    ("16 : 9   —  Widescreen  (web / presentations)",    16 / 9),
    ("3 : 2    —  Standard photography / print",          3 / 2),
    ("4 : 3    —  Traditional architectural photo",       4 / 3),
    ("1 : 1    —  Square  (social media)",                1.0),
    ("2 : 1    —  Panoramic  (building facade)",          2.0),
    ("3 : 1    —  Wide panoramic  (skyline / urban)",     3.0),
    ("4 : 5    —  Portrait  (tall building)",             4 / 5),
    ("2 : 3    —  Portrait  (magazine / print)",          2 / 3),
    ("φ  1.618 : 1  —  Golden ratio",                    1.618),
    ("√2 : 1  —  A-series paper  (A4 / A3 landscape)",   2 ** 0.5),
    ("Custom", None),
]
_PRESET_LABELS  = [label for label, _ in RATIO_PRESETS]
_PRESET_BY_LABEL = {label: ratio for label, ratio in RATIO_PRESETS}


def _parse_ratio(s: str) -> float | None:
    """Accept '16:9', '16/9', or '1.777'. Returns None on invalid input."""
    s = s.strip()
    if not s:
        return None
    for sep in (":", "/"):
        if sep in s:
            parts = s.split(sep, 1)
            try:
                return float(parts[0].strip()) / float(parts[1].strip())
            except (ValueError, ZeroDivisionError):
                return None
    try:
        return float(s)
    except ValueError:
        return None


def _margin_group(parent, items: list[tuple[str, str, int]], vars_dict: dict) -> tb.Frame:
    """Render 4 labelled int entries in a single row, registering vars into vars_dict."""
    frame = tb.Frame(parent)
    for i, (key, label, default) in enumerate(items):
        var = tk.StringVar(value=str(default))
        vars_dict[key] = var
        tb.Label(frame, text=label).grid(row=0, column=i * 2, padx=(12 if i else 0, 4), pady=6)
        tb.Entry(frame, textvariable=var, width=7).grid(row=0, column=i * 2 + 1, padx=(0, 8), pady=6)
    return frame


class WhitespaceTab(BaseTab):
    def __init__(self, parent, root):
        super().__init__(parent, root, config_key="whitespace", run_label="Crop Images")

        # ── Mode selector ──────────────────────────────────────────────────
        self._mode = tk.StringVar(value="crop")
        self._vars["mode"] = self._mode

        mode_row = tb.Frame(self._ctrl_frame)
        mode_row.grid(row=self._ctrl_row, column=0, columnspan=3, sticky="w", padx=4, pady=(10, 6))
        for text, val in [("  Crop  ", "crop"), ("  Pad  ", "pad"), ("  Audit  ", "audit")]:
            tb.Radiobutton(
                mode_row, text=text, variable=self._mode, value=val,
                bootstyle="toolbutton-outline", command=self._on_mode_change,
            ).pack(side="left", padx=(0, 4))
        self._ctrl_row += 1

        # ── Shared controls ────────────────────────────────────────────────
        self._add_section("Paths")
        self._add_path_row("folder", "Folder")
        self._add_checkbox("recursive", "Recursive (include subfolders)", default=False)

        self._add_section("Content Detection")
        self._add_int_row("white_tolerance", "White Tolerance  (0–255)", default=10)
        self._add_int_row("alpha_threshold", "Alpha Threshold  (0–255)", default=0)

        # ── Aspect ratio section (shared, always visible) ──────────────────
        self._add_section("Aspect Ratio")
        self._make_aspect_ratio_section()

        # ── Mode-specific sections (same grid row, toggled) ────────────────
        self._mode_row = self._ctrl_row
        self._ctrl_row += 1

        self._crop_frame  = self._make_crop_frame()
        self._pad_frame   = self._make_pad_frame()
        self._audit_frame = self._make_audit_frame()
        for f in (self._crop_frame, self._pad_frame, self._audit_frame):
            f.grid(row=self._mode_row, column=0, columnspan=3, sticky="ew", padx=4, pady=6)

        self._finish_build()
        self._on_mode_change()

    # ── Aspect ratio section ───────────────────────────────────────────────

    def _make_aspect_ratio_section(self):
        self._vars["use_aspect_ratio"] = tk.BooleanVar(value=False)
        tb.Checkbutton(
            self._ctrl_frame,
            text="Enforce / check aspect ratio  (pads to target ratio — never crops content)",
            variable=self._vars["use_aspect_ratio"],
            bootstyle="round-toggle",
            command=self._on_aspect_toggle,
        ).grid(row=self._ctrl_row, column=0, columnspan=3, sticky="w", padx=4, pady=(2, 4))
        self._ctrl_row += 1

        # Container shown/hidden by the toggle
        self._aspect_inner = tb.Frame(self._ctrl_frame)
        self._aspect_inner.grid(row=self._ctrl_row, column=0, columnspan=3,
                                 sticky="ew", padx=20, pady=(0, 6))
        self._ctrl_row += 1

        # Preset dropdown
        self._vars["_aspect_preset"] = tk.StringVar(value=_PRESET_LABELS[0])
        combo = ttk.Combobox(
            self._aspect_inner, textvariable=self._vars["_aspect_preset"],
            values=_PRESET_LABELS, state="readonly", width=46,
        )
        combo.grid(row=0, column=0, columnspan=3, sticky="w", padx=(0, 8), pady=4)
        combo.bind("<<ComboboxSelected>>", lambda _e: self._on_preset_change())

        # Custom entry (shown only when "Custom" is selected)
        self._custom_label = tb.Label(self._aspect_inner, text="Custom  (e.g. 16:9 or 1.777):")
        self._custom_label.grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self._vars["_aspect_custom"] = tk.StringVar()
        self._custom_entry = tb.Entry(
            self._aspect_inner, textvariable=self._vars["_aspect_custom"], width=16)
        self._custom_entry.grid(row=1, column=1, sticky="w", pady=4)

        # Ratio pad colour (for Crop / Pad modes)
        self._ratio_color_rgba = (255, 255, 255, 255)
        self._vars["_ratio_color"] = tk.StringVar(value="255,255,255,255")
        self._vars["ratio_transparent"] = tk.BooleanVar(value=False)

        color_row = tb.Frame(self._aspect_inner)
        color_row.grid(row=2, column=0, columnspan=3, sticky="w", pady=4)
        tb.Label(color_row, text="Padding colour:").pack(side="left", padx=(0, 8))
        self._ratio_swatch = tk.Label(
            color_row, bg="#ffffff", width=5, relief="solid", bd=1)
        self._ratio_swatch.pack(side="left", padx=(0, 8), ipady=4)
        tb.Button(color_row, text="Pick", bootstyle="secondary-outline", width=6,
                  command=self._pick_ratio_color).pack(side="left", padx=(0, 12))
        tb.Checkbutton(color_row, text="Transparent",
                       variable=self._vars["ratio_transparent"],
                       bootstyle="round-toggle").pack(side="left")

        # Start hidden; toggled by checkbox
        self._aspect_inner.grid_remove()

    def _on_aspect_toggle(self):
        if self._vars["use_aspect_ratio"].get():
            self._aspect_inner.grid()
            self._on_preset_change()
        else:
            self._aspect_inner.grid_remove()

    def _on_preset_change(self):
        is_custom = self._vars["_aspect_preset"].get() == "Custom"
        if is_custom:
            self._custom_label.grid()
            self._custom_entry.grid()
        else:
            self._custom_label.grid_remove()
            self._custom_entry.grid_remove()

    def _pick_ratio_color(self):
        result = colorchooser.askcolor(
            color=self._ratio_color_rgba[:3], title="Pick aspect ratio pad colour")
        if result and result[0]:
            r, g, b = tuple(int(c) for c in result[0])
            a = 0 if self._vars["ratio_transparent"].get() else 255
            self._ratio_color_rgba = (r, g, b, a)
            self._vars["_ratio_color"].set(f"{r},{g},{b},{a}")
            self._ratio_swatch.configure(bg="#{:02x}{:02x}{:02x}".format(r, g, b))

    def _resolved_target_ratio(self) -> float | None:
        if not self._vars["use_aspect_ratio"].get():
            return None
        preset = self._vars["_aspect_preset"].get()
        if preset == "Custom":
            return _parse_ratio(self._vars["_aspect_custom"].get())
        return _PRESET_BY_LABEL.get(preset)

    # ── Mode-specific section builders ────────────────────────────────────

    def _make_crop_frame(self) -> ttk.LabelFrame:
        lf = ttk.LabelFrame(self._ctrl_frame, text="Crop Settings")
        tb.Label(lf, text="Target buffer in px — image is cropped to content bbox + these margins",
                 foreground="gray").grid(row=0, column=0, sticky="w", padx=8, pady=(6, 2))
        _margin_group(lf, [
            ("crop_left",   "Left",   200), ("crop_right",  "Right",  200),
            ("crop_top",    "Top",    200), ("crop_bottom", "Bottom", 200),
        ], self._vars).grid(row=1, column=0, sticky="w", padx=4)
        self._vars["crop_overwrite"] = tk.BooleanVar(value=True)
        tb.Checkbutton(lf, text="Overwrite original files",
                       variable=self._vars["crop_overwrite"],
                       bootstyle="round-toggle").grid(row=2, column=0, sticky="w", padx=8, pady=6)
        return lf

    def _make_pad_frame(self) -> ttk.LabelFrame:
        lf = ttk.LabelFrame(self._ctrl_frame, text="Pad Settings")
        tb.Label(lf, text="Minimum margin in px — padding added only where current margin is less",
                 foreground="gray").grid(row=0, column=0, columnspan=4, sticky="w", padx=8, pady=(6, 2))
        _margin_group(lf, [
            ("pad_left",   "Left",   50), ("pad_right",  "Right",  50),
            ("pad_top",    "Top",    50), ("pad_bottom", "Bottom", 50),
        ], self._vars).grid(row=1, column=0, sticky="w", padx=4)

        self._pad_color_rgba = (255, 255, 255, 255)
        self._vars["_pad_color"] = tk.StringVar(value="255,255,255,255")
        self._vars["pad_transparent"] = tk.BooleanVar(value=False)

        color_row = tb.Frame(lf)
        color_row.grid(row=2, column=0, sticky="w", padx=8, pady=4)
        tb.Label(color_row, text="Pad colour:").pack(side="left", padx=(0, 8))
        self._pad_swatch = tk.Label(color_row, bg="#ffffff", width=5, relief="solid", bd=1)
        self._pad_swatch.pack(side="left", padx=(0, 8), ipady=4)
        tb.Button(color_row, text="Pick", bootstyle="secondary-outline", width=6,
                  command=self._pick_pad_color).pack(side="left", padx=(0, 12))
        tb.Checkbutton(color_row, text="Transparent",
                       variable=self._vars["pad_transparent"],
                       bootstyle="round-toggle").pack(side="left")

        self._vars["pad_overwrite"] = tk.BooleanVar(value=True)
        tb.Checkbutton(lf, text="Overwrite original files",
                       variable=self._vars["pad_overwrite"],
                       bootstyle="round-toggle").grid(row=3, column=0, sticky="w", padx=8, pady=(0, 6))
        return lf

    def _make_audit_frame(self) -> ttk.LabelFrame:
        lf = ttk.LabelFrame(self._ctrl_frame, text="Audit Settings")
        tb.Label(lf, text="Minimum required margins — files below these values are reported as FAIL",
                 foreground="gray").grid(row=0, column=0, sticky="w", padx=8, pady=(6, 2))
        _margin_group(lf, [
            ("audit_left",   "Left",   50), ("audit_right",  "Right",  50),
            ("audit_top",    "Top",    50), ("audit_bottom", "Bottom", 25),
        ], self._vars).grid(row=1, column=0, sticky="w", padx=4)

        report_row = tb.Frame(lf)
        report_row.grid(row=2, column=0, sticky="w", padx=8, pady=6)
        tb.Label(report_row, text="Report filename:").pack(side="left", padx=(0, 8))
        self._vars["report_name"] = tk.StringVar(value="_PNG_WhitespaceAudit.csv")
        tb.Entry(report_row, textvariable=self._vars["report_name"], width=30).pack(side="left")
        return lf

    # ── Colour pickers ─────────────────────────────────────────────────────

    def _pick_pad_color(self):
        result = colorchooser.askcolor(color=self._pad_color_rgba[:3], title="Pick pad colour")
        if result and result[0]:
            r, g, b = tuple(int(c) for c in result[0])
            a = 0 if self._vars["pad_transparent"].get() else 255
            self._pad_color_rgba = (r, g, b, a)
            self._vars["_pad_color"].set(f"{r},{g},{b},{a}")
            self._pad_swatch.configure(bg="#{:02x}{:02x}{:02x}".format(r, g, b))

    # ── Live Preview ───────────────────────────────────────────────────────

    def _on_build_preview(self):
        outer = ttk.LabelFrame(self, text="Live Preview")
        outer.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 4))
        outer.grid_columnconfigure(0, weight=1)

        # Top row: image path picker + refresh button
        ctrl = tb.Frame(outer)
        ctrl.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ctrl.grid_columnconfigure(1, weight=1)

        tb.Label(ctrl, text="Preview Image:", width=16, anchor="w").grid(
            row=0, column=0, padx=(0, 6))
        self._preview_path_var = tk.StringVar()
        tb.Entry(ctrl, textvariable=self._preview_path_var).grid(
            row=0, column=1, sticky="ew", padx=(0, 4))
        tb.Button(ctrl, text="Browse", bootstyle="secondary-outline", width=8,
                  command=self._browse_preview).grid(row=0, column=2, padx=(0, 8))
        tb.Button(ctrl, text="↻ Refresh", bootstyle="info-outline", width=10,
                  command=self._refresh_preview).grid(row=0, column=3)

        # Thumbnails: Before | After
        thumbs = tb.Frame(outer)
        thumbs.grid(row=1, column=0)

        for col, title in enumerate(("Before", "After")):
            tb.Label(thumbs, text=title, anchor="center").grid(
                row=0, column=col, padx=20)
            lbl = tk.Label(thumbs, bg="#2b2b2b")
            lbl.grid(row=1, column=col, padx=20, pady=2)
            if col == 0:
                self._preview_before_lbl = lbl
            else:
                self._preview_after_lbl = lbl

        self._preview_status_lbl = tb.Label(
            outer, text="Select a preview image above", foreground="gray")
        self._preview_status_lbl.grid(row=2, column=0, sticky="w", pady=(4, 0))

        # Init placeholder thumbnails
        self._photo_before: ImageTk.PhotoImage | None = None
        self._photo_after:  ImageTk.PhotoImage | None = None
        self._preview_job:  str | None = None
        self._preview_busy  = False
        self._set_placeholder(self._preview_before_lbl)
        self._set_placeholder(self._preview_after_lbl)

    def _set_placeholder(self, lbl: tk.Label):
        ph = PILImage.new("RGB", (_THUMB_W, _THUMB_H), (43, 43, 43))
        photo = ImageTk.PhotoImage(ph)
        lbl.configure(image=photo)
        lbl._photo = photo  # keep reference

    def _make_thumb(self, pil_img: PILImage.Image) -> ImageTk.PhotoImage:
        thumb = pil_img.copy()
        thumb.thumbnail((_THUMB_W, _THUMB_H), PILImage.LANCZOS)
        canvas = PILImage.new("RGB", (_THUMB_W, _THUMB_H), (43, 43, 43))
        x = (_THUMB_W - thumb.width) // 2
        y = (_THUMB_H - thumb.height) // 2
        if "A" in thumb.getbands():
            canvas.paste(thumb.convert("RGB"), (x, y), thumb.split()[-1])
        else:
            canvas.paste(thumb.convert("RGB"), (x, y))
        return ImageTk.PhotoImage(canvas)

    def _browse_preview(self):
        result = filedialog.askopenfilename(
            title="Select Preview Image",
            filetypes=[("PNG files", "*.png"), ("All files", "*.*")],
        )
        if result:
            self._preview_path_var.set(result)

    def _on_preview_ready(self):
        self._preview_path_var.trace_add("write", lambda *_: self._schedule_preview())
        for var in self._vars.values():
            var.trace_add("write", lambda *_: self._schedule_preview())

    def _schedule_preview(self, delay: int = 500):
        if not hasattr(self, "_preview_job"):
            return
        if self._preview_job is not None:
            self._win.after_cancel(self._preview_job)
        self._preview_job = self._win.after(delay, self._refresh_preview)

    def _refresh_preview(self):
        self._preview_job = None
        if not hasattr(self, "_preview_path_var"):
            return
        path_str = self._preview_path_var.get().strip()
        if not path_str:
            return
        p = Path(path_str)
        if not p.is_file():
            self._preview_status_lbl.configure(text=f"File not found: {p.name}")
            return
        if self._preview_busy:
            return
        self._preview_busy = True
        self._preview_status_lbl.configure(text="Computing…")

        mode = self._mode.get()
        kwargs = self._get_preview_kwargs()

        def worker():
            try:
                with PILImage.open(p) as im:
                    before = im.copy()
                if mode == "crop":
                    after = wsc.preview_image(p, **kwargs)
                elif mode == "pad":
                    after = wsp.preview_image(p, **kwargs)
                else:
                    after = None  # audit — no visual transform
                self._win.after(0, lambda: self._set_preview(before, after, mode))
            except Exception as e:
                self._win.after(0, lambda msg=str(e): self._preview_status_lbl.configure(
                    text=f"Error: {msg}"))
            finally:
                self._preview_busy = False

        threading.Thread(target=worker, daemon=True).start()

    def _get_preview_kwargs(self) -> dict:
        def i(k): return int(self._vars[k].get() or 0)
        mode = self._mode.get()
        target_ratio = self._resolved_target_ratio()
        try:
            ratio_bg = tuple(int(p.strip()) for p in self._vars["_ratio_color"].get().split(","))
        except Exception:
            ratio_bg = (255, 255, 255, 255)
        shared = dict(white_tolerance=i("white_tolerance"), alpha_threshold=i("alpha_threshold"))
        if mode == "crop":
            return {**shared,
                    "buffer_left": i("crop_left"), "buffer_right": i("crop_right"),
                    "buffer_top":  i("crop_top"),  "buffer_bottom": i("crop_bottom"),
                    "target_ratio": target_ratio, "ratio_bg": ratio_bg,
                    "ratio_transparent": self._vars["ratio_transparent"].get()}
        if mode == "pad":
            try:
                pad_color = tuple(int(p.strip()) for p in self._vars["_pad_color"].get().split(","))
            except Exception:
                pad_color = (255, 255, 255, 255)
            return {**shared,
                    "min_left": i("pad_left"), "min_right": i("pad_right"),
                    "min_top":  i("pad_top"),  "min_bottom": i("pad_bottom"),
                    "pad_color": pad_color, "pad_transparent": self._vars["pad_transparent"].get(),
                    "target_ratio": target_ratio, "ratio_bg": ratio_bg,
                    "ratio_transparent": self._vars["ratio_transparent"].get()}
        return shared  # audit

    def _set_preview(self, before: PILImage.Image, after: PILImage.Image | None, mode: str):
        self._photo_before = self._make_thumb(before)
        self._preview_before_lbl.configure(image=self._photo_before)
        bw, bh = before.size
        if after is not None:
            self._photo_after = self._make_thumb(after)
            self._preview_after_lbl.configure(image=self._photo_after)
            aw, ah = after.size
            status = f"Before: {bw}×{bh}  →  After: {aw}×{ah}"
        else:
            self._set_placeholder(self._preview_after_lbl)
            if mode == "audit":
                status = f"Audit mode — no image transform  ({bw}×{bh})"
            else:
                status = f"No change needed  ({bw}×{bh})"
        self._preview_status_lbl.configure(text=status)

    # ── Mode toggle ────────────────────────────────────────────────────────

    def _on_mode_change(self):
        mode = self._mode.get()
        frame_map = {"crop": self._crop_frame, "pad": self._pad_frame, "audit": self._audit_frame}
        for m, f in frame_map.items():
            f.grid() if m == mode else f.grid_remove()
        self._run_label_default = _MODE_LABELS[mode]
        if hasattr(self, "_run_btn"):
            self._run_btn.configure(text=_MODE_LABELS[mode])
        self._schedule_preview(delay=100)

    # ── Config restore ─────────────────────────────────────────────────────

    def _on_config_loaded(self):
        # Restore pad swatch
        try:
            parts = self._vars["_pad_color"].get().split(",")
            self._pad_color_rgba = tuple(int(p.strip()) for p in parts)
            r, g, b = self._pad_color_rgba[:3]
            self._pad_swatch.configure(bg="#{:02x}{:02x}{:02x}".format(r, g, b))
        except Exception:
            pass
        # Restore ratio swatch
        try:
            parts = self._vars["_ratio_color"].get().split(",")
            self._ratio_color_rgba = tuple(int(p.strip()) for p in parts)
            r, g, b = self._ratio_color_rgba[:3]
            self._ratio_swatch.configure(bg="#{:02x}{:02x}{:02x}".format(r, g, b))
        except Exception:
            pass
        # Restore aspect ratio visibility
        self._on_aspect_toggle()
        self._on_preset_change()
        # Restore run button label
        mode = self._vars.get("mode")
        if mode:
            self._run_label_default = _MODE_LABELS.get(mode.get(), "Run")

    # ── Run ────────────────────────────────────────────────────────────────

    def _get_run_kwargs(self) -> dict:
        def i(k): return int(self._vars[k].get() or 0)

        mode = self._mode.get()
        target_ratio = self._resolved_target_ratio()

        try:
            parts = self._vars["_ratio_color"].get().split(",")
            ratio_bg = tuple(int(p.strip()) for p in parts)
        except Exception:
            ratio_bg = (255, 255, 255, 255)

        shared = dict(
            folder=self._vars["folder"].get(),
            recursive=self._vars["recursive"].get(),
            white_tolerance=i("white_tolerance"),
            alpha_threshold=i("alpha_threshold"),
        )

        if mode == "crop":
            return {**shared,
                    "overwrite":       self._vars["crop_overwrite"].get(),
                    "buffer_left":     i("crop_left"),
                    "buffer_right":    i("crop_right"),
                    "buffer_top":      i("crop_top"),
                    "buffer_bottom":   i("crop_bottom"),
                    "target_ratio":    target_ratio,
                    "ratio_bg":        ratio_bg,
                    "ratio_transparent": self._vars["ratio_transparent"].get()}

        if mode == "pad":
            try:
                parts = self._vars["_pad_color"].get().split(",")
                pad_color = tuple(int(p.strip()) for p in parts)
            except Exception:
                pad_color = (255, 255, 255, 255)
            return {**shared,
                    "overwrite":       self._vars["pad_overwrite"].get(),
                    "min_left":        i("pad_left"),
                    "min_right":       i("pad_right"),
                    "min_top":         i("pad_top"),
                    "min_bottom":      i("pad_bottom"),
                    "pad_color":       pad_color,
                    "pad_transparent": self._vars["pad_transparent"].get(),
                    "target_ratio":    target_ratio,
                    "ratio_bg":        ratio_bg,
                    "ratio_transparent": self._vars["ratio_transparent"].get()}

        # audit
        return {**shared,
                "required_left":   i("audit_left"),
                "required_right":  i("audit_right"),
                "required_top":    i("audit_top"),
                "required_bottom": i("audit_bottom"),
                "report_name":     self._vars["report_name"].get() or "_PNG_WhitespaceAudit.csv",
                "target_ratio":    target_ratio}

    def _run(self, **kwargs):
        mode = self._mode.get()
        if mode == "crop":
            wsc.run(**kwargs)
        elif mode == "pad":
            wsp.run(**kwargs)
        else:
            wsa.run(**kwargs)
