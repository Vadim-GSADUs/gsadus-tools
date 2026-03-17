import tkinter as tk
from tkinter import colorchooser

import ttkbootstrap as tb

import core.transparent_crop as tc
from ui.base_tab import BaseTab


class TransparentCropTab(BaseTab):
    def __init__(self, parent, root):
        super().__init__(parent, root, config_key="transparent_crop")
        self._color_rgb = (255, 255, 255)

        self._add_section("Paths")
        self._add_path_row("folder", "Input Folder")

        self._add_section("Background Colour")
        # Hidden var stores colour as "R,G,B" for config persistence
        self._vars["_color_rgb"] = tk.StringVar(value="255,255,255")

        color_row = tb.Frame(self._ctrl_frame)
        color_row.grid(row=self._ctrl_row, column=0, columnspan=3, sticky="w", padx=4, pady=5)
        tb.Label(color_row, text="Target Colour", width=22, anchor="w").grid(
            row=0, column=0, padx=(0, 8))
        self._swatch = tk.Label(color_row, bg="#ffffff", width=5, relief="solid", bd=1)
        self._swatch.grid(row=0, column=1, padx=(0, 10), ipady=6)
        tb.Button(color_row, text="Pick Colour", bootstyle="secondary-outline",
                  command=self._pick_color).grid(row=0, column=2)
        self._ctrl_row += 1

        self._add_int_row("color_tolerance", "Colour Tolerance  (0–255)", default=0)

        self._add_section("Removal Mode")
        self._add_checkbox(
            "use_border_connected",
            "Border-connected only  (flood fill from edges — preserves interior detail)",
            default=False,
        )

        self._add_section("Morphology")
        self._add_int_row("blur_kernel_size", "Blur Kernel Size  (0 = off, must be odd)", default=3)
        self._add_int_row("erode_iters",      "Erode Iterations",  default=0)
        self._add_int_row("dilate_iters",     "Dilate Iterations", default=1)

        self._finish_build()

    def _pick_color(self):
        result = colorchooser.askcolor(color=self._color_rgb, title="Pick background colour to remove")
        if result and result[0]:
            rgb = tuple(int(c) for c in result[0])
            self._color_rgb = rgb
            self._vars["_color_rgb"].set(f"{rgb[0]},{rgb[1]},{rgb[2]}")
            self._swatch.configure(bg="#{:02x}{:02x}{:02x}".format(*rgb))

    def _on_config_loaded(self):
        try:
            parts = self._vars["_color_rgb"].get().split(",")
            self._color_rgb = tuple(int(p.strip()) for p in parts)
            self._swatch.configure(bg="#{:02x}{:02x}{:02x}".format(*self._color_rgb))
        except Exception:
            pass

    def _get_run_kwargs(self) -> dict:
        def i(k): return int(self._vars[k].get() or 0)
        try:
            parts = self._vars["_color_rgb"].get().split(",")
            r, g, b = tuple(int(p.strip()) for p in parts)
        except Exception:
            r, g, b = 255, 255, 255
        return dict(
            folder=self._vars["folder"].get(),
            use_border_connected=self._vars["use_border_connected"].get(),
            target_color_bgr=(b, g, r),   # UI stores RGB; cv2 needs BGR
            color_tolerance=i("color_tolerance"),
            blur_kernel_size=i("blur_kernel_size"),
            erode_iters=i("erode_iters"),
            dilate_iters=i("dilate_iters"),
        )

    def _run(self, **kwargs):
        tc.run(**kwargs)
