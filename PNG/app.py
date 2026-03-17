"""PNG Post-Processing Tools — main entry point."""
import ttkbootstrap as tb

from ui.tab_filename_normalization import FilenameNormalizationTab
from ui.tab_transparent_crop import TransparentCropTab
from ui.tab_whitespace import WhitespaceTab

TABS = [
    ("Whitespace",            WhitespaceTab),
    ("Transparent Crop",      TransparentCropTab),
    ("Filename Normalization", FilenameNormalizationTab),
]


def main():
    root = tb.Window(title="PNG Post-Processing Tools", themename="darkly")
    root.geometry("960x780")
    root.minsize(720, 620)

    notebook = tb.Notebook(root, bootstyle="dark")
    notebook.pack(fill="both", expand=True, padx=12, pady=12)

    for name, cls in TABS:
        frame = tb.Frame(notebook)
        notebook.add(frame, text=f"  {name}  ")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        cls(frame, root).grid(row=0, column=0, sticky="nsew")

    root.mainloop()


if __name__ == "__main__":
    main()
