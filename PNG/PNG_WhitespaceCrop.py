"""Standalone runner for whitespace crop. Edit defaults below, or use app.py for a GUI."""
# ====== CONFIG: edit these ======
FOLDER        = r"G:\Shared drives\GSADUs Projects\Our Models\0 - CATALOG\Output"
BUFFER_LEFT   = 200
BUFFER_RIGHT  = 200
BUFFER_TOP    = 200
BUFFER_BOTTOM = 200
RECURSIVE      = False
WHITE_TOLERANCE = 10
ALPHA_THRESHOLD = 0
ONLY_PNG       = True
OVERWRITE      = True
# ====== END CONFIG ======

if __name__ == "__main__":
    from core.whitespace_crop import run
    run(
        FOLDER,
        recursive=RECURSIVE,
        only_png=ONLY_PNG,
        overwrite=OVERWRITE,
        buffer_left=BUFFER_LEFT,
        buffer_right=BUFFER_RIGHT,
        buffer_top=BUFFER_TOP,
        buffer_bottom=BUFFER_BOTTOM,
        white_tolerance=WHITE_TOLERANCE,
        alpha_threshold=ALPHA_THRESHOLD,
    )
