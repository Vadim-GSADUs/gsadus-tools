"""Standalone runner for whitespace audit. Edit defaults below, or use app.py for a GUI."""
# ====== CONFIG: edit these ======
FOLDER          = r"G:\Shared drives\GSADUs Projects\Our Models\0 - CATALOG\Working\Support\PNG"
REQUIRED_LEFT   = 50
REQUIRED_RIGHT  = 50
REQUIRED_TOP    = 50
REQUIRED_BOTTOM = 25
WHITE_TOLERANCE = 10
ALPHA_THRESHOLD = 0
RECURSIVE       = False
ONLY_PNG        = True
REPORT_NAME     = "_PNG_WhitespaceAudit.csv"
# ====== END CONFIG ======

if __name__ == "__main__":
    from core.whitespace_audit import run
    run(
        FOLDER,
        recursive=RECURSIVE,
        only_png=ONLY_PNG,
        required_left=REQUIRED_LEFT,
        required_right=REQUIRED_RIGHT,
        required_top=REQUIRED_TOP,
        required_bottom=REQUIRED_BOTTOM,
        white_tolerance=WHITE_TOLERANCE,
        alpha_threshold=ALPHA_THRESHOLD,
        report_name=REPORT_NAME,
    )
