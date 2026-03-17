"""Standalone runner for transparent crop. Edit defaults below, or use app.py for a GUI."""
# ====== CONFIG: edit these ======
FOLDER              = r"G:\Shared drives\GSADUs Projects\Our Models\0 - CATALOG\Output"
USE_BORDER_CONNECTED = False
TARGET_COLOR_BGR    = (255, 255, 255)
COLOR_TOLERANCE     = 0
BLUR_KERNEL_SIZE    = 3
ERODE_ITERS         = 0
DILATE_ITERS        = 1
# ====== END CONFIG ======

if __name__ == "__main__":
    from core.transparent_crop import run
    run(
        FOLDER,
        use_border_connected=USE_BORDER_CONNECTED,
        target_color_bgr=TARGET_COLOR_BGR,
        color_tolerance=COLOR_TOLERANCE,
        blur_kernel_size=BLUR_KERNEL_SIZE,
        erode_iters=ERODE_ITERS,
        dilate_iters=DILATE_ITERS,
    )
