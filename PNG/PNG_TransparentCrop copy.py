import cv2
import numpy as np
from pathlib import Path
from collections import deque

# ====== CONFIG: edit these ======
FOLDER = r"G:\Shared drives\GSADUs Projects\Our Models\0 - CATALOG\Output"

BACKGROUND_THRESHOLD = 245  # 0–255, near-white background candidate
BLUR_KERNEL_SIZE = 3        # must be odd; 0 or 1 disables blur
ERODE_ITERS = 0             # optional tweak of final mask
DILATE_ITERS = 1

# ====== SCRIPT ======

def get_border_connected_background_mask(gray_img, threshold):
    """
    Returns a mask (uint8, 0/255) of background pixels that are:
      - brighter than 'threshold' AND
      - connected to the image border.
    """
    h, w = gray_img.shape

    # 1) Background candidates: near-white pixels
    #    255 = candidate background, 0 = foreground
    _, bg_candidates = cv2.threshold(
        gray_img, threshold, 255, cv2.THRESH_BINARY
    )

    # 2) BFS from border over candidate pixels
    visited = np.zeros((h, w), dtype=bool)
    border_bg = np.zeros((h, w), dtype=np.uint8)
    q = deque()

    # Seed from all border pixels that are candidate background
    for x in range(w):
        if bg_candidates[0, x] == 255:
            q.append((0, x))
        if bg_candidates[h - 1, x] == 255:
            q.append((h - 1, x))
    for y in range(h):
        if bg_candidates[y, 0] == 255:
            q.append((y, 0))
        if bg_candidates[y, w - 1] == 255:
            q.append((y, w - 1))

    # 4-connected neighbors
    neighbors = [(-1, 0), (1, 0), (0, -1), (0, 1)]

    while q:
        y, x = q.popleft()
        if visited[y, x]:
            continue
        visited[y, x] = True
        border_bg[y, x] = 255

        for dy, dx in neighbors:
            ny, nx = y + dy, x + dx
            if 0 <= ny < h and 0 <= nx < w:
                if not visited[ny, nx] and bg_candidates[ny, nx] == 255:
                    q.append((ny, nx))

    return border_bg


def process_image(path: Path):
    print(f"Processing: {path}")

    img = cv2.imread(str(path), cv2.IMREAD_UNCHANGED)
    if img is None:
        print("  Skipped (could not read).")
        return

    # Handle channels
    if img.ndim == 2:  # grayscale
        bgr = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
        alpha_in = None
    elif img.shape[2] == 3:
        bgr = img
        alpha_in = None
    else:  # BGRA
        bgr = img[:, :, :3]
        alpha_in = img[:, :, 3]

    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)

    if BLUR_KERNEL_SIZE and BLUR_KERNEL_SIZE > 1:
        gray_blur = cv2.GaussianBlur(gray, (BLUR_KERNEL_SIZE, BLUR_KERNEL_SIZE), 0)
    else:
        gray_blur = gray

    # Get mask of background connected to border only
    border_bg_mask = get_border_connected_background_mask(
        gray_blur, BACKGROUND_THRESHOLD
    )

    # Foreground = everything NOT border background
    fg_mask = cv2.bitwise_not(border_bg_mask)

    # Optional morphology on foreground mask
    if ERODE_ITERS > 0:
        fg_mask = cv2.erode(fg_mask, None, iterations=ERODE_ITERS)
    if DILATE_ITERS > 0:
        fg_mask = cv2.dilate(fg_mask, None, iterations=DILATE_ITERS)

    # Combine with any existing alpha
    if alpha_in is not None:
        combined_alpha = (alpha_in.astype(np.float32) / 255.0) * (
            fg_mask.astype(np.float32) / 255.0
        )
        alpha_out = (combined_alpha * 255.0).astype(np.uint8)
    else:
        alpha_out = fg_mask

    bgra = cv2.cvtColor(bgr, cv2.COLOR_BGR2BGRA)
    bgra[:, :, 3] = alpha_out

    cv2.imwrite(str(path), bgra)
    print("  Saved (overwritten).")


def main():
    folder = Path(FOLDER)
    if not folder.is_dir():
        raise SystemExit(f"Folder does not exist: {folder}")

    png_files = list(folder.glob("*.png"))
    if not png_files:
        print(f"No .png files found in {folder}")
        return

    print(f"Found {len(png_files)} PNG(s) in {folder}")
    for p in png_files:
        process_image(p)


if __name__ == "__main__":
    main()
