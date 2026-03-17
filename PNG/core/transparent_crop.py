"""Remove a background colour from PNGs, writing BGRA output with transparency."""
from collections import deque
from pathlib import Path

import cv2
import numpy as np


def _get_color_mask(img, target_bgr: tuple, tolerance: int):
    lower = np.array([max(0, c - tolerance) for c in target_bgr], dtype=np.uint8)
    upper = np.array([min(255, c + tolerance) for c in target_bgr], dtype=np.uint8)
    return cv2.inRange(img, lower, upper)


def _get_border_connected_mask(mask):
    """Return only the region of `mask` that is 4-connected to the image border."""
    h, w = mask.shape
    visited = np.zeros((h, w), dtype=bool)
    border_connected = np.zeros((h, w), dtype=np.uint8)
    q = deque()

    for x in range(w):
        if mask[0, x] == 255:
            q.append((0, x))
        if mask[h - 1, x] == 255:
            q.append((h - 1, x))
    for y in range(h):
        if mask[y, 0] == 255:
            q.append((y, 0))
        if mask[y, w - 1] == 255:
            q.append((y, w - 1))

    neighbors = [(-1, 0), (1, 0), (0, -1), (0, 1)]
    while q:
        y, x = q.popleft()
        if visited[y, x]:
            continue
        visited[y, x] = True
        border_connected[y, x] = 255
        for dy, dx in neighbors:
            ny, nx = y + dy, x + dx
            if 0 <= ny < h and 0 <= nx < w and not visited[ny, nx] and mask[ny, nx] == 255:
                q.append((ny, nx))

    return border_connected


def _process_image(
    path: Path,
    use_border_connected: bool,
    target_bgr: tuple,
    color_tolerance: int,
    blur_kernel_size: int,
    erode_iters: int,
    dilate_iters: int,
):
    img = cv2.imread(str(path), cv2.IMREAD_UNCHANGED)
    if img is None:
        print(f"  Skipped (could not read): {path.name}")
        return

    if img.ndim == 2:
        bgr, alpha_in = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR), None
    elif img.shape[2] == 3:
        bgr, alpha_in = img, None
    else:
        bgr, alpha_in = img[:, :, :3], img[:, :, 3]

    bgr_blur = (
        cv2.GaussianBlur(bgr, (blur_kernel_size, blur_kernel_size), 0)
        if blur_kernel_size > 1
        else bgr
    )

    color_mask = _get_color_mask(bgr_blur, target_bgr, color_tolerance)
    bg_mask = _get_border_connected_mask(color_mask) if use_border_connected else color_mask
    fg_mask = cv2.bitwise_not(bg_mask)

    if erode_iters > 0:
        fg_mask = cv2.erode(fg_mask, None, iterations=erode_iters)
    if dilate_iters > 0:
        fg_mask = cv2.dilate(fg_mask, None, iterations=dilate_iters)

    if alpha_in is not None:
        combined = (alpha_in.astype(np.float32) / 255.0) * (fg_mask.astype(np.float32) / 255.0)
        alpha_out = (combined * 255.0).astype(np.uint8)
    else:
        alpha_out = fg_mask

    bgra = cv2.cvtColor(bgr, cv2.COLOR_BGR2BGRA)
    bgra[:, :, 3] = alpha_out
    cv2.imwrite(str(path), bgra)
    print(f"  Saved: {path.name}")


def run(
    folder,
    *,
    use_border_connected: bool = False,
    target_color_bgr: tuple = (255, 255, 255),
    color_tolerance: int = 0,
    blur_kernel_size: int = 3,
    erode_iters: int = 0,
    dilate_iters: int = 1,
) -> None:
    """
    Remove target_color_bgr from all PNGs in folder, writing BGRA output.
    Files are overwritten in place.
    """
    root = Path(folder)
    if not root.is_dir():
        print(f"Folder not found: {root}")
        return

    png_files = list(root.glob("*.png"))
    if not png_files:
        print("No PNG files found.")
        return

    # Ensure blur kernel is odd
    if blur_kernel_size > 1 and blur_kernel_size % 2 == 0:
        blur_kernel_size += 1

    print(f"Processing {len(png_files)} PNG(s)…")
    for p in sorted(png_files):
        print(f"Processing: {p.name}")
        _process_image(p, use_border_connected, target_color_bgr, color_tolerance,
                       blur_kernel_size, erode_iters, dilate_iters)
    print("Done.")
