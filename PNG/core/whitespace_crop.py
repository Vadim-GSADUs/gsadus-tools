"""Crop uniform whitespace/transparent borders from PNGs with configurable per-side buffers."""
from pathlib import Path

from PIL import Image

from core._image_utils import content_bbox_rgb, content_bbox_rgba, is_uniform, pad_to_ratio


def _expand_bbox(bbox, w: int, h: int, bl: int, bt: int, br: int, bb: int):
    if not bbox:
        return None
    left, top, right, bottom = bbox
    left = max(0, left - bl)
    top = max(0, top - bt)
    right = min(w, right + br)
    bottom = min(h, bottom + bb)
    return (left, top, right, bottom) if right > left and bottom > top else None


def _trim_file(
    path: Path,
    overwrite: bool,
    buffer_left: int,
    buffer_top: int,
    buffer_right: int,
    buffer_bottom: int,
    white_tolerance: int,
    alpha_threshold: int,
    target_ratio: float | None,
    ratio_bg: tuple,
    ratio_transparent: bool,
) -> str:
    try:
        with Image.open(path) as im:
            w, h = im.size
            if is_uniform(im):
                return f"skip uniform: {path.name}"
            bbox = content_bbox_rgba(im, alpha_threshold) if "A" in im.getbands() else None
            if not bbox:
                bbox = content_bbox_rgb(im, white_tolerance)
            if not bbox:
                return f"skip no-content: {path.name}"
            bbox = _expand_bbox(bbox, w, h, buffer_left, buffer_top, buffer_right, buffer_bottom)
            if not bbox:
                return f"skip invalid-bbox: {path.name}"

            result = im.crop(bbox) if bbox != (0, 0, w, h) else im.copy()

            if target_ratio is not None:
                result = pad_to_ratio(result, target_ratio, ratio_bg, ratio_transparent)

            out_path = path if overwrite else path.with_name(path.stem + "._trimmed" + path.suffix)
            result.save(out_path)
            return f"ok: {path.name}"
    except Exception as e:
        return f"error: {path.name}: {e}"


def run(
    folder,
    *,
    recursive: bool = False,
    only_png: bool = True,
    overwrite: bool = True,
    buffer_left: int = 200,
    buffer_right: int = 200,
    buffer_top: int = 200,
    buffer_bottom: int = 200,
    white_tolerance: int = 10,
    alpha_threshold: int = 0,
    target_ratio: float | None = None,
    ratio_bg: tuple = (255, 255, 255, 255),
    ratio_transparent: bool = False,
) -> None:
    """
    Crop whitespace/transparent borders from PNGs in folder.
    If target_ratio is set, the cropped image is padded symmetrically to reach
    that aspect ratio (width / height) — content is never clipped.
    """
    root = Path(folder)
    patterns = ["*.png"] if only_png else ["*.png", "*.PNG"]
    files: list[Path] = []
    if recursive:
        for ptn in patterns:
            files.extend(root.rglob(ptn))
    else:
        for ptn in patterns:
            files.extend(root.glob(ptn))

    if not files:
        print("No PNG files found.")
        return

    for f in sorted(files):
        print(_trim_file(f, overwrite, buffer_left, buffer_top, buffer_right, buffer_bottom,
                         white_tolerance, alpha_threshold, target_ratio, ratio_bg, ratio_transparent))


def preview_image(
    path,
    *,
    white_tolerance: int = 10,
    alpha_threshold: int = 0,
    buffer_left: int = 200,
    buffer_top: int = 200,
    buffer_right: int = 200,
    buffer_bottom: int = 200,
    target_ratio: float | None = None,
    ratio_bg: tuple = (255, 255, 255, 255),
    ratio_transparent: bool = False,
) -> "Image.Image | None":
    """Return the processed image without saving, or None if skipped/error."""
    try:
        with Image.open(path) as im:
            im = im.copy()
        w, h = im.size
        if is_uniform(im):
            return None
        bbox = content_bbox_rgba(im, alpha_threshold) if "A" in im.getbands() else None
        if not bbox:
            bbox = content_bbox_rgb(im, white_tolerance)
        if not bbox:
            return None
        bbox = _expand_bbox(bbox, w, h, buffer_left, buffer_top, buffer_right, buffer_bottom)
        if not bbox:
            return None
        result = im.crop(bbox) if bbox != (0, 0, w, h) else im.copy()
        if target_ratio is not None:
            result = pad_to_ratio(result, target_ratio, ratio_bg, ratio_transparent)
        return result
    except Exception:
        return None
