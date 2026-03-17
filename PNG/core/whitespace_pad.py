"""Add whitespace padding to PNGs where any margin falls below a minimum."""
from pathlib import Path

from PIL import Image

from core._image_utils import content_bbox_rgb, content_bbox_rgba, is_uniform, pad_to_ratio


def _pad_file(
    path: Path,
    overwrite: bool,
    min_left: int,
    min_top: int,
    min_right: int,
    min_bottom: int,
    white_tolerance: int,
    alpha_threshold: int,
    pad_color: tuple,
    pad_transparent: bool,
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

            left, top, right, bottom = bbox
            extra_left   = max(0, min_left   - left)
            extra_top    = max(0, min_top    - top)
            extra_right  = max(0, min_right  - (w - right))
            extra_bottom = max(0, min_bottom - (h - bottom))

            has_alpha = "A" in im.getbands()
            out_mode = "RGBA" if (has_alpha or pad_transparent) else im.mode
            bg = (0, 0, 0, 0) if pad_transparent else (
                pad_color if out_mode == "RGBA" else pad_color[:3]
            )

            if any((extra_left, extra_top, extra_right, extra_bottom)):
                new_w = w + extra_left + extra_right
                new_h = h + extra_top  + extra_bottom
                canvas = Image.new(out_mode, (new_w, new_h), bg)
                canvas.paste(im.convert(out_mode), (extra_left, extra_top))
                result = canvas
            else:
                result = im.copy()

            if target_ratio is not None:
                result = pad_to_ratio(result, target_ratio, ratio_bg, ratio_transparent)

            if result.size == im.size and target_ratio is None:
                return f"skip margins-ok: {path.name}"

            out_path = path if overwrite else path.with_name(path.stem + "._padded" + path.suffix)
            result.save(out_path)

            delta = f"+{extra_left}L +{extra_top}T +{extra_right}R +{extra_bottom}B"
            ratio_note = f"  ratio→{result.size[0]}x{result.size[1]}" if target_ratio else ""
            return f"ok: {path.name}  ({delta}){ratio_note}"

    except Exception as e:
        return f"error: {path.name}: {e}"


def run(
    folder,
    *,
    recursive: bool = False,
    only_png: bool = True,
    overwrite: bool = True,
    min_left: int = 50,
    min_right: int = 50,
    min_top: int = 50,
    min_bottom: int = 50,
    white_tolerance: int = 10,
    alpha_threshold: int = 0,
    pad_color: tuple = (255, 255, 255, 255),
    pad_transparent: bool = False,
    target_ratio: float | None = None,
    ratio_bg: tuple = (255, 255, 255, 255),
    ratio_transparent: bool = False,
) -> None:
    """
    Add padding to PNGs where any margin is below the minimum.
    If target_ratio is set, images are also padded symmetrically to reach
    that aspect ratio (width / height) after margin padding.
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
        print(_pad_file(f, overwrite, min_left, min_top, min_right, min_bottom,
                        white_tolerance, alpha_threshold, pad_color, pad_transparent,
                        target_ratio, ratio_bg, ratio_transparent))


def preview_image(
    path,
    *,
    white_tolerance: int = 10,
    alpha_threshold: int = 0,
    min_left: int = 50,
    min_top: int = 50,
    min_right: int = 50,
    min_bottom: int = 50,
    pad_color: tuple = (255, 255, 255, 255),
    pad_transparent: bool = False,
    target_ratio: float | None = None,
    ratio_bg: tuple = (255, 255, 255, 255),
    ratio_transparent: bool = False,
) -> "Image.Image | None":
    """Return the padded image without saving, or None if skipped/error."""
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
        left, top, right, bottom = bbox
        extra_left   = max(0, min_left   - left)
        extra_top    = max(0, min_top    - top)
        extra_right  = max(0, min_right  - (w - right))
        extra_bottom = max(0, min_bottom - (h - bottom))
        has_alpha = "A" in im.getbands()
        out_mode = "RGBA" if (has_alpha or pad_transparent) else im.mode
        bg = (0, 0, 0, 0) if pad_transparent else (pad_color if out_mode == "RGBA" else pad_color[:3])
        if any((extra_left, extra_top, extra_right, extra_bottom)):
            new_w = w + extra_left + extra_right
            new_h = h + extra_top  + extra_bottom
            canvas = Image.new(out_mode, (new_w, new_h), bg)
            canvas.paste(im.convert(out_mode), (extra_left, extra_top))
            result = canvas
        else:
            result = im.copy()
        if target_ratio is not None:
            result = pad_to_ratio(result, target_ratio, ratio_bg, ratio_transparent)
        return result
    except Exception:
        return None
