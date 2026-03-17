"""Shared PIL image helpers used by whitespace_crop and whitespace_audit."""
from PIL import Image, ImageStat


def is_uniform(im: Image.Image) -> bool:
    """True if the image is a single solid colour (no content to crop)."""
    stat = ImageStat.Stat(im.convert("RGBA"))
    return all(v == 0 for v in stat.var)


def content_bbox_rgba(im: Image.Image, alpha_threshold: int):
    """Bounding box of pixels whose alpha > alpha_threshold. Returns None if empty."""
    if im.mode != "RGBA":
        im = im.convert("RGBA")
    alpha = im.split()[-1]
    content = alpha.point(lambda a: 255 if a > alpha_threshold else 0, mode="L")
    return content.getbbox()


def content_bbox_rgb(im: Image.Image, white_tolerance: int):
    """Bounding box of pixels that are not near-white. Returns None if empty."""
    if im.mode not in ("RGB", "L"):
        im = im.convert("RGB")
    gray = im.convert("L")
    cutoff = 255 - int(white_tolerance)
    content = gray.point(lambda p: 255 if p < cutoff else 0, mode="L")
    return content.getbbox()


def pad_to_ratio(
    im: Image.Image,
    target_ratio: float,
    bg_rgba: tuple = (255, 255, 255, 255),
    transparent: bool = False,
) -> Image.Image:
    """
    Pad image symmetrically to reach target aspect ratio (width / height).
    Never crops — only adds whitespace. Returns original if already at ratio.
    """
    w, h = im.size
    if abs(w / h - target_ratio) < 0.002:
        return im

    if w / h > target_ratio:
        new_w, new_h = w, max(h + 1, round(w / target_ratio))
    else:
        new_w, new_h = max(w + 1, round(h * target_ratio)), h

    has_alpha = "A" in im.getbands()
    out_mode = "RGBA" if (has_alpha or transparent) else im.mode
    bg = (0, 0, 0, 0) if transparent else (bg_rgba if out_mode == "RGBA" else bg_rgba[:3])

    canvas = Image.new(out_mode, (new_w, new_h), bg)
    canvas.paste(im.convert(out_mode), ((new_w - w) // 2, (new_h - h) // 2))
    return canvas
