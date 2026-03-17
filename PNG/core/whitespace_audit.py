"""Audit PNGs for minimum whitespace margins; write a CSV report of failures."""
import csv
from pathlib import Path

from PIL import Image

from core._image_utils import content_bbox_rgb, content_bbox_rgba, is_uniform


def _measure_margins(im: Image.Image, white_tolerance: int, alpha_threshold: int):
    """Return ((mL, mT, mR, mB), bbox, reason). margins is None on failure."""
    w, h = im.size
    if is_uniform(im):
        return None, None, "uniform_image"
    bbox = content_bbox_rgba(im, alpha_threshold) if "A" in im.getbands() else None
    if not bbox:
        bbox = content_bbox_rgb(im, white_tolerance)
    if not bbox:
        return None, None, "no_content_bbox"
    left, top, right, bottom = bbox
    return (left, top, w - right, h - bottom), bbox, "ok"


def _audit_file(
    p: Path,
    white_tolerance: int,
    alpha_threshold: int,
    req_left: int,
    req_top: int,
    req_right: int,
    req_bottom: int,
) -> dict:
    try:
        with Image.open(p) as im:
            margins, _, reason = _measure_margins(im, white_tolerance, alpha_threshold)
            base = {
                "file": p.name,
                "width": im.size[0],
                "height": im.size[1],
                "req_left": req_left, "req_top": req_top,
                "req_right": req_right, "req_bottom": req_bottom,
            }
            if margins is None:
                return {**base, "m_left": "", "m_top": "", "m_right": "", "m_bottom": "",
                        "status": "FAIL", "reason": reason}
            mL, mT, mR, mB = margins
            ok = mL >= req_left and mT >= req_top and mR >= req_right and mB >= req_bottom
            return {**base, "m_left": mL, "m_top": mT, "m_right": mR, "m_bottom": mB,
                    "status": "PASS" if ok else "FAIL",
                    "reason": "ok" if ok else "insufficient_margin"}
    except Exception as e:
        return {"file": p.name, "width": "", "height": "",
                "m_left": "", "m_top": "", "m_right": "", "m_bottom": "",
                "req_left": req_left, "req_top": req_top,
                "req_right": req_right, "req_bottom": req_bottom,
                "status": "FAIL", "reason": f"error:{e}"}


def run(
    folder,
    *,
    recursive: bool = False,
    only_png: bool = True,
    required_left: int = 50,
    required_right: int = 50,
    required_top: int = 50,
    required_bottom: int = 25,
    white_tolerance: int = 10,
    alpha_threshold: int = 0,
    report_name: str = "_PNG_WhitespaceAudit.csv",
    target_ratio: float | None = None,
    ratio_tolerance: float = 0.01,
) -> Path | None:
    """
    Measure content margins for all PNGs in folder.
    If target_ratio is set, also checks whether each image matches that aspect ratio
    (width / height) within ratio_tolerance.
    Writes a CSV of failures to folder/report_name.
    Returns the report Path, or None if no files were found.
    """
    root = Path(folder)
    if not root.exists():
        print(f"Folder not found: {root}")
        return None

    patterns = ["*.png"] if only_png else ["*.png", "*.PNG"]
    files: list[Path] = []
    if recursive:
        for ptn in patterns:
            files.extend(root.rglob(ptn))
    else:
        for ptn in patterns:
            files.extend(root.glob(ptn))
    files = [p for p in files if p.is_file()]

    if not files:
        print("No files found.")
        return None

    results = [
        _audit_file(p, white_tolerance, alpha_threshold,
                    required_left, required_top, required_right, required_bottom)
        for p in files
    ]

    # Apply aspect ratio check on top of margin results
    if target_ratio is not None:
        for r in results:
            if r["width"] and r["height"]:
                try:
                    actual = int(r["width"]) / int(r["height"])
                    if abs(actual - target_ratio) > ratio_tolerance:
                        r["status"] = "FAIL"
                        r["reason"] = (r["reason"] + "; " if r["reason"] not in ("ok", "") else "") + \
                                      f"aspect_ratio {actual:.3f} != {target_ratio:.3f}"
                except (ValueError, ZeroDivisionError):
                    pass

    fails = [r for r in results if r["status"] == "FAIL"]

    report_path = root / report_name
    fieldnames = ["file", "width", "height",
                  "m_left", "m_top", "m_right", "m_bottom",
                  "req_left", "req_top", "req_right", "req_bottom",
                  "status", "reason"]
    with open(report_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for r in fails:
            writer.writerow(r)

    total = len(results)
    n_fail = len(fails)
    ratio_note = f"  |  Target ratio: {target_ratio:.3f}" if target_ratio else ""
    print(f"Audited: {total}  |  Pass: {total - n_fail}  |  Fail: {n_fail}{ratio_note}")
    print(f"Report:  {report_path}")
    return report_path
