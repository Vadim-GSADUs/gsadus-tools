# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

A `ttkbootstrap` desktop GUI app (`app.py`) for batch PNG post-processing, backed by a `core/` package of independent workflow modules. Targets the GSADUs product catalog PNG pipeline on a Google Shared Drive (`G:\Shared drives\GSADUs Projects\Our Models\0 - CATALOG\`).

## Running

```bash
# Launch the GUI (double-click "PNG Tools.pyw" in Explorer, or:)
python app.py

# Run any workflow as a headless script (edit CONFIG block at top)
python PNG_WhitespaceCrop.py
python PNG_TransparentCrop.py
python PNG_FilenameNormalization.py   # supports --dry-run, --folder, etc.
python PNG_WhitespaceAudit.py

# Install dependencies
pip install -r requirements.txt
```

## Architecture

```
app.py                  Entry point: tb.Notebook with one tab per workflow
config.py               JSON load/save for last-used settings (→ config.json, gitignored)
core/                   Pure logic — no UI, no globals, callable from GUI or CLI
  _image_utils.py       Shared PIL helpers: is_uniform, content_bbox_rgba/rgb
  whitespace_crop.py    run(folder, *, recursive, overwrite, buffer_*, white_tolerance, alpha_threshold)
  transparent_crop.py   run(folder, *, use_border_connected, target_color_bgr, color_tolerance, ...)
  filename_normalization.py  run(folder, *, recursive, dry_run, add_folder, truncate)
  whitespace_audit.py   run(folder, *, recursive, required_*, white_tolerance, ...) → Path
ui/
  base_tab.py           BaseTab(tb.Frame): path pickers, log widget, threading, config persistence
  tab_*.py              One file per workflow — subclasses BaseTab, implements _get_run_kwargs + _run
PNG_*.py                Thin wrappers with editable CONFIG block at top; call into core/
```

## Adding a New Workflow

1. Add `core/my_tool.py` with a `run(folder, *, ...)` function that uses `print()` for output.
2. Add `ui/tab_my_tool.py` subclassing `BaseTab` — use `_add_section`, `_add_path_row`, `_add_checkbox`, `_add_int_row`, `_add_inline_int_group` to build controls, then call `_finish_build()` last.
3. Register it in `app.py`'s `TABS` list.

## Key Patterns

**Threading** — `BaseTab._on_run()` runs the workflow in a daemon thread with `sys.stdout` redirected to a `Queue`. A 50ms `after()` poll drains the queue and appends to the log widget. Core `run()` functions need zero changes — their `print()` calls stream to the UI automatically.

**Config** — all `tk.Variable` instances registered in `self._vars` are automatically persisted to `config.json` on each Run and restored on next launch. For non-Variable state (e.g. the colour tuple in `TransparentCropTab`), store a serialisable string in `_vars` and override `_on_config_loaded()`.

**Content detection** — both `whitespace_crop` and `whitespace_audit` prefer alpha channel when present, falling back to near-white grayscale detection. Shared helpers live in `core/_image_utils.py`.

**BGR vs RGB** — `transparent_crop` uses `cv2` (BGR). The UI presents colours as RGB and converts to BGR in `_get_run_kwargs()`.

## Dependencies

| Package | Used by |
|---|---|
| `ttkbootstrap` | entire UI (darkly theme, Python 3.13 compatible) |
| `Pillow` | `whitespace_crop`, `whitespace_audit`, `_image_utils` |
| `opencv-python` + `numpy` | `transparent_crop` only |
