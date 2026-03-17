"""Batch rename PNGs: strip suffix, find/replace, case, prefix/suffix, numbering, and more."""
import os
import re


# ── Per-stem transforms (applied before truncation) ───────────────────────────

def _apply_per_stem(
    stem: str,
    *,
    strip_after_delimiter: bool,
    delimiter: str,
    strip_strings: list,
    find_replace: bool,
    find_str: str,
    replace_str: str,
    find_case_sensitive: bool,
    strip_digits: bool,
    strip_special: bool,
    case_mode: str,
) -> str:
    if strip_after_delimiter and delimiter and delimiter in stem:
        stem = stem.split(delimiter, 1)[0]

    for s in strip_strings:
        if s:
            stem = stem.replace(s, "")

    if find_replace and find_str:
        flags = 0 if find_case_sensitive else re.IGNORECASE
        stem = re.sub(re.escape(find_str), replace_str, stem, flags=flags)

    if strip_digits:
        stem = re.sub(r"^\d+[\s\-_]*", "", stem)
        stem = re.sub(r"[\s\-_]*\d+$", "", stem)

    if strip_special:
        stem = re.sub(r"[^\w\s\-]", "", stem)  # keep alphanumeric, underscore, space, hyphen

    stem = re.sub(r"\s+", " ", stem).strip()

    if case_mode == "title":
        stem = stem.title()
    elif case_mode == "lower":
        stem = stem.lower()
    elif case_mode == "upper":
        stem = stem.upper()

    return stem


# ── Post-truncate transforms (applied after uniqueness pass) ───────────────────

def _apply_post_truncate(
    stem: str,
    *,
    add_folder: bool,
    folder_name: str,
    prefix: str,
    suffix: str,
    replace_spaces: str,
) -> str:
    if add_folder and folder_name:
        token = f" {folder_name}"
        idx = stem.lower().rfind(token.lower())
        stem = stem[:idx] + token if idx != -1 else f"{stem}{token}"

    stem = f"{prefix}{stem}{suffix}"
    stem = re.sub(r"\s+", " ", stem).strip()

    if replace_spaces == "_":
        stem = stem.replace(" ", "_")
    elif replace_spaces == "-":
        stem = stem.replace(" ", "-")
    elif replace_spaces == "remove":
        stem = stem.replace(" ", "")

    return stem


# ── Core shared logic ──────────────────────────────────────────────────────────

def _compute_renames(
    folder_path: str,
    *,
    recursive: bool,
    add_folder: bool,
    truncate: bool,
    strip_after_delimiter: bool,
    delimiter: str,
    strip_strings: list,
    find_replace: bool,
    find_str: str,
    replace_str: str,
    find_case_sensitive: bool,
    prefix: str,
    suffix: str,
    case_mode: str,
    replace_spaces: str,
    strip_special: bool,
    strip_digits: bool,
    sequential: bool,
    seq_start: int,
    seq_step: int,
    seq_padding: int,
    seq_position: str,
) -> list[tuple[str, str, str]]:
    """Return list of (root_dir, old_name, new_name). Does not touch the filesystem."""

    def iter_files():
        if recursive:
            for root, _, files in os.walk(folder_path):
                for f in files:
                    yield root, f
        else:
            for f in os.listdir(folder_path):
                yield folder_path, f

    files_by_root: dict[str, list[str]] = {}
    for root, filename in iter_files():
        if filename.lower().endswith(".png"):
            files_by_root.setdefault(root, []).append(filename)

    # Phase 1: per-stem transforms
    all_entries: list[dict] = []
    for root in sorted(files_by_root):
        for filename in sorted(files_by_root[root]):
            stem, ext = os.path.splitext(filename)
            stem = _apply_per_stem(
                stem,
                strip_after_delimiter=strip_after_delimiter,
                delimiter=delimiter,
                strip_strings=strip_strings,
                find_replace=find_replace,
                find_str=find_str,
                replace_str=replace_str,
                find_case_sensitive=find_case_sensitive,
                strip_digits=strip_digits,
                strip_special=strip_special,
                case_mode=case_mode,
            )
            all_entries.append({"root": root, "orig": filename, "ext": ext, "words": stem.split()})

    # Phase 2: truncate within each folder group for uniqueness
    if truncate:
        for root in sorted(files_by_root):
            group = [e for e in all_entries if e["root"] == root]
            for idx, e in enumerate(group):
                words = e["words"]
                for k in range(1, len(words) + 1):
                    pfx = words[:k]
                    collision = any(
                        j != idx
                        and len(other["words"]) >= k
                        and other["words"][:k] == pfx
                        for j, other in enumerate(group)
                    )
                    if not collision:
                        e["words"] = pfx
                        break

    # Phase 3: post-truncate transforms + sequential numbering
    sep = {"keep": " ", "_": "_", "-": "-", "remove": ""}.get(replace_spaces, " ")
    seq_counter = seq_start
    triples: list[tuple[str, str, str]] = []

    for e in all_entries:
        folder_name = os.path.basename(e["root"])
        stem_part = " ".join(e["words"])
        stem_part = _apply_post_truncate(
            stem_part,
            add_folder=add_folder,
            folder_name=folder_name,
            prefix=prefix,
            suffix=suffix,
            replace_spaces=replace_spaces,
        )
        if sequential:
            num_str = str(seq_counter).zfill(seq_padding)
            if seq_position == "prefix":
                stem_part = f"{num_str}{sep}{stem_part}"
            else:
                stem_part = f"{stem_part}{sep}{num_str}"
            seq_counter += seq_step

        new_name = f"{stem_part}{e['ext']}"
        if new_name != e["orig"]:
            triples.append((e["root"], e["orig"], new_name))

    return triples


# ── Public API ─────────────────────────────────────────────────────────────────

_DEFAULTS = dict(
    recursive=True,
    add_folder=True,
    truncate=True,
    strip_after_delimiter=True,
    delimiter=" - ",
    strip_strings=[],
    find_replace=False,
    find_str="",
    replace_str="",
    find_case_sensitive=False,
    prefix="",
    suffix="",
    case_mode="none",
    replace_spaces="keep",
    strip_special=False,
    strip_digits=False,
    sequential=False,
    seq_start=1,
    seq_step=1,
    seq_padding=3,
    seq_position="prefix",
)


def get_renames(folder, **kwargs) -> list[tuple[str, str]]:
    """Return (old_name, new_name) pairs without touching the filesystem."""
    folder_path = os.path.normpath(str(folder))
    if not os.path.isdir(folder_path):
        return []
    opts = {**_DEFAULTS, **kwargs}
    triples = _compute_renames(folder_path, **opts)
    return [(old, new) for _, old, new in triples]


def run(folder, *, dry_run: bool = True, **kwargs) -> None:
    """
    Rename PNG files in folder according to the configured options.
    dry_run=True (default) prints proposed renames without touching the filesystem.
    """
    folder_path = os.path.normpath(str(folder))
    if not os.path.isdir(folder_path):
        print("Invalid folder path.")
        return

    opts = {**_DEFAULTS, **kwargs}
    triples = _compute_renames(folder_path, **opts)

    if not triples:
        print("No renames needed.")
        return

    changed = 0
    for root, old_name, new_name in triples:
        old_path = os.path.join(root, old_name)
        new_path = os.path.join(root, new_name)

        if dry_run:
            print(f"[DRY RUN]  {old_name}  →  {new_name}")
        else:
            # Avoid collisions with existing files
            if os.path.exists(new_path) and new_path.lower() != old_path.lower():
                base, ext = os.path.splitext(new_path)
                n = 1
                while os.path.exists(new_path):
                    new_path = f"{base} ({n}){ext}"
                    n += 1
                new_name = os.path.basename(new_path)
            os.rename(old_path, new_path)
            print(f"Renamed:   {old_name}  →  {new_name}")
            changed += 1

    if not dry_run:
        print(f"Done. {changed} file(s) renamed.")
