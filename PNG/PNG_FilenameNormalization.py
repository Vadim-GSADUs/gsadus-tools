"""Standalone runner for filename normalization. Edit defaults below, or use app.py for a GUI."""
# ====== CONFIG: edit these ======
DEFAULT_FOLDER    = r"G:\Shared drives\GSADUs Projects\Our Models\0 - CATALOG\Output\3D Plan"
DEFAULT_RECURSIVE = True
DEFAULT_DRY_RUN   = True
DEFAULT_ADD_FOLDER = True
DEFAULT_TRUNCATE  = True
# ====== END CONFIG ======

if __name__ == "__main__":
    import argparse
    from core.filename_normalization import run

    parser = argparse.ArgumentParser(description="Normalize PNG filenames.")
    parser.add_argument("--folder",    "-f", default=DEFAULT_FOLDER)
    parser.add_argument("--recursive", "-r", action="store_true", default=DEFAULT_RECURSIVE)
    parser.add_argument("--dry-run",   "-d", action="store_true", default=DEFAULT_DRY_RUN)
    parser.add_argument("--add-folder","-a", action="store_true", default=DEFAULT_ADD_FOLDER)
    parser.add_argument("--truncate",  "-t", action="store_true", default=DEFAULT_TRUNCATE)
    args = parser.parse_args()

    run(
        args.folder,
        recursive=args.recursive,
        dry_run=args.dry_run,
        add_folder=args.add_folder,
        truncate=args.truncate,
    )
