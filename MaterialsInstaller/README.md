# GSADUs Materials Installer

Aggregates Revit material textures into the shared GSADUs/GSDE materials folder so every drafter renders against the same library.

## Default Paths

- **Source:** `G:\Shared drives\GSDE Projects\CADD\RevitFamily.Biz` (the family library — change as needed)
- **Destination:** `G:\Shared drives\GSDE Projects\CADD\Materials` (canonical shared folder)

The destination lives on the GSDE drive because GSDE handles post-contract Revit drafting (see `Vault\wiki\curated\key-locations.md` -> "GSDE Drive"). Both paths can be changed in the GUI.

## How It Works

One source, one destination, one **Install Materials** button. The source can be:

- A folder of Revit families (e.g. `RevitFamily.Biz`).
- A folder of downloaded bundles (e.g. `Downloads/`).
- A single archive (`.zip` / `.rar` / `.7z`) picked via the **File...** button.

The tool walks the source tree and copies image files into the destination, deduplicating by filename. It does two things at every step:

1. **Folder rule (legacy):** any image whose immediate parent folder is named `Materials` or `Textures` (case-insensitive) is pulled.
2. **Archive rule (new):** any archive encountered is extracted to a temp folder. If the extracted contents contain **no** Revit family (`.rfa/.rvt/.rte/.rft`), every image inside is pulled — that archive is a textures bundle. If a family is present, the tool recurses with the folder rule above.

This handles common vendor patterns:

| Pattern | Example | Outcome |
|---------|---------|---------|
| `family/Materials/*.jpg` | Generic library | Folder rule pulls images |
| `.rfa` + `Materials.zip` (flat .jpgs) | BIMobject Bok chair | Archive rule pulls images |
| `.rfa` + `<MaterialName>.zip` (often RAR-as-zip) | BIMobject Roller Max | Archive rule pulls (needs 7-Zip) |
| `.rfa` + `.txt` only | BIMobject Alphabet Sofa | Clean no-op |

### 7-Zip Auto-Install

The tool extracts `.zip` natively. Some vendors ship RAR archives renamed `.zip` (e.g. `ETH_Oak_Natural.zip` is actually RAR), which needs 7-Zip. If the source contains any archives and 7-Zip isn't installed, the tool prompts to install it via:

```
winget install --id 7zip.7zip -e --silent
```

If you decline, only standard `.zip` archives will extract; non-zip archives are skipped with a warning.

## Per-PC Revit Setup (one-time, separate from this tool)

This tool only populates the shared folder. Each machine's Revit/Architextures still needs three pointers wired up — see `Vault\wiki\curated\architextures-material-sync.md`:

1. **Revit:** File -> Options -> Rendering -> Additional render appearance paths -> add the shared destination path above.
2. **Architextures:** Revit ribbon -> Add-Ins -> Architextures -> Relocate Textures Folder -> same path.
3. **Google Drive:** right-click the shared `Materials` folder -> Offline access -> Available offline.

Without all three, materials still render with pink "missing texture" placeholders even though the textures are present on the shared drive.

## Quick Start

- **Batch:** double-click `MaterialsInstaller.bat`.
- **PowerShell:** right-click `Install-Materials.ps1` -> Run with PowerShell.
- **Standalone exe:** run `Build-Installer.ps1`, then `dist\MaterialsInstaller.exe`.

Admin rights are no longer required (destination is no longer under Program Files).

## Files

| File | Description |
|------|-------------|
| `Install-Materials.ps1` | Main PowerShell script with GUI |
| `MaterialsInstaller.bat` | Batch launcher |
| `Build-Installer.ps1` | Builds standalone .exe via PS2EXE |
| `dist/` | Build output (rebuild after script changes) |

## Image Extensions Recognized

`.jpg .jpeg .png .bmp .tif .tiff .dds .tga`
