# GSADUs Family Catalog Generator

A tool to scan Revit family folders and generate CSV catalogs for inventory management and bulk operations.

## Purpose

When managing large Revit family libraries (3000+ files), it's essential to have visibility into what's available. This tool:

1. Scans specified category folders for `.rfa` files
2. Generates detailed CSV catalogs with file metadata
3. Creates both per-category and master catalogs
4. Enables bulk management and planning for family replacements

## CSV Output Fields

Each generated catalog includes:

| Field | Description |
|-------|-------------|
| FileName | Full file name with extension |
| FamilyName | File name without extension |
| RelativePath | Path relative to category folder |
| FullPath | Complete file path |
| Category | Category folder name (Windows, Doors, etc.) |
| FileSizeBytes | File size in bytes (for sorting/filtering) |
| FileSize | Human-readable file size (KB/MB) |
| DateModified | Last modification date |
| DateCreated | File creation date |
| ParentFolder | Immediate parent folder name |

## Quick Start

### Option 1: Run the Batch File (Simplest)
Double-click `FamilyCatalogGenerator.bat` to launch the tool.

### Option 2: Run the PowerShell Script
Right-click `Generate-FamilyCatalog.ps1` and select "Run with PowerShell"

### Option 3: Use the Standalone Executable
1. Run `Build-CatalogGenerator.ps1` to create the executable
2. Find `FamilyCatalogGenerator.exe` in the `dist` folder
3. Copy the exe to any Windows machine and run it

## Building the Executable

To create a standalone `.exe` file:

1. Open PowerShell
2. Navigate to this folder
3. Run: `.\Build-CatalogGenerator.ps1`

The executable will be created in the `dist` folder.

## Default Category Paths

The tool comes pre-configured with these category folders:

- `G:\Shared drives\GSDE Projects\CADD\RevitFamily.Biz\Windows`
- `G:\Shared drives\GSDE Projects\CADD\RevitFamily.Biz\Cabinets`
- `G:\Shared drives\GSDE Projects\CADD\RevitFamily.Biz\Doors`
- `G:\Shared drives\GSDE Projects\CADD\RevitFamily.Biz\Electrical`
- `G:\Shared drives\GSDE Projects\CADD\RevitFamily.Biz\Finish Carpentry`
- `G:\Shared drives\GSDE Projects\CADD\RevitFamily.Biz\Railings`

You can add, remove, or modify these paths in the GUI.

## Output Options

- **Save CSVs in category folders**: Places a `_FamilyCatalog_[Category].csv` file in each scanned category folder
- **Generate Master Catalog**: Creates a combined `FamilyCatalog_MASTER.csv` with all families from all categories

## Use Cases

### Family Replacement Planning
1. Generate catalogs for all categories
2. Open the master CSV in Excel
3. Filter/sort to identify families by name patterns
4. Plan which families will replace existing project families

### Inventory Management
- Track total family count per category
- Identify duplicate family names across categories
- Monitor file sizes for optimization

### Future Expansion
The CSV format enables future automation:
- Bulk parameter updates via Revit API
- Family name standardization
- Automated family loading scripts

## Requirements

- Windows 10 or later
- PowerShell 5.1 or later (included with Windows 10)
- Access to the source folders (Google Shared Drive or local)

## Files

| File | Description |
|------|-------------|
| `Generate-FamilyCatalog.ps1` | Main PowerShell script with GUI |
| `FamilyCatalogGenerator.bat` | Batch file launcher |
| `Build-CatalogGenerator.ps1` | Script to create standalone .exe |
| `dist/` | Output folder for built executable |
