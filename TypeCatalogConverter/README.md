# GSADUs TypeCatalog Converter

A bidirectional conversion tool for Revit family type catalogs. Convert between Revit's native `.txt` format and Excel `.xlsx` for batch editing multiple families at once.

## Purpose

When managing large Revit family libraries, editing type parameters one family at a time is tedious. This tool enables:

1. **Batch Import**: Merge multiple family type catalogs into a single Excel spreadsheet
2. **Bulk Editing**: Edit types across multiple families in Excel (with formulas, find/replace, etc.)
3. **Batch Export**: Convert the edited spreadsheet back to individual Revit-compatible `.txt` files

## Workflow

### Import Mode (TXT → XLSX)

```
1. Select folder(s) containing .rfa + .txt files
2. (Optional) Click "Scan Columns..." to configure column order/visibility
3. Click "Generate XLSX Files"
4. Tool scans for .txt files with matching .rfa files
5. Merges all type catalogs into one XLSX per folder
6. Output: FolderName_TypeCatalog.xlsx
7. (Optional) XLSX auto-opens if checkbox is checked
```

### Column Configuration

The column configuration feature helps manage spreadsheets with many parameters (70+ columns):

1. **Scan Columns**: Click "Scan Columns..." after adding folders to discover all available parameters
2. **Configure Dialog**: Opens a dialog where you can:
   - ☑ Check/uncheck columns to show/hide them in the XLSX
   - Drag or use buttons to reorder columns (most-used at top)
   - Use "Check Common" to auto-select frequently used parameters
3. **Persistence**: Configuration is saved in the `_Settings` sheet of the XLSX
4. **Reload**: When regenerating, existing configuration is loaded from the XLSX

### Export Mode (XLSX → TXT)

```
1. Switch to "Export XLSX → TXT" mode
2. Select the XLSX file(s) to convert
3. Click "Export to TXT Files"
4. Tool recreates individual .txt files for each family
5. Output: Original locations or custom folder
```

## XLSX Structure

### Sheet 1: "TypeCatalog" (Main Data)

| Family | Type Name | Height | Width | Frame Depth | Glazing Mtrl | ... |
|--------|-----------|--------|-------|-------------|--------------|-----|
| Window - Fixed | 3'-0" x 5'-0" | 5.0 | 3.0 | 0.333 | Glass | ... |
| Window - Fixed | 4'-0" x 6'-0" | 6.0 | 4.0 | 0.333 | Glass | ... |
| Window - Awning | 2'-6" x 3'-0" | 3.0 | 2.5 | 0.291 | Glass | ... |

- **Column A (Family)**: Family name - matches .txt filename
- **Column B (Type Name)**: The type name that appears in Revit
- **Remaining columns**: Parameters with clean readable names
- **Empty cells**: Parameter doesn't apply to that family

### Sheet 2: "_Metadata" (Required for Export)

Contains the original Revit header format for each parameter per family. **Do not delete this sheet** - it's required for converting back to TXT.

| Family | ParameterName | FullHeader | ColumnIndex | DataType | Unit |
|--------|---------------|------------|-------------|----------|------|
| Window - Fixed | Height | Height[-1001300]##LENGTH##FEET | 49 | LENGTH | FEET |

### Sheet 3: "_Settings" (Configuration Persistence)

Stores configuration including unit conversion settings and column preferences.

| Setting | Value |
|---------|-------|
| UnitConversion | ConvertToInches |
| GeneratedDate | 2026-01-30 14:30:00 |
| SourceFolder | C:\Families\Windows |
| ColumnConfig | [{"Name":"Family","Visible":true,"Order":0},...] |

The `ColumnConfig` setting preserves your column order and visibility preferences for future regeneration.

## Editing Tips

### Safe Operations
- **Edit values**: Change any parameter value in the TypeCatalog sheet
- **Delete rows**: Remove types you don't want (they won't be exported)
- **Add rows**: Add new types by copying an existing row and modifying values
- **Use formulas**: Excel formulas work - values are exported, not formulas

### What to Avoid
- **Don't rename Column A or B**: "Family" and "Type Name" must stay as-is
- **Don't add new columns**: They'll be ignored (no metadata for them)
- **Don't delete _Metadata sheet**: Required for export
- **Don't change Family names**: Must match original .txt filenames

### Adding New Types
1. Copy an existing row for that family
2. Paste it as a new row
3. Change the "Type Name" to your new type name
4. Modify parameter values as needed
5. Export - the new type will be included in that family's .txt

## Quick Start

### Option 1: Run the Batch File
Double-click `TypeCatalogConverter.bat`

### Option 2: Run the PowerShell Script
Right-click `Convert-TypeCatalog.ps1` and select "Run with PowerShell"

### Option 3: Use the Standalone Executable
1. Run `Build-TypeCatalogConverter.ps1` to create the executable
2. Find `TypeCatalogConverter.exe` in the `dist` folder
3. Copy the exe anywhere and run it

## Requirements

- Windows 10 or later
- PowerShell 5.1 or later
- **ImportExcel module** (auto-installs on first run if missing)

## Validation Rules

| Scenario | Behavior |
|----------|----------|
| .txt without matching .rfa | Skipped - not processed |
| Family in XLSX not in metadata | Error - cannot export |
| New columns added in Excel | Ignored on export |
| Row deleted in Excel | Type not exported |
| New row with existing Family | Exported as new type |
| Empty cells | Omitted from that family's .txt |

## File Format Notes

- **Encoding**: Revit uses UTF-16 (Unicode) for type catalogs
- **Delimiter**: Comma-separated values
- **Header format**: `ParameterName[-BuiltInId]##DATATYPE##UNIT`
- **Values**: Lengths in decimal feet, booleans as 0/1

## Files

| File | Description |
|------|-------------|
| `Convert-TypeCatalog.ps1` | Main PowerShell script with GUI |
| `TypeCatalogConverter.bat` | Batch file launcher |
| `Build-TypeCatalogConverter.ps1` | Script to create standalone .exe |
| `dist/` | Output folder for built executable |

## Troubleshooting

### "ImportExcel module not found"
The tool will attempt to auto-install. If it fails:
```powershell
Install-Module -Name ImportExcel -Scope CurrentUser -Force
```

### "No valid type catalogs found"
- Ensure .txt files have matching .rfa files (same name)
- Check that files are actual Revit type catalogs (not other .txt files)

### Export creates empty/corrupted files
- Verify the _Metadata sheet exists and has data
- Check that Family names in TypeCatalog match those in _Metadata
