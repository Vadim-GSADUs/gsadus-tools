#Requires -Version 5.1
<#
.SYNOPSIS
    TypeCatalog Converter - Converts Revit type catalogs between TXT and XLSX formats.

.DESCRIPTION
    This tool provides bidirectional conversion between Revit family type catalogs (.txt)
    and Excel spreadsheets (.xlsx) for batch editing of family types across multiple families.

.NOTES
    Author: GSADUs
    Created: 2026-01-30
    Updated: 2026-01-30 - Added modern folder dialog, unit conversion
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Add shell COM interop for modern folder dialog
$modernFolderDialogSource = @"
using System;
using System.Runtime.InteropServices;

[ComImport]
[Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
public class FileOpenDialogClass
{
}

[ComImport]
[Guid("42f85136-db7e-439c-85f1-e4075d135fc8")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IFileOpenDialog
{
    [PreserveSig] int Show([In] IntPtr parent);
    void SetFileTypes();
    void SetFileTypeIndex([In] uint iFileType);
    void GetFileTypeIndex(out uint piFileType);
    void Advise();
    void Unadvise();
    void SetOptions([In] uint fos);
    void GetOptions(out uint pfos);
    void SetDefaultFolder(IShellItem psi);
    void SetFolder(IShellItem psi);
    void GetFolder(out IShellItem ppsi);
    void GetCurrentSelection(out IShellItem ppsi);
    void SetFileName([In, MarshalAs(UnmanagedType.LPWStr)] string pszName);
    void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
    void SetTitle([In, MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
    void SetOkButtonLabel([In, MarshalAs(UnmanagedType.LPWStr)] string pszText);
    void SetFileNameLabel([In, MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
    void GetResult(out IShellItem ppsi);
    void AddPlace(IShellItem psi, int alignment);
    void SetDefaultExtension([In, MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
    void Close(int hr);
    void SetClientGuid();
    void ClearClientData();
    void SetFilter([MarshalAs(UnmanagedType.Interface)] IntPtr pFilter);
    void GetResults([MarshalAs(UnmanagedType.Interface)] out IntPtr ppenum);
    void GetSelectedItems([MarshalAs(UnmanagedType.Interface)] out IntPtr ppsai);
}

[ComImport]
[Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IShellItem
{
    void BindToHandler();
    void GetParent();
    void GetDisplayName([In] uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
    void GetAttributes();
    void Compare();
}

public class ModernFolderBrowser
{
    private const uint FOS_PICKFOLDERS = 0x00000020;
    private const uint FOS_FORCEFILESYSTEM = 0x00000040;
    private const uint FOS_NOVALIDATE = 0x00000100;
    private const uint SIGDN_FILESYSPATH = 0x80058000;

    public static string ShowDialog(IntPtr owner, string title)
    {
        IFileOpenDialog dialog = (IFileOpenDialog)new FileOpenDialogClass();
        try
        {
            dialog.SetOptions(FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_NOVALIDATE);
            dialog.SetTitle(title);

            int hr = dialog.Show(owner);
            if (hr != 0) return null;

            IShellItem item;
            dialog.GetResult(out item);
            string path;
            item.GetDisplayName(SIGDN_FILESYSPATH, out path);
            return path;
        }
        finally
        {
            Marshal.ReleaseComObject(dialog);
        }
    }
}
"@

try {
    Add-Type -TypeDefinition $modernFolderDialogSource -ErrorAction SilentlyContinue
}
catch {
    # Type already added, ignore
}

function Show-ModernFolderDialog {
    param(
        [string]$Title = "Select Folder",
        [System.Windows.Forms.Form]$Owner = $null
    )

    try {
        $handle = if ($Owner) { $Owner.Handle } else { [IntPtr]::Zero }
        $result = [ModernFolderBrowser]::ShowDialog($handle, $Title)
        return $result
    }
    catch {
        # Fallback to legacy dialog if modern fails
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = $Title
        if ($folderBrowser.ShowDialog() -eq "OK") {
            return $folderBrowser.SelectedPath
        }
        return $null
    }
}

# Check and import ImportExcel module
function Ensure-ImportExcelModule {
    $module = Get-Module -ListAvailable -Name ImportExcel
    if (-not $module) {
        Write-Log "ImportExcel module not found. Installing..." -Level "WARN"
        try {
            Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
            Write-Log "ImportExcel module installed successfully."
            Import-Module ImportExcel -Force
        }
        catch {
            Write-Log "Failed to install ImportExcel module: $_" -Level "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to install required module 'ImportExcel'.`n`nPlease run PowerShell as Administrator and execute:`nInstall-Module -Name ImportExcel -Scope CurrentUser -Force",
                "Module Required",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return $false
        }
    }
    else {
        Import-Module ImportExcel -Force
    }
    return $true
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    if ($script:LogTextBox) {
        $script:LogTextBox.AppendText("$logEntry`r`n")
        $script:LogTextBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Parse-TypeCatalogHeader {
    param([string]$HeaderLine)

    $columns = $HeaderLine -split ','
    $parameters = @()

    for ($i = 0; $i -lt $columns.Count; $i++) {
        $col = $columns[$i].Trim()

        if ($i -eq 0) {
            # First column is Type Name (no header in Revit format)
            $parameters += [PSCustomObject]@{
                Index = $i
                CleanName = "Type Name"
                FullHeader = ""
                ParameterName = ""
                DataType = ""
                Unit = ""
                BuiltInId = ""
            }
            continue
        }

        # Parse format: ParameterName[-BuiltInId]##DATATYPE##UNIT or ParameterName##DATATYPE##UNIT
        $fullHeader = $col
        $paramName = $col
        $dataType = ""
        $unit = ""
        $builtInId = ""

        # Extract ##DATATYPE##UNIT
        if ($col -match '^(.+?)##([^#]+)##(.*)$') {
            $paramName = $matches[1]
            $dataType = $matches[2]
            $unit = $matches[3]
        }
        elseif ($col -match '^(.+?)##([^#]+)$') {
            $paramName = $matches[1]
            $dataType = $matches[2]
        }

        # Extract [-BuiltInId] from parameter name
        if ($paramName -match '^(.+?)\[(-?\d+)\]$') {
            $cleanName = $matches[1]
            $builtInId = $matches[2]
        }
        else {
            $cleanName = $paramName
        }

        $parameters += [PSCustomObject]@{
            Index = $i
            CleanName = $cleanName.Trim()
            FullHeader = $fullHeader
            ParameterName = $paramName.Trim()
            DataType = $dataType
            Unit = $unit
            BuiltInId = $builtInId
        }
    }

    return $parameters
}

function Parse-TypeCatalogData {
    param(
        [string]$DataLine,
        [array]$Parameters
    )

    # Handle CSV parsing with potential quoted values
    $values = @()
    $currentValue = ""
    $inQuotes = $false

    for ($i = 0; $i -lt $DataLine.Length; $i++) {
        $char = $DataLine[$i]

        if ($char -eq '"') {
            if ($inQuotes -and ($i + 1 -lt $DataLine.Length) -and $DataLine[$i + 1] -eq '"') {
                # Escaped quote
                $currentValue += '"'
                $i++
            }
            else {
                $inQuotes = -not $inQuotes
            }
        }
        elseif ($char -eq ',' -and -not $inQuotes) {
            $values += $currentValue
            $currentValue = ""
        }
        else {
            $currentValue += $char
        }
    }
    $values += $currentValue

    $typeData = [ordered]@{}

    for ($i = 0; $i -lt $Parameters.Count; $i++) {
        $param = $Parameters[$i]
        $value = if ($i -lt $values.Count) { $values[$i].Trim() } else { "" }
        $typeData[$param.CleanName] = $value
    }

    return $typeData
}

function Read-TypeCatalog {
    param([string]$FilePath)

    $familyName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)

    try {
        # Read with UTF-16 encoding (Revit's native format)
        $content = Get-Content -Path $FilePath -Encoding Unicode -ErrorAction Stop

        if ($content.Count -lt 2) {
            Write-Log "  Skipping $familyName - insufficient data" -Level "WARN"
            return $null
        }

        $headerLine = $content[0]
        $parameters = Parse-TypeCatalogHeader -HeaderLine $headerLine

        $types = @()
        for ($i = 1; $i -lt $content.Count; $i++) {
            $line = $content[$i].Trim()
            if ([string]::IsNullOrWhiteSpace($line)) { continue }

            $typeData = Parse-TypeCatalogData -DataLine $line -Parameters $parameters
            $typeData["Family"] = $familyName
            $types += $typeData
        }

        return @{
            FamilyName = $familyName
            Parameters = $parameters
            Types = $types
            FilePath = $FilePath
        }
    }
    catch {
        Write-Log "  Error reading $familyName`: $_" -Level "ERROR"
        return $null
    }
}

function Convert-FeetToInches {
    param(
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return $Value }

    try {
        $numValue = [double]$Value
        $inchValue = $numValue * 12
        return $inchValue.ToString()
    }
    catch {
        # Not a number, return as-is
        return $Value
    }
}

function Convert-InchesToFeet {
    param(
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return $Value }

    try {
        $numValue = [double]$Value
        $feetValue = $numValue / 12
        return $feetValue.ToString()
    }
    catch {
        # Not a number, return as-is
        return $Value
    }
}

function Import-TypeCatalogsToXlsx {
    param(
        [string]$FolderPath,
        [string]$OutputPath,
        [string]$UnitConversion = "KeepFeet"  # "KeepFeet", "ConvertToInches"
    )

    $folderName = Split-Path -Path $FolderPath -Leaf
    Write-Log "Scanning folder: $folderName"

    # Find all .txt files that have matching .rfa files
    $txtFiles = Get-ChildItem -Path $FolderPath -Filter "*.txt" -File -Recurse -ErrorAction SilentlyContinue

    $validCatalogs = @()
    $skippedCount = 0

    foreach ($txtFile in $txtFiles) {
        $rfaPath = [System.IO.Path]::ChangeExtension($txtFile.FullName, ".rfa")

        if (Test-Path -Path $rfaPath) {
            $validCatalogs += $txtFile
        }
        else {
            $skippedCount++
        }
    }

    if ($skippedCount -gt 0) {
        Write-Log "  Skipped $skippedCount .txt files without matching .rfa"
    }

    if ($validCatalogs.Count -eq 0) {
        Write-Log "  No valid type catalogs found in $folderName" -Level "WARN"
        return $null
    }

    Write-Log "  Found $($validCatalogs.Count) valid type catalogs"

    # Parse all catalogs
    $allCatalogs = @()
    $allParameters = [ordered]@{}
    $allParameters["Family"] = $true
    $allParameters["Type Name"] = $true

    foreach ($txtFile in $validCatalogs) {
        $catalog = Read-TypeCatalog -FilePath $txtFile.FullName

        if ($catalog) {
            $allCatalogs += $catalog
            Write-Log "  Parsed: $($catalog.FamilyName) ($($catalog.Types.Count) types)"

            # Collect all unique parameter names
            foreach ($param in $catalog.Parameters) {
                if ($param.CleanName -ne "Type Name" -and -not $allParameters.Contains($param.CleanName)) {
                    $allParameters[$param.CleanName] = $true
                }
            }
        }
    }

    if ($allCatalogs.Count -eq 0) {
        Write-Log "  No catalogs could be parsed" -Level "ERROR"
        return $null
    }

    # Build a lookup for parameter metadata (to check if LENGTH + FEET)
    $paramMetaLookup = @{}
    foreach ($catalog in $allCatalogs) {
        foreach ($param in $catalog.Parameters) {
            $key = "$($catalog.FamilyName)|$($param.CleanName)"
            $paramMetaLookup[$key] = $param
        }
    }

    # Build merged data with optional unit conversion
    $mergedData = @()
    $columnOrder = @($allParameters.Keys)
    $convertedCount = 0

    foreach ($catalog in $allCatalogs) {
        foreach ($typeData in $catalog.Types) {
            $row = [ordered]@{}

            foreach ($colName in $columnOrder) {
                if ($typeData.Contains($colName)) {
                    $value = $typeData[$colName]

                    # Check if we need to convert this value
                    if ($UnitConversion -eq "ConvertToInches" -and $colName -ne "Family" -and $colName -ne "Type Name") {
                        $metaKey = "$($catalog.FamilyName)|$colName"
                        if ($paramMetaLookup.ContainsKey($metaKey)) {
                            $paramMeta = $paramMetaLookup[$metaKey]
                            # Only convert if DataType is LENGTH and Unit is FEET
                            if ($paramMeta.DataType -eq "LENGTH" -and $paramMeta.Unit -eq "FEET") {
                                $originalValue = $value
                                $value = Convert-FeetToInches -Value $value
                                if ($originalValue -ne $value) {
                                    $convertedCount++
                                }
                            }
                        }
                    }

                    $row[$colName] = $value
                }
                else {
                    $row[$colName] = ""
                }
            }

            $mergedData += [PSCustomObject]$row
        }
    }

    if ($UnitConversion -eq "ConvertToInches") {
        Write-Log "  Converted $convertedCount values from Feet to Inches"
    }

    # Build metadata
    $metadata = @()
    foreach ($catalog in $allCatalogs) {
        foreach ($param in $catalog.Parameters) {
            if ($param.CleanName -eq "Type Name") { continue }

            $metadata += [PSCustomObject]@{
                Family = $catalog.FamilyName
                ParameterName = $param.CleanName
                FullHeader = $param.FullHeader
                ColumnIndex = $param.Index
                DataType = $param.DataType
                Unit = $param.Unit
                BuiltInId = $param.BuiltInId
                SourceFile = $catalog.FilePath
            }
        }
    }

    # Build settings
    $settings = @(
        [PSCustomObject]@{
            Setting = "UnitConversion"
            Value = $UnitConversion
        },
        [PSCustomObject]@{
            Setting = "GeneratedDate"
            Value = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        },
        [PSCustomObject]@{
            Setting = "SourceFolder"
            Value = $FolderPath
        }
    )

    # Generate output filename
    $xlsxFileName = "${folderName}_TypeCatalog.xlsx"
    $xlsxPath = Join-Path -Path $OutputPath -ChildPath $xlsxFileName

    # Remove existing file if present
    if (Test-Path -Path $xlsxPath) {
        Remove-Item -Path $xlsxPath -Force
    }

    try {
        # Export main data sheet
        $mergedData | Export-Excel -Path $xlsxPath -WorksheetName "TypeCatalog" -AutoSize -FreezeTopRow -BoldTopRow

        # Export metadata sheet
        $metadata | Export-Excel -Path $xlsxPath -WorksheetName "_Metadata" -AutoSize -FreezeTopRow -BoldTopRow

        # Export settings sheet
        $settings | Export-Excel -Path $xlsxPath -WorksheetName "_Settings" -AutoSize -FreezeTopRow -BoldTopRow

        Write-Log "  Created: $xlsxPath"
        Write-Log "  Total types: $($mergedData.Count) from $($allCatalogs.Count) families"

        return @{
            OutputPath = $xlsxPath
            FamilyCount = $allCatalogs.Count
            TypeCount = $mergedData.Count
            ParameterCount = $columnOrder.Count - 2  # Exclude Family and Type Name
        }
    }
    catch {
        Write-Log "  Error creating XLSX: $_" -Level "ERROR"
        return $null
    }
}

function Export-XlsxToTypeCatalogs {
    param(
        [string]$XlsxPath,
        [string]$OutputMode  # "Original" or custom path
    )

    $xlsxName = [System.IO.Path]::GetFileNameWithoutExtension($XlsxPath)
    Write-Log "Processing: $xlsxName"

    try {
        # Check if file was modified since generation
        $xlsxFile = Get-Item -Path $XlsxPath
        $xlsxLastModified = $xlsxFile.LastWriteTime

        # Read settings first to check GeneratedDate
        $settings = @{}
        try {
            $settingsData = Import-Excel -Path $XlsxPath -WorksheetName "_Settings"
            foreach ($setting in $settingsData) {
                $settings[$setting.Setting] = $setting.Value
            }
        }
        catch {
            Write-Log "  No settings sheet found, using defaults" -Level "WARN"
        }

        # Check if file was modified since generation (skip if unchanged)
        if ($settings.ContainsKey("GeneratedDate")) {
            try {
                $generatedDate = [DateTime]::ParseExact($settings["GeneratedDate"], "yyyy-MM-dd HH:mm:ss", $null)
                # Allow 5 second tolerance for file system timing differences
                $timeDiff = ($xlsxLastModified - $generatedDate).TotalSeconds

                if ($timeDiff -lt 5) {
                    Write-Log "  SKIPPED - File not modified since generation (preserving original TXT files)" -Level "SKIP"
                    return @{
                        ExportedCount = 0
                        SkippedCount = 1
                        ErrorCount = 0
                        Families = @()
                    }
                }
                else {
                    Write-Log "  File modified: $([Math]::Round($timeDiff / 60, 1)) minutes after generation"
                }
            }
            catch {
                Write-Log "  Could not parse GeneratedDate, proceeding with export" -Level "WARN"
            }
        }

        # Read main data
        $typeData = Import-Excel -Path $XlsxPath -WorksheetName "TypeCatalog"

        if (-not $typeData -or $typeData.Count -eq 0) {
            Write-Log "  No data found in TypeCatalog sheet" -Level "ERROR"
            return $null
        }

        # Read metadata
        $metadata = Import-Excel -Path $XlsxPath -WorksheetName "_Metadata"

        if (-not $metadata -or $metadata.Count -eq 0) {
            Write-Log "  No metadata found - cannot reconstruct TXT files" -Level "ERROR"
            return $null
        }

        $unitConversion = if ($settings.ContainsKey("UnitConversion")) { $settings["UnitConversion"] } else { "KeepFeet" }
        Write-Log "  Unit conversion setting: $unitConversion"
        Write-Log "  Loaded $($typeData.Count) types"

        # Group data by Family
        $familyGroups = $typeData | Group-Object -Property Family

        # Build metadata lookup per family
        $familyMetadata = @{}
        foreach ($meta in $metadata) {
            $family = $meta.Family
            if (-not $familyMetadata.ContainsKey($family)) {
                $familyMetadata[$family] = @{
                    Parameters = @()
                    SourceFile = $meta.SourceFile
                }
            }
            $familyMetadata[$family].Parameters += $meta
        }

        $results = @{
            ExportedCount = 0
            ErrorCount = 0
            Families = @()
        }

        foreach ($group in $familyGroups) {
            $familyName = $group.Name
            $familyTypes = $group.Group

            if (-not $familyMetadata.ContainsKey($familyName)) {
                Write-Log "  Skipping $familyName - no metadata found" -Level "WARN"
                $results.ErrorCount++
                continue
            }

            $meta = $familyMetadata[$familyName]

            # Sort parameters by original column index
            $sortedParams = $meta.Parameters | Sort-Object -Property ColumnIndex

            # Build header row
            $headerParts = @("")  # First column empty (Type Name column)
            foreach ($param in $sortedParams) {
                $headerParts += $param.FullHeader
            }
            $headerLine = $headerParts -join ","

            # Build data rows
            $dataLines = @()
            foreach ($type in $familyTypes) {
                $valueParts = @()

                # Type Name (first column)
                $typeName = $type."Type Name"
                if ($null -eq $typeName) { $typeName = "" }
                $typeName = $typeName.ToString()
                # Always escape double quotes by doubling them (Revit CSV format)
                $typeName = $typeName -replace '"', '""'
                # Wrap in outer quotes only if contains comma
                if ($typeName -match ',') {
                    $typeName = "`"$typeName`""
                }
                $valueParts += $typeName

                # Parameter values in original order
                foreach ($param in $sortedParams) {
                    $value = $type.($param.ParameterName)
                    if ($null -eq $value) { $value = "" }

                    $valueStr = $value.ToString()

                    # Convert inches back to feet if needed
                    if ($unitConversion -eq "ConvertToInches" -and $param.DataType -eq "LENGTH" -and $param.Unit -eq "FEET") {
                        $valueStr = Convert-InchesToFeet -Value $valueStr
                    }

                    # Always escape double quotes by doubling them (Revit CSV format)
                    $valueStr = $valueStr -replace '"', '""'
                    # Wrap in outer quotes only if contains comma
                    if ($valueStr -match ',') {
                        $valueStr = "`"$valueStr`""
                    }
                    $valueParts += $valueStr
                }

                $dataLines += ($valueParts -join ",")
            }

            # Determine output path
            if ($OutputMode -eq "Original" -and $meta.SourceFile) {
                $outputFile = $meta.SourceFile
            }
            else {
                $outputFile = Join-Path -Path $OutputMode -ChildPath "$familyName.txt"
            }

            # Ensure output directory exists
            $outputDir = Split-Path -Path $outputFile -Parent
            if (-not (Test-Path -Path $outputDir)) {
                New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
            }

            try {
                # Write with UTF-16 encoding (Revit's native format)
                $allLines = @($headerLine) + $dataLines
                $allLines | Out-File -FilePath $outputFile -Encoding Unicode -Force

                Write-Log "  Exported: $familyName ($($familyTypes.Count) types)"
                $results.ExportedCount++
                $results.Families += $familyName
            }
            catch {
                Write-Log "  Error exporting $familyName`: $_" -Level "ERROR"
                $results.ErrorCount++
            }
        }

        return $results
    }
    catch {
        Write-Log "  Error processing XLSX: $_" -Level "ERROR"
        return $null
    }
}

function Show-GUI {
    # Ensure ImportExcel module is available
    if (-not (Ensure-ImportExcelModule)) {
        return
    }

    # Create the main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "GSADUs TypeCatalog Converter"
    $form.Size = New-Object System.Drawing.Size(780, 720)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # Mode Selection GroupBox
    $grpMode = New-Object System.Windows.Forms.GroupBox
    $grpMode.Location = New-Object System.Drawing.Point(20, 15)
    $grpMode.Size = New-Object System.Drawing.Size(720, 50)
    $grpMode.Text = "Conversion Mode"
    $form.Controls.Add($grpMode)

    $rbImport = New-Object System.Windows.Forms.RadioButton
    $rbImport.Location = New-Object System.Drawing.Point(20, 20)
    $rbImport.Size = New-Object System.Drawing.Size(200, 24)
    $rbImport.Text = "Import TXT → XLSX"
    $rbImport.Checked = $true
    $grpMode.Controls.Add($rbImport)

    $rbExport = New-Object System.Windows.Forms.RadioButton
    $rbExport.Location = New-Object System.Drawing.Point(250, 20)
    $rbExport.Size = New-Object System.Drawing.Size(200, 24)
    $rbExport.Text = "Export XLSX → TXT"
    $grpMode.Controls.Add($rbExport)

    # Import Panel
    $pnlImport = New-Object System.Windows.Forms.Panel
    $pnlImport.Location = New-Object System.Drawing.Point(20, 75)
    $pnlImport.Size = New-Object System.Drawing.Size(720, 240)
    $form.Controls.Add($pnlImport)

    # Import - Folders Label
    $lblImportFolders = New-Object System.Windows.Forms.Label
    $lblImportFolders.Location = New-Object System.Drawing.Point(0, 0)
    $lblImportFolders.Size = New-Object System.Drawing.Size(300, 20)
    $lblImportFolders.Text = "Folders to scan (containing .rfa + .txt files):"
    $pnlImport.Controls.Add($lblImportFolders)

    $script:ImportFolderList = New-Object System.Windows.Forms.ListBox
    $script:ImportFolderList.Location = New-Object System.Drawing.Point(0, 23)
    $script:ImportFolderList.Size = New-Object System.Drawing.Size(590, 80)
    $script:ImportFolderList.SelectionMode = "MultiExtended"
    $script:ImportFolderList.HorizontalScrollbar = $true
    $pnlImport.Controls.Add($script:ImportFolderList)

    $btnImportAdd = New-Object System.Windows.Forms.Button
    $btnImportAdd.Location = New-Object System.Drawing.Point(600, 23)
    $btnImportAdd.Size = New-Object System.Drawing.Size(100, 25)
    $btnImportAdd.Text = "Add..."
    $btnImportAdd.Add_Click({
        $selectedPath = Show-ModernFolderDialog -Title "Select a folder containing Revit families" -Owner $form
        if ($selectedPath -and -not $script:ImportFolderList.Items.Contains($selectedPath)) {
            $script:ImportFolderList.Items.Add($selectedPath) | Out-Null
        }
    })
    $pnlImport.Controls.Add($btnImportAdd)

    $btnImportRemove = New-Object System.Windows.Forms.Button
    $btnImportRemove.Location = New-Object System.Drawing.Point(600, 53)
    $btnImportRemove.Size = New-Object System.Drawing.Size(100, 25)
    $btnImportRemove.Text = "Remove"
    $btnImportRemove.Add_Click({
        $selected = @($script:ImportFolderList.SelectedItems)
        foreach ($item in $selected) {
            $script:ImportFolderList.Items.Remove($item)
        }
    })
    $pnlImport.Controls.Add($btnImportRemove)

    $btnImportClear = New-Object System.Windows.Forms.Button
    $btnImportClear.Location = New-Object System.Drawing.Point(600, 83)
    $btnImportClear.Size = New-Object System.Drawing.Size(100, 25)
    $btnImportClear.Text = "Clear All"
    $btnImportClear.Add_Click({
        $script:ImportFolderList.Items.Clear()
    })
    $pnlImport.Controls.Add($btnImportClear)

    # Import - Output Location
    $lblImportOutput = New-Object System.Windows.Forms.Label
    $lblImportOutput.Location = New-Object System.Drawing.Point(0, 115)
    $lblImportOutput.Size = New-Object System.Drawing.Size(100, 20)
    $lblImportOutput.Text = "Output location:"
    $pnlImport.Controls.Add($lblImportOutput)

    $script:cmbImportOutput = New-Object System.Windows.Forms.ComboBox
    $script:cmbImportOutput.Location = New-Object System.Drawing.Point(105, 112)
    $script:cmbImportOutput.Size = New-Object System.Drawing.Size(200, 25)
    $script:cmbImportOutput.DropDownStyle = "DropDownList"
    $script:cmbImportOutput.Items.Add("Same as source folder") | Out-Null
    $script:cmbImportOutput.Items.Add("Custom location...") | Out-Null
    $script:cmbImportOutput.SelectedIndex = 0
    $pnlImport.Controls.Add($script:cmbImportOutput)

    $script:txtImportCustomPath = New-Object System.Windows.Forms.TextBox
    $script:txtImportCustomPath.Location = New-Object System.Drawing.Point(315, 112)
    $script:txtImportCustomPath.Size = New-Object System.Drawing.Size(275, 23)
    $script:txtImportCustomPath.Enabled = $false
    $pnlImport.Controls.Add($script:txtImportCustomPath)

    $btnImportBrowseOutput = New-Object System.Windows.Forms.Button
    $btnImportBrowseOutput.Location = New-Object System.Drawing.Point(600, 111)
    $btnImportBrowseOutput.Size = New-Object System.Drawing.Size(100, 25)
    $btnImportBrowseOutput.Text = "Browse..."
    $btnImportBrowseOutput.Enabled = $false
    $btnImportBrowseOutput.Add_Click({
        $selectedPath = Show-ModernFolderDialog -Title "Select output folder for XLSX files" -Owner $form
        if ($selectedPath) {
            $script:txtImportCustomPath.Text = $selectedPath
        }
    })
    $pnlImport.Controls.Add($btnImportBrowseOutput)

    $script:cmbImportOutput.Add_SelectedIndexChanged({
        $isCustom = $script:cmbImportOutput.SelectedIndex -eq 1
        $script:txtImportCustomPath.Enabled = $isCustom
        $btnImportBrowseOutput.Enabled = $isCustom
    })

    # Unit Conversion Options
    $grpUnits = New-Object System.Windows.Forms.GroupBox
    $grpUnits.Location = New-Object System.Drawing.Point(0, 145)
    $grpUnits.Size = New-Object System.Drawing.Size(590, 50)
    $grpUnits.Text = "Length Unit Conversion (only applies to LENGTH parameters with FEET unit)"
    $pnlImport.Controls.Add($grpUnits)

    $script:rbKeepFeet = New-Object System.Windows.Forms.RadioButton
    $script:rbKeepFeet.Location = New-Object System.Drawing.Point(15, 20)
    $script:rbKeepFeet.Size = New-Object System.Drawing.Size(150, 24)
    $script:rbKeepFeet.Text = "Keep as Feet"
    $grpUnits.Controls.Add($script:rbKeepFeet)

    $script:rbConvertInches = New-Object System.Windows.Forms.RadioButton
    $script:rbConvertInches.Location = New-Object System.Drawing.Point(180, 20)
    $script:rbConvertInches.Size = New-Object System.Drawing.Size(220, 24)
    $script:rbConvertInches.Text = "Convert to Inches (default)"
    $script:rbConvertInches.Checked = $true
    $grpUnits.Controls.Add($script:rbConvertInches)

    # Import Button
    $btnDoImport = New-Object System.Windows.Forms.Button
    $btnDoImport.Location = New-Object System.Drawing.Point(0, 205)
    $btnDoImport.Size = New-Object System.Drawing.Size(180, 35)
    $btnDoImport.Text = "Generate XLSX Files"
    $btnDoImport.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnDoImport.ForeColor = [System.Drawing.Color]::White
    $btnDoImport.FlatStyle = "Flat"
    $pnlImport.Controls.Add($btnDoImport)

    # Export Panel (initially hidden)
    $pnlExport = New-Object System.Windows.Forms.Panel
    $pnlExport.Location = New-Object System.Drawing.Point(20, 75)
    $pnlExport.Size = New-Object System.Drawing.Size(720, 240)
    $pnlExport.Visible = $false
    $form.Controls.Add($pnlExport)

    # Export - Files Label
    $lblExportFiles = New-Object System.Windows.Forms.Label
    $lblExportFiles.Location = New-Object System.Drawing.Point(0, 0)
    $lblExportFiles.Size = New-Object System.Drawing.Size(300, 20)
    $lblExportFiles.Text = "XLSX files to convert back to TXT:"
    $pnlExport.Controls.Add($lblExportFiles)

    $script:ExportFileList = New-Object System.Windows.Forms.ListBox
    $script:ExportFileList.Location = New-Object System.Drawing.Point(0, 23)
    $script:ExportFileList.Size = New-Object System.Drawing.Size(590, 80)
    $script:ExportFileList.SelectionMode = "MultiExtended"
    $script:ExportFileList.HorizontalScrollbar = $true
    $pnlExport.Controls.Add($script:ExportFileList)

    $btnExportAdd = New-Object System.Windows.Forms.Button
    $btnExportAdd.Location = New-Object System.Drawing.Point(600, 23)
    $btnExportAdd.Size = New-Object System.Drawing.Size(100, 25)
    $btnExportAdd.Text = "Add..."
    $btnExportAdd.Add_Click({
        $openDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
        $openDialog.Multiselect = $true
        $openDialog.Title = "Select TypeCatalog XLSX files"
        if ($openDialog.ShowDialog() -eq "OK") {
            foreach ($file in $openDialog.FileNames) {
                if (-not $script:ExportFileList.Items.Contains($file)) {
                    $script:ExportFileList.Items.Add($file) | Out-Null
                }
            }
        }
    })
    $pnlExport.Controls.Add($btnExportAdd)

    $btnExportRemove = New-Object System.Windows.Forms.Button
    $btnExportRemove.Location = New-Object System.Drawing.Point(600, 53)
    $btnExportRemove.Size = New-Object System.Drawing.Size(100, 25)
    $btnExportRemove.Text = "Remove"
    $btnExportRemove.Add_Click({
        $selected = @($script:ExportFileList.SelectedItems)
        foreach ($item in $selected) {
            $script:ExportFileList.Items.Remove($item)
        }
    })
    $pnlExport.Controls.Add($btnExportRemove)

    $btnExportClear = New-Object System.Windows.Forms.Button
    $btnExportClear.Location = New-Object System.Drawing.Point(600, 83)
    $btnExportClear.Size = New-Object System.Drawing.Size(100, 25)
    $btnExportClear.Text = "Clear All"
    $btnExportClear.Add_Click({
        $script:ExportFileList.Items.Clear()
    })
    $pnlExport.Controls.Add($btnExportClear)

    # Export - Output Location
    $lblExportOutput = New-Object System.Windows.Forms.Label
    $lblExportOutput.Location = New-Object System.Drawing.Point(0, 115)
    $lblExportOutput.Size = New-Object System.Drawing.Size(100, 20)
    $lblExportOutput.Text = "Output location:"
    $pnlExport.Controls.Add($lblExportOutput)

    $script:cmbExportOutput = New-Object System.Windows.Forms.ComboBox
    $script:cmbExportOutput.Location = New-Object System.Drawing.Point(105, 112)
    $script:cmbExportOutput.Size = New-Object System.Drawing.Size(200, 25)
    $script:cmbExportOutput.DropDownStyle = "DropDownList"
    $script:cmbExportOutput.Items.Add("Original .txt locations") | Out-Null
    $script:cmbExportOutput.Items.Add("Custom location...") | Out-Null
    $script:cmbExportOutput.SelectedIndex = 0
    $pnlExport.Controls.Add($script:cmbExportOutput)

    $script:txtExportCustomPath = New-Object System.Windows.Forms.TextBox
    $script:txtExportCustomPath.Location = New-Object System.Drawing.Point(315, 112)
    $script:txtExportCustomPath.Size = New-Object System.Drawing.Size(275, 23)
    $script:txtExportCustomPath.Enabled = $false
    $pnlExport.Controls.Add($script:txtExportCustomPath)

    $btnExportBrowseOutput = New-Object System.Windows.Forms.Button
    $btnExportBrowseOutput.Location = New-Object System.Drawing.Point(600, 111)
    $btnExportBrowseOutput.Size = New-Object System.Drawing.Size(100, 25)
    $btnExportBrowseOutput.Text = "Browse..."
    $btnExportBrowseOutput.Enabled = $false
    $btnExportBrowseOutput.Add_Click({
        $selectedPath = Show-ModernFolderDialog -Title "Select output folder for TXT files" -Owner $form
        if ($selectedPath) {
            $script:txtExportCustomPath.Text = $selectedPath
        }
    })
    $pnlExport.Controls.Add($btnExportBrowseOutput)

    $script:cmbExportOutput.Add_SelectedIndexChanged({
        $isCustom = $script:cmbExportOutput.SelectedIndex -eq 1
        $script:txtExportCustomPath.Enabled = $isCustom
        $btnExportBrowseOutput.Enabled = $isCustom
    })

    # Export Note
    $lblExportNote = New-Object System.Windows.Forms.Label
    $lblExportNote.Location = New-Object System.Drawing.Point(0, 150)
    $lblExportNote.Size = New-Object System.Drawing.Size(590, 40)
    $lblExportNote.Text = "Note: Unit conversion will be automatically reversed on export based on the _Settings sheet in the XLSX file."
    $lblExportNote.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
    $pnlExport.Controls.Add($lblExportNote)

    # Export Button
    $btnDoExport = New-Object System.Windows.Forms.Button
    $btnDoExport.Location = New-Object System.Drawing.Point(0, 205)
    $btnDoExport.Size = New-Object System.Drawing.Size(180, 35)
    $btnDoExport.Text = "Export to TXT Files"
    $btnDoExport.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnDoExport.ForeColor = [System.Drawing.Color]::White
    $btnDoExport.FlatStyle = "Flat"
    $pnlExport.Controls.Add($btnDoExport)

    # Mode switch handler
    $rbImport.Add_CheckedChanged({
        $pnlImport.Visible = $rbImport.Checked
        $pnlExport.Visible = -not $rbImport.Checked
    })

    $rbExport.Add_CheckedChanged({
        $pnlImport.Visible = -not $rbExport.Checked
        $pnlExport.Visible = $rbExport.Checked
    })

    # Log Label
    $lblLog = New-Object System.Windows.Forms.Label
    $lblLog.Location = New-Object System.Drawing.Point(20, 325)
    $lblLog.Size = New-Object System.Drawing.Size(100, 20)
    $lblLog.Text = "Activity Log:"
    $form.Controls.Add($lblLog)

    # Clear Log Button
    $btnClearLog = New-Object System.Windows.Forms.Button
    $btnClearLog.Location = New-Object System.Drawing.Point(640, 320)
    $btnClearLog.Size = New-Object System.Drawing.Size(100, 25)
    $btnClearLog.Text = "Clear Log"
    $btnClearLog.Add_Click({
        $script:LogTextBox.Clear()
    })
    $form.Controls.Add($btnClearLog)

    # Log TextBox
    $script:LogTextBox = New-Object System.Windows.Forms.TextBox
    $script:LogTextBox.Location = New-Object System.Drawing.Point(20, 348)
    $script:LogTextBox.Size = New-Object System.Drawing.Size(720, 280)
    $script:LogTextBox.Multiline = $true
    $script:LogTextBox.ScrollBars = "Both"
    $script:LogTextBox.ReadOnly = $true
    $script:LogTextBox.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $script:LogTextBox.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $script:LogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $script:LogTextBox.WordWrap = $false
    $form.Controls.Add($script:LogTextBox)

    # Status Label
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(20, 640)
    $lblStatus.Size = New-Object System.Drawing.Size(720, 25)
    $lblStatus.Text = "Ready. Select a mode and add files/folders to convert."
    $form.Controls.Add($lblStatus)

    # Import Button Handler
    $btnDoImport.Add_Click({
        if ($script:ImportFolderList.Items.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please add at least one folder to scan.",
                "Validation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        $unitConversion = if ($script:rbConvertInches.Checked) { "ConvertToInches" } else { "KeepFeet" }

        $btnDoImport.Enabled = $false
        $btnDoImport.Text = "Processing..."
        $lblStatus.Text = "Generating XLSX files..."

        Write-Log "=========================================="
        Write-Log "Starting Import: TXT → XLSX"
        Write-Log "Unit conversion: $unitConversion"
        Write-Log "=========================================="

        $totalFamilies = 0
        $totalTypes = 0
        $successCount = 0
        $generatedXlsxPaths = @()

        foreach ($folder in $script:ImportFolderList.Items) {
            # Determine output path
            if ($script:cmbImportOutput.SelectedIndex -eq 0) {
                $outputPath = $folder
            }
            else {
                $outputPath = $script:txtImportCustomPath.Text
                if ([string]::IsNullOrWhiteSpace($outputPath)) {
                    Write-Log "No custom output path specified" -Level "ERROR"
                    continue
                }
            }

            $result = Import-TypeCatalogsToXlsx -FolderPath $folder -OutputPath $outputPath -UnitConversion $unitConversion

            if ($result) {
                $totalFamilies += $result.FamilyCount
                $totalTypes += $result.TypeCount
                $successCount++
                # Track generated XLSX path for auto-populating Export list
                if ($result.OutputPath) {
                    $generatedXlsxPaths += $result.OutputPath
                }
            }
        }

        Write-Log "=========================================="
        Write-Log "Import Complete!"
        Write-Log "  Folders processed: $successCount"
        Write-Log "  Total families: $totalFamilies"
        Write-Log "  Total types: $totalTypes"
        Write-Log "=========================================="

        # Auto-populate Export file list with generated XLSX paths
        if ($generatedXlsxPaths.Count -gt 0) {
            Write-Log ""
            Write-Log "Adding generated XLSX files to Export list..."
            foreach ($xlsxPath in $generatedXlsxPaths) {
                if (-not $script:ExportFileList.Items.Contains($xlsxPath)) {
                    $script:ExportFileList.Items.Add($xlsxPath) | Out-Null
                    Write-Log "  Added: $xlsxPath"
                }
            }
            Write-Log "Export list now contains $($script:ExportFileList.Items.Count) file(s)"
        }

        $lblStatus.Text = "Import complete! $totalFamilies families, $totalTypes types"
        $btnDoImport.Enabled = $true
        $btnDoImport.Text = "Generate XLSX Files"

        [System.Windows.Forms.MessageBox]::Show(
            "Import complete!`n`nFolders processed: $successCount`nTotal families: $totalFamilies`nTotal types: $totalTypes`n`nGenerated XLSX files have been added to the Export list.",
            "Import Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    })

    # Export Button Handler
    $btnDoExport.Add_Click({
        if ($script:ExportFileList.Items.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please add at least one XLSX file to convert.",
                "Validation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        $btnDoExport.Enabled = $false
        $btnDoExport.Text = "Processing..."
        $lblStatus.Text = "Exporting to TXT files..."

        Write-Log "=========================================="
        Write-Log "Starting Export: XLSX → TXT"
        Write-Log "=========================================="

        $totalExported = 0
        $totalSkipped = 0
        $totalErrors = 0

        foreach ($xlsxFile in $script:ExportFileList.Items) {
            # Determine output mode
            if ($script:cmbExportOutput.SelectedIndex -eq 0) {
                $outputMode = "Original"
            }
            else {
                $outputMode = $script:txtExportCustomPath.Text
                if ([string]::IsNullOrWhiteSpace($outputMode)) {
                    Write-Log "No custom output path specified" -Level "ERROR"
                    continue
                }
            }

            $result = Export-XlsxToTypeCatalogs -XlsxPath $xlsxFile -OutputMode $outputMode

            if ($result) {
                $totalExported += $result.ExportedCount
                $totalErrors += $result.ErrorCount
                if ($result.SkippedCount) {
                    $totalSkipped += $result.SkippedCount
                }
            }
        }

        Write-Log "=========================================="
        Write-Log "Export Complete!"
        Write-Log "  Families exported: $totalExported"
        Write-Log "  Files skipped (unchanged): $totalSkipped"
        Write-Log "  Errors: $totalErrors"
        Write-Log "=========================================="

        $lblStatus.Text = "Export complete! $totalExported families exported, $totalSkipped skipped"
        $btnDoExport.Enabled = $true
        $btnDoExport.Text = "Export to TXT Files"

        $message = "Export complete!`n`nFamilies exported: $totalExported`nErrors: $totalErrors"
        if ($totalSkipped -gt 0) {
            $message += "`n`nSkipped $totalSkipped file(s) that were not modified`n(original TXT files preserved)"
        }

        [System.Windows.Forms.MessageBox]::Show(
            $message,
            "Export Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    })

    # Show the form
    Write-Log "GSADUs TypeCatalog Converter started"
    Write-Log "Modes: Import (TXT → XLSX) | Export (XLSX → TXT)"
    Write-Log ""
    Write-Log "ImportExcel module loaded successfully"

    [void]$form.ShowDialog()
}

# Run the GUI
Show-GUI
