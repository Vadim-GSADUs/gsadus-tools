#Requires -Version 5.1
<#
.SYNOPSIS
    Family Catalog Generator - Scans Revit family folders and generates CSV catalogs.

.DESCRIPTION
    This tool searches specified folders for .rfa files and compiles detailed CSV catalogs
    including file names, paths, sizes, and dates. Generates both per-category and master catalogs.

.NOTES
    Author: GSADUs
    Created: 2026-01-30
    Updated: 2026-01-30 - Added modern folder dialog
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

# Default category paths
$script:DefaultCategories = @(
    "G:\Shared drives\GSDE Projects\CADD\RevitFamily.Biz\Windows",
    "G:\Shared drives\GSDE Projects\CADD\RevitFamily.Biz\Cabinets",
    "G:\Shared drives\GSDE Projects\CADD\RevitFamily.Biz\Doors",
    "G:\Shared drives\GSDE Projects\CADD\RevitFamily.Biz\Electrical",
    "G:\Shared drives\GSDE Projects\CADD\RevitFamily.Biz\Finish Carpentry",
    "G:\Shared drives\GSDE Projects\CADD\RevitFamily.Biz\Railings"
)

$script:DefaultOutputPath = "G:\Shared drives\GSDE Projects\CADD\RevitFamily.Biz"

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

function Format-FileSize {
    param([long]$Bytes)

    if ($Bytes -ge 1MB) {
        return "{0:N2} MB" -f ($Bytes / 1MB)
    }
    elseif ($Bytes -ge 1KB) {
        return "{0:N2} KB" -f ($Bytes / 1KB)
    }
    else {
        return "$Bytes B"
    }
}

function Get-FamilyFiles {
    param([string]$FolderPath)

    $familyFiles = @()

    try {
        $rfaFiles = Get-ChildItem -Path $FolderPath -Filter "*.rfa" -File -Recurse -ErrorAction SilentlyContinue

        foreach ($file in $rfaFiles) {
            # Calculate relative path from the category folder
            $relativePath = $file.DirectoryName.Replace($FolderPath, "").TrimStart("\", "/")
            if ([string]::IsNullOrEmpty($relativePath)) {
                $relativePath = "\"
            }

            $familyFiles += [PSCustomObject]@{
                FileName        = $file.Name
                FamilyName      = $file.BaseName
                RelativePath    = $relativePath
                FullPath        = $file.FullName
                Category        = (Split-Path -Path $FolderPath -Leaf)
                FileSizeBytes   = $file.Length
                FileSize        = Format-FileSize -Bytes $file.Length
                DateModified    = $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                DateCreated     = $file.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
                ParentFolder    = $file.Directory.Name
            }
        }
    }
    catch {
        Write-Log "Error scanning folder $FolderPath`: $_" -Level "ERROR"
    }

    return $familyFiles
}

function Export-CategoryCatalog {
    param(
        [string]$CategoryPath,
        [string]$OutputFolder,
        [bool]$SaveInCategoryFolder = $false
    )

    $categoryName = Split-Path -Path $CategoryPath -Leaf
    Write-Log "Scanning category: $categoryName"

    $families = Get-FamilyFiles -FolderPath $CategoryPath

    if ($families.Count -eq 0) {
        Write-Log "  No .rfa files found in $categoryName" -Level "WARN"
        return @{
            CategoryName = $categoryName
            FileCount = 0
            FilePath = $null
            Families = @()
        }
    }

    Write-Log "  Found $($families.Count) family files"

    # Determine output location
    if ($SaveInCategoryFolder) {
        $csvPath = Join-Path -Path $CategoryPath -ChildPath "_FamilyCatalog_$categoryName.csv"
    }
    else {
        $csvPath = Join-Path -Path $OutputFolder -ChildPath "FamilyCatalog_$categoryName.csv"
    }

    try {
        $families | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Log "  Saved catalog to: $csvPath"
    }
    catch {
        Write-Log "  Error saving catalog: $_" -Level "ERROR"
        $csvPath = $null
    }

    return @{
        CategoryName = $categoryName
        FileCount = $families.Count
        FilePath = $csvPath
        Families = $families
    }
}

function Generate-Catalogs {
    param(
        [string[]]$CategoryPaths,
        [string]$OutputFolder,
        [bool]$SaveInCategoryFolders = $false,
        [bool]$GenerateMaster = $true
    )

    $results = @{
        TotalFiles = 0
        Categories = @()
        MasterCatalogPath = $null
        Errors = 0
    }

    $allFamilies = @()

    # Process each category
    foreach ($categoryPath in $CategoryPaths) {
        if (-not (Test-Path -Path $categoryPath)) {
            Write-Log "Category path does not exist: $categoryPath" -Level "WARN"
            $results.Errors++
            continue
        }

        $categoryResult = Export-CategoryCatalog -CategoryPath $categoryPath -OutputFolder $OutputFolder -SaveInCategoryFolder $SaveInCategoryFolders

        $results.Categories += $categoryResult
        $results.TotalFiles += $categoryResult.FileCount
        $allFamilies += $categoryResult.Families
    }

    # Generate master catalog
    if ($GenerateMaster -and $allFamilies.Count -gt 0) {
        Write-Log "=========================================="
        Write-Log "Generating Master Catalog..."

        $masterPath = Join-Path -Path $OutputFolder -ChildPath "FamilyCatalog_MASTER.csv"

        try {
            $allFamilies | Export-Csv -Path $masterPath -NoTypeInformation -Encoding UTF8
            Write-Log "Master catalog saved: $masterPath"
            Write-Log "Total families in master: $($allFamilies.Count)"
            $results.MasterCatalogPath = $masterPath
        }
        catch {
            Write-Log "Error saving master catalog: $_" -Level "ERROR"
            $results.Errors++
        }
    }

    return $results
}

function Show-GUI {
    # Create the main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "GSADUs Family Catalog Generator"
    $form.Size = New-Object System.Drawing.Size(750, 650)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # Category Paths Label
    $lblCategories = New-Object System.Windows.Forms.Label
    $lblCategories.Location = New-Object System.Drawing.Point(20, 15)
    $lblCategories.Size = New-Object System.Drawing.Size(200, 20)
    $lblCategories.Text = "Category Folders to Scan:"
    $form.Controls.Add($lblCategories)

    # Category ListBox
    $script:CategoryListBox = New-Object System.Windows.Forms.ListBox
    $script:CategoryListBox.Location = New-Object System.Drawing.Point(20, 38)
    $script:CategoryListBox.Size = New-Object System.Drawing.Size(590, 100)
    $script:CategoryListBox.SelectionMode = "MultiExtended"
    $script:CategoryListBox.HorizontalScrollbar = $true
    foreach ($cat in $script:DefaultCategories) {
        $script:CategoryListBox.Items.Add($cat) | Out-Null
    }
    $form.Controls.Add($script:CategoryListBox)

    # Add Category Button
    $btnAddCategory = New-Object System.Windows.Forms.Button
    $btnAddCategory.Location = New-Object System.Drawing.Point(620, 38)
    $btnAddCategory.Size = New-Object System.Drawing.Size(100, 28)
    $btnAddCategory.Text = "Add..."
    $btnAddCategory.Add_Click({
        $selectedPath = Show-ModernFolderDialog -Title "Select a family category folder to add" -Owner $form
        if ($selectedPath -and -not $script:CategoryListBox.Items.Contains($selectedPath)) {
            $script:CategoryListBox.Items.Add($selectedPath) | Out-Null
        }
    })
    $form.Controls.Add($btnAddCategory)

    # Remove Category Button
    $btnRemoveCategory = New-Object System.Windows.Forms.Button
    $btnRemoveCategory.Location = New-Object System.Drawing.Point(620, 72)
    $btnRemoveCategory.Size = New-Object System.Drawing.Size(100, 28)
    $btnRemoveCategory.Text = "Remove"
    $btnRemoveCategory.Add_Click({
        $selectedItems = @($script:CategoryListBox.SelectedItems)
        foreach ($item in $selectedItems) {
            $script:CategoryListBox.Items.Remove($item)
        }
    })
    $form.Controls.Add($btnRemoveCategory)

    # Clear All Button
    $btnClearCategories = New-Object System.Windows.Forms.Button
    $btnClearCategories.Location = New-Object System.Drawing.Point(620, 106)
    $btnClearCategories.Size = New-Object System.Drawing.Size(100, 28)
    $btnClearCategories.Text = "Clear All"
    $btnClearCategories.Add_Click({
        $script:CategoryListBox.Items.Clear()
    })
    $form.Controls.Add($btnClearCategories)

    # Output Path Label
    $lblOutput = New-Object System.Windows.Forms.Label
    $lblOutput.Location = New-Object System.Drawing.Point(20, 150)
    $lblOutput.Size = New-Object System.Drawing.Size(120, 23)
    $lblOutput.Text = "Output Folder:"
    $form.Controls.Add($lblOutput)

    # Output Path TextBox
    $txtOutput = New-Object System.Windows.Forms.TextBox
    $txtOutput.Location = New-Object System.Drawing.Point(140, 148)
    $txtOutput.Size = New-Object System.Drawing.Size(470, 23)
    $txtOutput.Text = $script:DefaultOutputPath
    $form.Controls.Add($txtOutput)

    # Output Browse Button
    $btnBrowseOutput = New-Object System.Windows.Forms.Button
    $btnBrowseOutput.Location = New-Object System.Drawing.Point(620, 147)
    $btnBrowseOutput.Size = New-Object System.Drawing.Size(100, 25)
    $btnBrowseOutput.Text = "Browse..."
    $btnBrowseOutput.Add_Click({
        $selectedPath = Show-ModernFolderDialog -Title "Select the output folder for CSV catalogs" -Owner $form
        if ($selectedPath) {
            $txtOutput.Text = $selectedPath
        }
    })
    $form.Controls.Add($btnBrowseOutput)

    # Options GroupBox
    $grpOptions = New-Object System.Windows.Forms.GroupBox
    $grpOptions.Location = New-Object System.Drawing.Point(20, 180)
    $grpOptions.Size = New-Object System.Drawing.Size(700, 55)
    $grpOptions.Text = "Options"
    $form.Controls.Add($grpOptions)

    # Save in category folders checkbox
    $chkSaveInCategory = New-Object System.Windows.Forms.CheckBox
    $chkSaveInCategory.Location = New-Object System.Drawing.Point(15, 22)
    $chkSaveInCategory.Size = New-Object System.Drawing.Size(250, 24)
    $chkSaveInCategory.Text = "Also save CSVs in category folders"
    $chkSaveInCategory.Checked = $true
    $grpOptions.Controls.Add($chkSaveInCategory)

    # Generate master checkbox
    $chkMaster = New-Object System.Windows.Forms.CheckBox
    $chkMaster.Location = New-Object System.Drawing.Point(280, 22)
    $chkMaster.Size = New-Object System.Drawing.Size(200, 24)
    $chkMaster.Text = "Generate Master Catalog"
    $chkMaster.Checked = $true
    $grpOptions.Controls.Add($chkMaster)

    # Generate Button
    $btnGenerate = New-Object System.Windows.Forms.Button
    $btnGenerate.Location = New-Object System.Drawing.Point(20, 245)
    $btnGenerate.Size = New-Object System.Drawing.Size(150, 35)
    $btnGenerate.Text = "Generate Catalogs"
    $btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnGenerate.ForeColor = [System.Drawing.Color]::White
    $btnGenerate.FlatStyle = "Flat"
    $form.Controls.Add($btnGenerate)

    # Open Output Button
    $btnOpenOutput = New-Object System.Windows.Forms.Button
    $btnOpenOutput.Location = New-Object System.Drawing.Point(180, 245)
    $btnOpenOutput.Size = New-Object System.Drawing.Size(130, 35)
    $btnOpenOutput.Text = "Open Output Folder"
    $btnOpenOutput.Add_Click({
        if (Test-Path -Path $txtOutput.Text) {
            Start-Process explorer.exe -ArgumentList $txtOutput.Text
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "Output folder does not exist.",
                "Folder Not Found",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    })
    $form.Controls.Add($btnOpenOutput)

    # Clear Log Button
    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Location = New-Object System.Drawing.Point(320, 245)
    $btnClear.Size = New-Object System.Drawing.Size(100, 35)
    $btnClear.Text = "Clear Log"
    $btnClear.Add_Click({
        $script:LogTextBox.Clear()
    })
    $form.Controls.Add($btnClear)

    # Progress Bar
    $script:ProgressBar = New-Object System.Windows.Forms.ProgressBar
    $script:ProgressBar.Location = New-Object System.Drawing.Point(430, 252)
    $script:ProgressBar.Size = New-Object System.Drawing.Size(290, 22)
    $script:ProgressBar.Style = "Continuous"
    $form.Controls.Add($script:ProgressBar)

    # Log Label
    $lblLog = New-Object System.Windows.Forms.Label
    $lblLog.Location = New-Object System.Drawing.Point(20, 295)
    $lblLog.Size = New-Object System.Drawing.Size(100, 20)
    $lblLog.Text = "Activity Log:"
    $form.Controls.Add($lblLog)

    # Log TextBox
    $script:LogTextBox = New-Object System.Windows.Forms.TextBox
    $script:LogTextBox.Location = New-Object System.Drawing.Point(20, 318)
    $script:LogTextBox.Size = New-Object System.Drawing.Size(700, 240)
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
    $lblStatus.Location = New-Object System.Drawing.Point(20, 570)
    $lblStatus.Size = New-Object System.Drawing.Size(700, 25)
    $lblStatus.Text = "Ready. Add category folders and click 'Generate Catalogs' to begin."
    $form.Controls.Add($lblStatus)

    # Generate Button Click Handler
    $btnGenerate.Add_Click({
        # Validate
        if ($script:CategoryListBox.Items.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please add at least one category folder to scan.",
                "Validation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        $outputPath = $txtOutput.Text.Trim()
        if ([string]::IsNullOrEmpty($outputPath)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please specify an output folder.",
                "Validation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        # Create output folder if needed
        if (-not (Test-Path -Path $outputPath)) {
            try {
                New-Item -Path $outputPath -ItemType Directory -Force | Out-Null
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Failed to create output folder: $_",
                    "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                return
            }
        }

        # Disable button during operation
        $btnGenerate.Enabled = $false
        $btnGenerate.Text = "Scanning..."
        $lblStatus.Text = "Generating catalogs... Please wait."
        $script:ProgressBar.Value = 0

        Write-Log "=========================================="
        Write-Log "Starting Family Catalog Generation"
        Write-Log "Output folder: $outputPath"
        Write-Log "Categories to scan: $($script:CategoryListBox.Items.Count)"
        Write-Log "=========================================="

        # Get category paths
        $categoryPaths = @($script:CategoryListBox.Items)

        # Update progress
        $script:ProgressBar.Maximum = $categoryPaths.Count + 1
        $script:ProgressBar.Value = 0

        # Generate catalogs
        $results = Generate-Catalogs -CategoryPaths $categoryPaths `
            -OutputFolder $outputPath `
            -SaveInCategoryFolders $chkSaveInCategory.Checked `
            -GenerateMaster $chkMaster.Checked

        $script:ProgressBar.Value = $script:ProgressBar.Maximum

        Write-Log "=========================================="
        Write-Log "Catalog Generation Complete!"
        Write-Log "  Total categories: $($results.Categories.Count)"
        Write-Log "  Total families: $($results.TotalFiles)"
        Write-Log "  Errors: $($results.Errors)"
        if ($results.MasterCatalogPath) {
            Write-Log "  Master catalog: $($results.MasterCatalogPath)"
        }
        Write-Log "=========================================="

        # Show summary per category
        Write-Log ""
        Write-Log "Summary by Category:"
        foreach ($cat in $results.Categories) {
            Write-Log "  $($cat.CategoryName): $($cat.FileCount) families"
        }

        $lblStatus.Text = "Complete! Total: $($results.TotalFiles) families in $($results.Categories.Count) categories"

        # Re-enable button
        $btnGenerate.Enabled = $true
        $btnGenerate.Text = "Generate Catalogs"

        # Show completion message
        $message = "Catalog generation complete!`n`n" +
                   "Categories scanned: $($results.Categories.Count)`n" +
                   "Total families found: $($results.TotalFiles)`n" +
                   "Errors: $($results.Errors)"

        if ($results.MasterCatalogPath) {
            $message += "`n`nMaster catalog saved to:`n$($results.MasterCatalogPath)"
        }

        [System.Windows.Forms.MessageBox]::Show(
            $message,
            "Generation Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    })

    # Show the form
    Write-Log "GSADUs Family Catalog Generator started"
    Write-Log "Default output: $script:DefaultOutputPath"
    Write-Log ""
    Write-Log "Loaded $($script:DefaultCategories.Count) default category paths"

    [void]$form.ShowDialog()
}

# Run the GUI
Show-GUI
