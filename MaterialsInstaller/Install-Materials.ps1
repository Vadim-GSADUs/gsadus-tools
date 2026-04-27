#Requires -Version 5.1
<#
.SYNOPSIS
    Materials Installer - Aggregates Revit material textures into the shared
    GSDE Projects materials folder, from either a Revit family library or
    one-off downloaded family bundles (.zip / .rar / .7z).

.NOTES
    Author: GSADUs
    Created: 2026-01-28
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Modern (Explorer-style) folder picker via IFileOpenDialog.
# Supports pasting a full path into the "Folder name:" field instead of clicking through a tree.
if (-not ('MaterialsInstaller.FolderPicker' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.IO;

namespace MaterialsInstaller {
    [ComImport, ClassInterface(ClassInterfaceType.None), Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")]
    public class FileOpenDialogRCW { }

    [ComImport, Guid("42F85136-DB7E-439C-85F1-E4075D135FC8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IFileOpenDialog {
        [PreserveSig] uint Show(IntPtr hwnd);
        void SetFileTypes();
        void SetFileTypeIndex();
        void GetFileTypeIndex();
        void Advise();
        void Unadvise();
        void SetOptions(uint fos);
        void GetOptions();
        void SetDefaultFolder(IShellItem psi);
        void SetFolder(IShellItem psi);
        void GetFolder();
        void GetCurrentSelection();
        void SetFileName();
        void GetFileName();
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string title);
        void SetOkButtonLabel();
        void SetFileNameLabel();
        void GetResult(out IShellItem ppsi);
        void AddPlace();
        void SetDefaultExtension();
        void Close();
        void SetClientGuid();
        void ClearClientData();
        void SetFilter();
        void GetResults();
        void GetSelectedItems();
    }

    [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IShellItem {
        void BindToHandler();
        void GetParent();
        void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string name);
        void GetAttributes();
        void Compare();
    }

    public static class FolderPicker {
        const uint FOS_PICKFOLDERS      = 0x20;
        const uint FOS_FORCEFILESYSTEM  = 0x40;
        const uint FOS_NOCHANGEDIR      = 0x8;
        const uint SIGDN_FILESYSPATH    = 0x80058000;

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        static extern int SHCreateItemFromParsingName(string path, IntPtr pbc, [In] ref Guid riid, [Out, MarshalAs(UnmanagedType.Interface)] out IShellItem ppv);

        public static string Pick(string initialPath, string title, IntPtr owner) {
            var dlg = (IFileOpenDialog)new FileOpenDialogRCW();
            try {
                dlg.SetOptions(FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_NOCHANGEDIR);
                if (!string.IsNullOrEmpty(title)) dlg.SetTitle(title);
                if (!string.IsNullOrEmpty(initialPath) && Directory.Exists(initialPath)) {
                    Guid iid = typeof(IShellItem).GUID;
                    IShellItem si;
                    if (SHCreateItemFromParsingName(initialPath, IntPtr.Zero, ref iid, out si) == 0) {
                        dlg.SetFolder(si);
                    }
                }
                if (dlg.Show(owner) != 0) return null;
                IShellItem result;
                dlg.GetResult(out result);
                string path;
                result.GetDisplayName(SIGDN_FILESYSPATH, out path);
                return path;
            } finally {
                Marshal.ReleaseComObject(dlg);
            }
        }
    }
}
"@
}

function Select-Folder {
    param([string]$InitialPath, [string]$Title, $Owner)
    $hwnd = [IntPtr]::Zero
    if ($Owner -and $Owner.Handle) { $hwnd = $Owner.Handle }
    return [MaterialsInstaller.FolderPicker]::Pick($InitialPath, $Title, $hwnd)
}

# Default paths
# Destination is the canonical shared materials folder used by everyone's Revit
# (see Vault/wiki/curated/architextures-material-sync.md for the per-PC Revit setup).
$script:DefaultSourcePath = "G:\Shared drives\GSDE Projects\CADD\RevitFamily.Biz"
$script:DefaultDestPath   = "G:\Shared drives\GSDE Projects\CADD\Materials"
$script:DownloadsPath     = [Environment]::GetFolderPath('UserProfile') + '\Downloads'

$script:ImageExtensions        = @('.jpg','.jpeg','.png','.bmp','.tif','.tiff','.dds','.tga')
$script:RevitFamilyExtensions  = @('.rfa','.rvt','.rte','.rft')
$script:ArchiveExtensions      = @('.zip','.rar','.7z')

$script:SevenZipPath = $null
$script:LogTextBox   = $null

#--------------------------------------------------------------------
# Logging
#--------------------------------------------------------------------
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp] [$Level] $Message"
    if ($script:LogTextBox) {
        $script:LogTextBox.AppendText("$entry`r`n")
        $script:LogTextBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

#--------------------------------------------------------------------
# 7-Zip helpers
#--------------------------------------------------------------------
function Find-7Zip {
    $candidates = @(
        "${env:ProgramFiles}\7-Zip\7z.exe",
        "${env:ProgramFiles(x86)}\7-Zip\7z.exe"
    )
    foreach ($c in $candidates) { if (Test-Path -LiteralPath $c) { return $c } }
    return $null
}

function Install-7Zip {
    Write-Log "7-Zip not detected. Installing via winget (this may trigger a UAC prompt)..."
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget) {
        Write-Log "winget not available. Install 7-Zip manually from https://www.7-zip.org/" -Level "ERROR"
        return $null
    }
    try {
        $proc = Start-Process -FilePath winget -ArgumentList @(
            'install','--id','7zip.7zip','-e',
            '--accept-source-agreements',
            '--accept-package-agreements',
            '--silent'
        ) -Wait -PassThru -WindowStyle Hidden
        if ($proc.ExitCode -eq 0) {
            $found = Find-7Zip
            if ($found) { Write-Log "7-Zip installed: $found" } else { Write-Log "winget reported success but 7z.exe not found." -Level "WARN" }
            return $found
        }
        Write-Log "winget exited with code $($proc.ExitCode)" -Level "ERROR"
        return $null
    } catch {
        Write-Log "winget install failed: $_" -Level "ERROR"
        return $null
    }
}

function Ensure-7Zip {
    param([switch]$PromptIfMissing)
    if (-not $script:SevenZipPath) { $script:SevenZipPath = Find-7Zip }
    if ($script:SevenZipPath) { return $script:SevenZipPath }

    if ($PromptIfMissing) {
        $answer = [System.Windows.Forms.MessageBox]::Show(
            "7-Zip is required to extract some bundle archives (e.g. RAR files renamed as .zip).`n`nInstall it now via winget?",
            "Install 7-Zip?",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
            $script:SevenZipPath = Install-7Zip
        } else {
            Write-Log "Continuing without 7-Zip (only standard .zip archives will be extracted)." -Level "WARN"
        }
    }
    return $script:SevenZipPath
}

function Expand-AnyArchive {
    param([string]$ArchivePath, [string]$Destination)
    $sevenZip = Ensure-7Zip
    if ($sevenZip) {
        # Build a single quoted command line — Start-Process -ArgumentList @() does not
        # auto-quote individual array elements, so paths with spaces get tokenized.
        $argLine = 'x -y "-o{0}" "{1}"' -f $Destination, $ArchivePath
        $proc = Start-Process -FilePath $sevenZip -ArgumentList $argLine -Wait -PassThru -WindowStyle Hidden
        if ($proc.ExitCode -eq 0) { return $true }
        Write-Log "  7z exit code $($proc.ExitCode) on $(Split-Path -Leaf $ArchivePath)" -Level "WARN"
        return $false
    }
    try {
        Expand-Archive -LiteralPath $ArchivePath -DestinationPath $Destination -Force -ErrorAction Stop
        return $true
    } catch {
        Write-Log "  Could not extract $(Split-Path -Leaf $ArchivePath): $_" -Level "WARN"
        return $false
    }
}

#--------------------------------------------------------------------
# Source scan: walks folders + extracts archives, applying conservative rule
#--------------------------------------------------------------------
function Get-SourceImages {
    param([string]$SourcePath, [string]$TempRoot, [int]$Depth = 0)
    $images = New-Object System.Collections.Generic.List[string]
    if ($Depth -gt 8) { Write-Log "  Depth limit reached at $SourcePath" -Level "WARN"; return $images }

    if (Test-Path -LiteralPath $SourcePath -PathType Leaf) {
        $ext = [System.IO.Path]::GetExtension($SourcePath).ToLower()
        if ($script:ArchiveExtensions -contains $ext) {
            return Invoke-ArchiveScan -ArchivePath $SourcePath -TempRoot $TempRoot -Depth $Depth
        }
        return $images
    }

    # Directory walk — conservative rule
    Get-ChildItem -LiteralPath $SourcePath -File -ErrorAction SilentlyContinue | ForEach-Object {
        $ext = $_.Extension.ToLower()
        if ($script:ImageExtensions -contains $ext) {
            $parentName = Split-Path -Leaf $_.DirectoryName
            if ($parentName -match '^(Materials|Textures)$') {
                $images.Add($_.FullName) | Out-Null
            }
        } elseif ($script:ArchiveExtensions -contains $ext) {
            $sub = Invoke-ArchiveScan -ArchivePath $_.FullName -TempRoot $TempRoot -Depth ($Depth + 1)
            foreach ($p in $sub) { $images.Add($p) | Out-Null }
        }
    }
    Get-ChildItem -LiteralPath $SourcePath -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $sub = Get-SourceImages -SourcePath $_.FullName -TempRoot $TempRoot -Depth ($Depth + 1)
        foreach ($p in $sub) { $images.Add($p) | Out-Null }
    }
    return $images
}

function Invoke-ArchiveScan {
    param([string]$ArchivePath, [string]$TempRoot, [int]$Depth)
    $images = New-Object System.Collections.Generic.List[string]
    $name = Split-Path -Leaf $ArchivePath
    $extractDir = Join-Path $TempRoot ([System.IO.Path]::GetRandomFileName())
    New-Item -Path $extractDir -ItemType Directory -Force | Out-Null

    Write-Log "  Extracting: $name"
    if (-not (Expand-AnyArchive -ArchivePath $ArchivePath -Destination $extractDir)) {
        Write-Log "    Skipped (could not extract): $name" -Level "WARN"
        return $images
    }

    $allFiles = Get-ChildItem -LiteralPath $extractDir -File -Recurse -ErrorAction SilentlyContinue
    $hasFamily = $false
    foreach ($f in $allFiles) {
        if ($script:RevitFamilyExtensions -contains $f.Extension.ToLower()) { $hasFamily = $true; break }
    }

    if (-not $hasFamily) {
        # Texture archive — pull all images regardless of folder name
        foreach ($f in $allFiles) {
            if ($script:ImageExtensions -contains $f.Extension.ToLower()) {
                $images.Add($f.FullName) | Out-Null
            }
        }
        Write-Log "    Texture archive: $($images.Count) image(s) found in $name"
    } else {
        # Mixed bundle — recurse with conservative rule
        Write-Log "    Mixed bundle (contains Revit family): walking $name with conservative rule"
        $sub = Get-SourceImages -SourcePath $extractDir -TempRoot $TempRoot -Depth ($Depth + 1)
        foreach ($p in $sub) { $images.Add($p) | Out-Null }
    }
    return $images
}

#--------------------------------------------------------------------
# Shared copy + dedupe
#--------------------------------------------------------------------
function Copy-ImageList {
    param(
        [System.Collections.Generic.List[string]]$ImageList,
        [string]$Destination
    )
    $results = @{ Copied = 0; Skipped = 0; Errors = 0; TotalFound = $ImageList.Count }

    if (-not (Test-Path -LiteralPath $Destination)) {
        Write-Log "Creating destination folder: $Destination"
        try { New-Item -Path $Destination -ItemType Directory -Force | Out-Null }
        catch { Write-Log "Failed to create destination: $_" -Level "ERROR"; return $results }
    }

    $existing = @{}
    Get-ChildItem -LiteralPath $Destination -File -ErrorAction SilentlyContinue | ForEach-Object {
        $existing[$_.Name.ToLower()] = $_.FullName
    }
    Write-Log "Found $($existing.Count) existing file(s) in destination"

    foreach ($srcPath in $ImageList) {
        $fileName = Split-Path -Leaf $srcPath
        $key = $fileName.ToLower()
        if ($existing.ContainsKey($key)) {
            Write-Log "  Skipped (duplicate): $fileName" -Level "SKIP"
            $results.Skipped++
            continue
        }
        $destFile = Join-Path $Destination $fileName
        try {
            Copy-Item -LiteralPath $srcPath -Destination $destFile -Force
            Write-Log "  Copied: $fileName"
            $results.Copied++
            $existing[$key] = $destFile
        } catch {
            Write-Log "  Error copying $fileName : $_" -Level "ERROR"
            $results.Errors++
        }
    }
    return $results
}

#--------------------------------------------------------------------
# GUI
#--------------------------------------------------------------------
function Show-GUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "GSADUs Materials Installer"
    $form.Size = New-Object System.Drawing.Size(720, 560)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # Source row
    $lblSource = New-Object System.Windows.Forms.Label
    $lblSource.Location = New-Object System.Drawing.Point(20, 20)
    $lblSource.Size = New-Object System.Drawing.Size(100, 23)
    $lblSource.Text = "Source:"
    $form.Controls.Add($lblSource)

    $txtSource = New-Object System.Windows.Forms.TextBox
    $txtSource.Location = New-Object System.Drawing.Point(120, 18)
    $txtSource.Size = New-Object System.Drawing.Size(385, 23)
    $txtSource.Text = $script:DefaultSourcePath
    $form.Controls.Add($txtSource)

    $btnBrowseSourceFolder = New-Object System.Windows.Forms.Button
    $btnBrowseSourceFolder.Location = New-Object System.Drawing.Point(515, 17)
    $btnBrowseSourceFolder.Size = New-Object System.Drawing.Size(80, 25)
    $btnBrowseSourceFolder.Text = "Browse..."
    $btnBrowseSourceFolder.Add_Click({
        $picked = Select-Folder -InitialPath $txtSource.Text -Title "Select source folder" -Owner $form
        if ($picked) { $txtSource.Text = $picked }
    })
    $form.Controls.Add($btnBrowseSourceFolder)

    # Destination row
    $lblDest = New-Object System.Windows.Forms.Label
    $lblDest.Location = New-Object System.Drawing.Point(20, 55)
    $lblDest.Size = New-Object System.Drawing.Size(100, 23)
    $lblDest.Text = "Destination:"
    $form.Controls.Add($lblDest)

    $txtDest = New-Object System.Windows.Forms.TextBox
    $txtDest.Location = New-Object System.Drawing.Point(120, 53)
    $txtDest.Size = New-Object System.Drawing.Size(470, 23)
    $txtDest.Text = $script:DefaultDestPath
    $form.Controls.Add($txtDest)

    $btnBrowseDest = New-Object System.Windows.Forms.Button
    $btnBrowseDest.Location = New-Object System.Drawing.Point(600, 52)
    $btnBrowseDest.Size = New-Object System.Drawing.Size(80, 25)
    $btnBrowseDest.Text = "Browse..."
    $btnBrowseDest.Add_Click({
        $picked = Select-Folder -InitialPath $txtDest.Text -Title "Select destination folder" -Owner $form
        if ($picked) { $txtDest.Text = $picked }
    })
    $form.Controls.Add($btnBrowseDest)

    # Source: also offer "File..." for a single archive
    $btnBrowseSourceFile = New-Object System.Windows.Forms.Button
    $btnBrowseSourceFile.Location = New-Object System.Drawing.Point(600, 17)
    $btnBrowseSourceFile.Size = New-Object System.Drawing.Size(80, 25)
    $btnBrowseSourceFile.Text = "File..."
    $btnBrowseSourceFile.Add_Click({
        $fd = New-Object System.Windows.Forms.OpenFileDialog
        $fd.Title = "Pick a single archive (.zip / .rar / .7z)"
        $fd.Filter = "Archives (*.zip;*.rar;*.7z)|*.zip;*.rar;*.7z|All files (*.*)|*.*"
        $fd.InitialDirectory = $script:DownloadsPath
        if ($fd.ShowDialog() -eq "OK") { $txtSource.Text = $fd.FileName }
    })
    $form.Controls.Add($btnBrowseSourceFile)

    # Action buttons row
    $btnInstall = New-Object System.Windows.Forms.Button
    $btnInstall.Location = New-Object System.Drawing.Point(20, 125)
    $btnInstall.Size = New-Object System.Drawing.Size(180, 35)
    $btnInstall.Text = "Install Materials"
    $btnInstall.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnInstall.ForeColor = [System.Drawing.Color]::White
    $btnInstall.FlatStyle = "Flat"
    $form.Controls.Add($btnInstall)

    $btnOpenDest = New-Object System.Windows.Forms.Button
    $btnOpenDest.Location = New-Object System.Drawing.Point(210, 125)
    $btnOpenDest.Size = New-Object System.Drawing.Size(150, 35)
    $btnOpenDest.Text = "Open Destination"
    $btnOpenDest.Add_Click({
        if (Test-Path -LiteralPath $txtDest.Text) {
            Start-Process explorer.exe -ArgumentList $txtDest.Text
        } else {
            [System.Windows.Forms.MessageBox]::Show("Destination does not exist yet.","Folder Not Found",
                [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })
    $form.Controls.Add($btnOpenDest)

    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Location = New-Object System.Drawing.Point(370, 125)
    $btnClear.Size = New-Object System.Drawing.Size(140, 35)
    $btnClear.Text = "Clear Log"
    $btnClear.Add_Click({ $script:LogTextBox.Clear() })
    $form.Controls.Add($btnClear)

    # Log
    $lblLog = New-Object System.Windows.Forms.Label
    $lblLog.Location = New-Object System.Drawing.Point(20, 175)
    $lblLog.Size = New-Object System.Drawing.Size(100, 20)
    $lblLog.Text = "Activity Log:"
    $form.Controls.Add($lblLog)

    $script:LogTextBox = New-Object System.Windows.Forms.TextBox
    $script:LogTextBox.Location = New-Object System.Drawing.Point(20, 198)
    $script:LogTextBox.Size = New-Object System.Drawing.Size(660, 290)
    $script:LogTextBox.Multiline = $true
    $script:LogTextBox.ScrollBars = "Vertical"
    $script:LogTextBox.ReadOnly = $true
    $script:LogTextBox.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $script:LogTextBox.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $script:LogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $form.Controls.Add($script:LogTextBox)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(20, 498)
    $lblStatus.Size = New-Object System.Drawing.Size(660, 25)
    $lblStatus.Text = "Ready. Source can be a folder (e.g. RevitFamily.Biz, Downloads) or a single archive."
    $form.Controls.Add($lblStatus)

    # Single click handler
    $btnInstall.Add_Click({
        $src = $txtSource.Text.Trim()
        $dst = $txtDest.Text.Trim()

        if ([string]::IsNullOrEmpty($src)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify a source.","Validation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning); return
        }
        if (-not (Test-Path -LiteralPath $src)) {
            [System.Windows.Forms.MessageBox]::Show("Source does not exist:`n$src","Validation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning); return
        }
        if ([string]::IsNullOrEmpty($dst)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify a destination.","Validation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning); return
        }

        $btnInstall.Enabled = $false
        $btnInstall.Text = "Installing..."
        $lblStatus.Text = "Install in progress..."

        Write-Log "=========================================="
        Write-Log "Install starting"
        Write-Log "Source: $src"
        Write-Log "Destination: $dst"
        Write-Log "=========================================="

        # Offer 7-Zip install only if archives are likely involved
        $needs7z = $false
        if (Test-Path -LiteralPath $src -PathType Leaf) {
            $needs7z = $script:ArchiveExtensions -contains [System.IO.Path]::GetExtension($src).ToLower()
        } else {
            $needs7z = [bool](Get-ChildItem -LiteralPath $src -File -Recurse -ErrorAction SilentlyContinue |
                              Where-Object { $script:ArchiveExtensions -contains $_.Extension.ToLower() } |
                              Select-Object -First 1)
        }
        if ($needs7z) { Ensure-7Zip -PromptIfMissing | Out-Null }

        $tempRoot = Join-Path $env:TEMP ("MaterialsInstaller_" + [System.IO.Path]::GetRandomFileName())
        New-Item -Path $tempRoot -ItemType Directory -Force | Out-Null
        Write-Log "Temp workspace: $tempRoot"

        try {
            $images = Get-SourceImages -SourcePath $src -TempRoot $tempRoot -Depth 0
            Write-Log "Total images found: $($images.Count)"
            $results = Copy-ImageList -ImageList $images -Destination $dst

            Write-Log "=========================================="
            Write-Log "Install complete!"
            Write-Log "  Files found: $($results.TotalFound)"
            Write-Log "  Files copied: $($results.Copied)"
            Write-Log "  Files skipped (duplicates): $($results.Skipped)"
            Write-Log "  Errors: $($results.Errors)"
            Write-Log "=========================================="
            $lblStatus.Text = "Complete. Copied: $($results.Copied) | Skipped: $($results.Skipped) | Errors: $($results.Errors)"
            $msg = "Install complete!`n`nFound: $($results.TotalFound)`nCopied: $($results.Copied)`nSkipped: $($results.Skipped)`nErrors: $($results.Errors)"
            [System.Windows.Forms.MessageBox]::Show($msg,"Install Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
        } finally {
            try {
                Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
                Write-Log "Cleaned up temp workspace."
            } catch {
                Write-Log "Could not fully clean temp workspace: $_" -Level "WARN"
            }
            $btnInstall.Text = "Install Materials"
            $btnInstall.Enabled = $true
        }
    })

    Write-Log "GSADUs Materials Installer started"
    Write-Log "Default source: $script:DefaultSourcePath"
    Write-Log "Default destination: $script:DefaultDestPath"
    $sevenZipAtStart = Find-7Zip
    if ($sevenZipAtStart) { Write-Log "7-Zip detected: $sevenZipAtStart" } else { Write-Log "7-Zip not detected (will offer to install on first run that needs it)." }

    [void]$form.ShowDialog()
}

Show-GUI
