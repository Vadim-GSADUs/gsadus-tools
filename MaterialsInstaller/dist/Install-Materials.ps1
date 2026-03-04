#Requires -Version 5.1
<#
.SYNOPSIS
    Materials Installer - Copies material texture files from Revit family folders to Autodesk shared location.

.DESCRIPTION
    This tool searches a root folder for subfolders named "Materials" and copies all .jpg files
    to the Autodesk shared materials textures folder, skipping duplicates.

.NOTES
    Author: GSADUs
    Created: 2026-01-28
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Default paths
$script:DefaultSourcePath = "G:\Shared drives\GSDE Projects\CADD\RevitFamily.Biz"
$script:DefaultDestPath = "C:\Program Files (x86)\Common Files\Autodesk Shared\Materials\Textures\1\Mats"

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

function Find-MaterialsFolders {
    param([string]$RootPath)
    
    Write-Log "Searching for 'Materials' folders in: $RootPath"
    
    $materialsFolders = @()
    
    try {
        $materialsFolders = Get-ChildItem -Path $RootPath -Directory -Recurse -Filter "Materials" -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -eq "Materials" }
        
        Write-Log "Found $($materialsFolders.Count) Materials folder(s)"
    }
    catch {
        Write-Log "Error searching folders: $_" -Level "ERROR"
    }
    
    return $materialsFolders
}

function Copy-MaterialFiles {
    param(
        [string]$SourceRoot,
        [string]$Destination
    )
    
    $results = @{
        Copied = 0
        Skipped = 0
        Errors = 0
        TotalFound = 0
    }
    
    # Create destination folder if it doesn't exist
    if (-not (Test-Path -Path $Destination)) {
        Write-Log "Creating destination folder: $Destination"
        try {
            New-Item -Path $Destination -ItemType Directory -Force | Out-Null
            Write-Log "Destination folder created successfully"
        }
        catch {
            Write-Log "Failed to create destination folder: $_" -Level "ERROR"
            Write-Log "Try running as Administrator" -Level "ERROR"
            return $results
        }
    }
    
    # Find all Materials folders
    $materialsFolders = Find-MaterialsFolders -RootPath $SourceRoot
    
    if ($materialsFolders.Count -eq 0) {
        Write-Log "No Materials folders found in the source path" -Level "WARN"
        return $results
    }
    
    # Get existing files in destination for duplicate detection
    $existingFiles = @{}
    Get-ChildItem -Path $Destination -Filter "*.jpg" -File -ErrorAction SilentlyContinue | ForEach-Object {
        $existingFiles[$_.Name.ToLower()] = $_.FullName
    }
    Write-Log "Found $($existingFiles.Count) existing .jpg files in destination"
    
    # Process each Materials folder
    foreach ($folder in $materialsFolders) {
        Write-Log "Processing: $($folder.FullName)"
        
        $jpgFiles = Get-ChildItem -Path $folder.FullName -Filter "*.jpg" -File -ErrorAction SilentlyContinue
        
        foreach ($file in $jpgFiles) {
            $results.TotalFound++
            $destFilePath = Join-Path -Path $Destination -ChildPath $file.Name
            $fileNameLower = $file.Name.ToLower()
            
            # Check if file already exists
            if ($existingFiles.ContainsKey($fileNameLower)) {
                Write-Log "  Skipped (duplicate): $($file.Name)" -Level "SKIP"
                $results.Skipped++
                continue
            }
            
            # Copy the file
            try {
                Copy-Item -Path $file.FullName -Destination $destFilePath -Force
                Write-Log "  Copied: $($file.Name)"
                $results.Copied++
                $existingFiles[$fileNameLower] = $destFilePath
            }
            catch {
                Write-Log "  Error copying $($file.Name): $_" -Level "ERROR"
                $results.Errors++
            }
        }
    }
    
    return $results
}

function Show-GUI {
    # Create the main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "GSADUs Materials Installer"
    $form.Size = New-Object System.Drawing.Size(700, 550)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Source Path Label
    $lblSource = New-Object System.Windows.Forms.Label
    $lblSource.Location = New-Object System.Drawing.Point(20, 20)
    $lblSource.Size = New-Object System.Drawing.Size(100, 23)
    $lblSource.Text = "Source Folder:"
    $form.Controls.Add($lblSource)
    
    # Source Path TextBox
    $txtSource = New-Object System.Windows.Forms.TextBox
    $txtSource.Location = New-Object System.Drawing.Point(120, 18)
    $txtSource.Size = New-Object System.Drawing.Size(450, 23)
    $txtSource.Text = $script:DefaultSourcePath
    $form.Controls.Add($txtSource)
    
    # Source Browse Button
    $btnBrowseSource = New-Object System.Windows.Forms.Button
    $btnBrowseSource.Location = New-Object System.Drawing.Point(580, 17)
    $btnBrowseSource.Size = New-Object System.Drawing.Size(80, 25)
    $btnBrowseSource.Text = "Browse..."
    $btnBrowseSource.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select the root folder containing Revit families"
        $folderBrowser.SelectedPath = $txtSource.Text
        if ($folderBrowser.ShowDialog() -eq "OK") {
            $txtSource.Text = $folderBrowser.SelectedPath
        }
    })
    $form.Controls.Add($btnBrowseSource)
    
    # Destination Path Label
    $lblDest = New-Object System.Windows.Forms.Label
    $lblDest.Location = New-Object System.Drawing.Point(20, 55)
    $lblDest.Size = New-Object System.Drawing.Size(100, 23)
    $lblDest.Text = "Destination:"
    $form.Controls.Add($lblDest)
    
    # Destination Path TextBox
    $txtDest = New-Object System.Windows.Forms.TextBox
    $txtDest.Location = New-Object System.Drawing.Point(120, 53)
    $txtDest.Size = New-Object System.Drawing.Size(450, 23)
    $txtDest.Text = $script:DefaultDestPath
    $form.Controls.Add($txtDest)
    
    # Destination Browse Button
    $btnBrowseDest = New-Object System.Windows.Forms.Button
    $btnBrowseDest.Location = New-Object System.Drawing.Point(580, 52)
    $btnBrowseDest.Size = New-Object System.Drawing.Size(80, 25)
    $btnBrowseDest.Text = "Browse..."
    $btnBrowseDest.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select the destination folder for materials"
        $folderBrowser.SelectedPath = $txtDest.Text
        if ($folderBrowser.ShowDialog() -eq "OK") {
            $txtDest.Text = $folderBrowser.SelectedPath
        }
    })
    $form.Controls.Add($btnBrowseDest)
    
    # Install Button
    $btnInstall = New-Object System.Windows.Forms.Button
    $btnInstall.Location = New-Object System.Drawing.Point(20, 95)
    $btnInstall.Size = New-Object System.Drawing.Size(150, 35)
    $btnInstall.Text = "Install Materials"
    $btnInstall.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnInstall.ForeColor = [System.Drawing.Color]::White
    $btnInstall.FlatStyle = "Flat"
    $form.Controls.Add($btnInstall)
    
    # Open Destination Button
    $btnOpenDest = New-Object System.Windows.Forms.Button
    $btnOpenDest.Location = New-Object System.Drawing.Point(180, 95)
    $btnOpenDest.Size = New-Object System.Drawing.Size(150, 35)
    $btnOpenDest.Text = "Open Destination"
    $btnOpenDest.Add_Click({
        if (Test-Path -Path $txtDest.Text) {
            Start-Process explorer.exe -ArgumentList $txtDest.Text
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "Destination folder does not exist yet.",
                "Folder Not Found",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    })
    $form.Controls.Add($btnOpenDest)
    
    # Clear Log Button
    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Location = New-Object System.Drawing.Point(340, 95)
    $btnClear.Size = New-Object System.Drawing.Size(100, 35)
    $btnClear.Text = "Clear Log"
    $btnClear.Add_Click({
        $script:LogTextBox.Clear()
    })
    $form.Controls.Add($btnClear)
    
    # Log Label
    $lblLog = New-Object System.Windows.Forms.Label
    $lblLog.Location = New-Object System.Drawing.Point(20, 145)
    $lblLog.Size = New-Object System.Drawing.Size(100, 20)
    $lblLog.Text = "Activity Log:"
    $form.Controls.Add($lblLog)
    
    # Log TextBox
    $script:LogTextBox = New-Object System.Windows.Forms.TextBox
    $script:LogTextBox.Location = New-Object System.Drawing.Point(20, 168)
    $script:LogTextBox.Size = New-Object System.Drawing.Size(640, 290)
    $script:LogTextBox.Multiline = $true
    $script:LogTextBox.ScrollBars = "Vertical"
    $script:LogTextBox.ReadOnly = $true
    $script:LogTextBox.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $script:LogTextBox.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $script:LogTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $form.Controls.Add($script:LogTextBox)
    
    # Status Label
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(20, 470)
    $lblStatus.Size = New-Object System.Drawing.Size(640, 25)
    $lblStatus.Text = "Ready. Click 'Install Materials' to begin."
    $form.Controls.Add($lblStatus)
    
    # Install Button Click Handler
    $btnInstall.Add_Click({
        $sourcePath = $txtSource.Text.Trim()
        $destPath = $txtDest.Text.Trim()
        
        # Validate paths
        if ([string]::IsNullOrEmpty($sourcePath)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please specify a source folder.",
                "Validation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }
        
        if (-not (Test-Path -Path $sourcePath)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Source folder does not exist: $sourcePath",
                "Validation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }
        
        if ([string]::IsNullOrEmpty($destPath)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please specify a destination folder.",
                "Validation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }
        
        # Disable button during operation
        $btnInstall.Enabled = $false
        $btnInstall.Text = "Installing..."
        $lblStatus.Text = "Installing materials... Please wait."
        
        Write-Log "=========================================="
        Write-Log "Starting Materials Installation"
        Write-Log "Source: $sourcePath"
        Write-Log "Destination: $destPath"
        Write-Log "=========================================="
        
        # Perform the copy operation
        $results = Copy-MaterialFiles -SourceRoot $sourcePath -Destination $destPath
        
        Write-Log "=========================================="
        Write-Log "Installation Complete!"
        Write-Log "  Files found: $($results.TotalFound)"
        Write-Log "  Files copied: $($results.Copied)"
        Write-Log "  Files skipped (duplicates): $($results.Skipped)"
        Write-Log "  Errors: $($results.Errors)"
        Write-Log "=========================================="
        
        $lblStatus.Text = "Complete! Copied: $($results.Copied) | Skipped: $($results.Skipped) | Errors: $($results.Errors)"
        
        # Re-enable button
        $btnInstall.Enabled = $true
        $btnInstall.Text = "Install Materials"
        
        # Show completion message
        $message = "Installation complete!`n`n" +
                   "Files found: $($results.TotalFound)`n" +
                   "Files copied: $($results.Copied)`n" +
                   "Files skipped (duplicates): $($results.Skipped)`n" +
                   "Errors: $($results.Errors)"
        
        [System.Windows.Forms.MessageBox]::Show(
            $message,
            "Installation Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    })
    
    # Show the form
    Write-Log "GSADUs Materials Installer started"
    Write-Log "Default source: $script:DefaultSourcePath"
    Write-Log "Default destination: $script:DefaultDestPath"
    
    [void]$form.ShowDialog()
}

# Auto-elevate to administrator (needed for Program Files) - no prompt, just elevate
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    # Silently restart as Administrator
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Run the GUI
Show-GUI
