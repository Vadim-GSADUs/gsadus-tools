#Requires -Version 5.1
<#
.SYNOPSIS
    Builds the Materials Installer into a standalone executable.

.DESCRIPTION
    This script uses PS2EXE to convert the PowerShell script into a standalone .exe file
    that can be distributed and run on any Windows machine without requiring VS Code.

.NOTES
    Requires PS2EXE module. Will install automatically if not present.
#>

param(
    [switch]$SkipModuleInstall
)

$ErrorActionPreference = "Stop"

$scriptPath = $PSScriptRoot
$sourceScript = Join-Path $scriptPath "Install-Materials.ps1"
$outputExe = Join-Path $scriptPath "dist\MaterialsInstaller.exe"
$distFolder = Join-Path $scriptPath "dist"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " GSADUs Materials Installer Builder" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if PS2EXE is installed
$ps2exeModule = Get-Module -ListAvailable -Name ps2exe

if (-not $ps2exeModule -and -not $SkipModuleInstall) {
    Write-Host "PS2EXE module not found. Installing..." -ForegroundColor Yellow
    
    try {
        Install-Module -Name ps2exe -Scope CurrentUser -Force -AllowClobber
        Write-Host "PS2EXE module installed successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to install PS2EXE module: $_" -ForegroundColor Red
        Write-Host ""
        Write-Host "You can install it manually with:" -ForegroundColor Yellow
        Write-Host "  Install-Module -Name ps2exe -Scope CurrentUser -Force" -ForegroundColor White
        exit 1
    }
}

# Create dist folder if it doesn't exist
if (-not (Test-Path $distFolder)) {
    New-Item -Path $distFolder -ItemType Directory -Force | Out-Null
    Write-Host "Created dist folder: $distFolder" -ForegroundColor Gray
}

# Verify source script exists
if (-not (Test-Path $sourceScript)) {
    Write-Host "Source script not found: $sourceScript" -ForegroundColor Red
    exit 1
}

Write-Host "Building executable..." -ForegroundColor Cyan
Write-Host "  Source: $sourceScript" -ForegroundColor Gray
Write-Host "  Output: $outputExe" -ForegroundColor Gray
Write-Host ""

try {
    # Import the module
    Import-Module ps2exe -Force
    
    # Convert to executable
    Invoke-PS2EXE -InputFile $sourceScript `
                  -OutputFile $outputExe `
                  -Title "GSADUs Materials Installer" `
                  -Description "Installs Revit material textures to Autodesk shared folder" `
                  -Company "GSADUs" `
                  -Product "Materials Installer" `
                  -Version "1.0.0.0" `
                  -Copyright "(c) 2026 GSADUs" `
                  -RequireAdmin `
                  -NoConsole `
                  -NoError
    
    if (Test-Path $outputExe) {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Green
        Write-Host " Build Successful!" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "Executable created at:" -ForegroundColor White
        Write-Host "  $outputExe" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "You can distribute this file to other machines." -ForegroundColor Gray
        Write-Host "It will run without requiring PowerShell scripts or VS Code." -ForegroundColor Gray
        
        # Copy the batch file to dist as well (as a fallback option)
        $batchFile = Join-Path $scriptPath "MaterialsInstaller.bat"
        $scriptCopy = Join-Path $distFolder "Install-Materials.ps1"
        $batchCopy = Join-Path $distFolder "MaterialsInstaller.bat"
        
        Copy-Item -Path $sourceScript -Destination $scriptCopy -Force
        Copy-Item -Path $batchFile -Destination $batchCopy -Force
        
        Write-Host ""
        Write-Host "Also copied script and batch files to dist folder as fallback." -ForegroundColor Gray
    }
    else {
        Write-Host "Build failed - output file not created." -ForegroundColor Red
        exit 1
    }
}
catch {
    Write-Host "Build failed: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Alternative: Use the batch file (MaterialsInstaller.bat) directly." -ForegroundColor Yellow
    Write-Host "It will work on any Windows machine with PowerShell installed." -ForegroundColor Yellow
    exit 1
}
