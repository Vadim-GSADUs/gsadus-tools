# GSADUs Materials Installer

A simple tool to install Revit material texture files from your family library to the Autodesk shared materials folder.

## Purpose

When using Revit families from RevitFamily.Biz or similar sources, the render material textures need to be placed in a specific location for them to display correctly:

```
C:\Program Files (x86)\Common Files\Autodesk Shared\Materials\Textures\1\Mats
```

This tool automates the process of:
1. Searching your family library for `Materials` folders
2. Copying all `.jpg` texture files to the Autodesk shared location
3. Skipping any duplicate files that already exist

## Quick Start

### Option 1: Run the Batch File (Simplest)
Double-click `MaterialsInstaller.bat` to launch the tool.

### Option 2: Run the PowerShell Script
Right-click `Install-Materials.ps1` and select "Run with PowerShell"

### Option 3: Use the Standalone Executable
1. Run `Build-Installer.ps1` to create the executable
2. Find `MaterialsInstaller.exe` in the `dist` folder
3. Copy the exe to any Windows machine and run it

## Building the Executable

To create a standalone `.exe` file that can run on any Windows machine:

1. Open PowerShell as Administrator
2. Navigate to this folder
3. Run: `.\Build-Installer.ps1`

The executable will be created in the `dist` folder.

**Note:** The build script will automatically install the `PS2EXE` PowerShell module if needed.

## Default Paths

- **Source**: `G:\Shared drives\GSDE Projects\CADD\RevitFamily.Biz`
- **Destination**: `C:\Program Files (x86)\Common Files\Autodesk Shared\Materials\Textures\1\Mats`

Both paths can be changed in the GUI before running the installation.

## Requirements

- Windows 10 or later
- PowerShell 5.1 or later (included with Windows 10)
- Administrator rights (to write to Program Files)

## Troubleshooting

### "Access Denied" errors
- Run the tool as Administrator
- The tool will prompt you to restart with elevated privileges if needed

### "Source folder does not exist"
- Make sure you have access to the Google Shared Drive
- Verify the path is correct and the drive is mounted

### PS2EXE module installation fails
- Run PowerShell as Administrator
- Try: `Install-Module -Name ps2exe -Scope CurrentUser -Force`
- Or just use the batch file instead of building an exe

## Files

| File | Description |
|------|-------------|
| `Install-Materials.ps1` | Main PowerShell script with GUI |
| `MaterialsInstaller.bat` | Batch file launcher (simple option) |
| `Build-Installer.ps1` | Script to create standalone .exe |
| `dist/` | Output folder for built executable |
