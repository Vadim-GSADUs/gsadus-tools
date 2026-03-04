@echo off
:: GSADUs TypeCatalog Converter Launcher
:: This batch file launches the PowerShell script with the correct execution policy

title GSADUs TypeCatalog Converter

:: Check for PowerShell
where powershell >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo PowerShell is required but not found.
    echo Please install PowerShell and try again.
    pause
    exit /b 1
)

:: Get the directory where this batch file is located
set "SCRIPT_DIR=%~dp0"

:: Launch the PowerShell script
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%SCRIPT_DIR%Convert-TypeCatalog.ps1"

exit /b 0
