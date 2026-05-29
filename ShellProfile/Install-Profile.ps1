# Installs the GSADUs profile shim into this user's $PROFILE.
# The shim simply dot-sources the tracked logic (profile.ps1) from the Tools repo,
# so future changes propagate via git/unwip-all with no per-machine edits.
#
# Idempotent and safe to re-run. Run once per machine after the first clone.
[CmdletBinding()]
param()

$tracked = Join-Path $PSScriptRoot 'profile.ps1'
if (-not (Test-Path $tracked)) { throw "Tracked profile not found at $tracked" }

$profilePath = $PROFILE.CurrentUserCurrentHost
$dir = Split-Path $profilePath -Parent
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Force $dir | Out-Null }

$shim = @"
# GSADUs shell profile shim — managed by Tools\ShellProfile\Install-Profile.ps1
# Real logic is TRACKED in the gsadus-tools repo; edit/commit there, not here.
`$GSADUsProfile = "C:\GSADUs\Tools\ShellProfile\profile.ps1"
if (Test-Path `$GSADUsProfile) { . `$GSADUsProfile }
else { Write-Warning "GSADUs profile not found at `$GSADUsProfile — clone the Tools repo or run Tools\ShellProfile\Install-Profile.ps1" }
"@

if ((Test-Path $profilePath) -and ((Get-Content $profilePath -Raw).Trim() -eq $shim.Trim())) {
    Write-Host "Profile shim already installed at $profilePath" -ForegroundColor DarkGray
    return
}

if (Test-Path $profilePath) {
    $bak = "$profilePath.bak-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    Copy-Item $profilePath $bak -Force
    Write-Host "Backed up existing profile -> $bak" -ForegroundColor Yellow
}

Set-Content -Path $profilePath -Value $shim -Encoding UTF8
Write-Host "Installed GSADUs profile shim -> $profilePath" -ForegroundColor Green
Write-Host "Open a new PowerShell session (or run: . `$PROFILE) to load it." -ForegroundColor Cyan
