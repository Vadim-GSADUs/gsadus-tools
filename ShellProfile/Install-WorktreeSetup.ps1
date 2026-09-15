# Install a small, repo-local Git hook. No global hooksPath or harness policy changes.
[CmdletBinding()]
param([string[]]$Repos = @('WebApp', 'PM', 'WebCatalog', 'PostProcess\PNGTools'), [switch]$Uninstall)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'profile.ps1')
$marker = '# GSADUs managed worktree setup v1'
$hook = @'
#!/bin/sh
# GSADUs managed worktree setup v1
# Only initial worktree checkout, never branch switches or file checkouts.
[ "$3" = "1" ] || exit 0
case "$1" in *[!0]*) exit 0 ;; esac
exec pwsh -NoLogo -NoProfile -File 'C:/GSADUs/Tools/ShellProfile/Initialize-Worktree.ps1' -FromHook
'@
foreach ($repo in $Repos) {
    $path = Join-Path $GSADUsRoot $repo
    $context = Get-GSADUsWorktreeContext -Path $path
    $configuredHooks = git -C $path config --get core.hooksPath
    if ($configuredHooks) { throw "Existing core.hooksPath for $repo; integrate explicitly instead of replacing it." }
    $hooksDir = Join-Path $path '.git\hooks'
    $target = Join-Path $hooksDir 'post-checkout'
    if (Test-Path -LiteralPath $target) {
        $old = Get-Content -LiteralPath $target -Raw
        if (-not $old.Contains($marker)) { throw "Existing post-checkout hook for $repo; left untouched." }
        if ($Uninstall) { Remove-Item -LiteralPath $target; continue }
        if ($old.Trim() -eq $hook.Trim()) { Write-Host "Already installed: $repo"; continue }
        Copy-Item -LiteralPath $target -Destination "$target.bak-$(Get-Date -Format yyyyMMdd-HHmmss)"
    }
    if (-not $Uninstall) {
        New-Item -ItemType Directory -Force -Path $hooksDir | Out-Null
        [IO.File]::WriteAllText($target, $hook.Replace("`r`n", "`n") + "`n", [Text.UTF8Encoding]::new($false))
        Write-Host "Installed worktree setup: $repo"
    }
}
