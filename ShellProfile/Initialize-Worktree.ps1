# Shared by Git, Claude and Codex. Never copies env files or shares node_modules.
[CmdletBinding()]
param([string]$Path = (Get-Location).Path, [switch]$FromHook)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'profile.ps1')
try {
    $context = Get-GSADUsWorktreeContext -Path $Path
    if ($context.Root -eq $context.Main) {
        if ($FromHook) { exit 0 }
        throw 'Initialize-Worktree is for linked worktrees, not the main checkout.'
    }
    # Serialize setup in this worktree; the file is outside the tracked tree.
    $gitDir = git -C $context.Root rev-parse --absolute-git-dir
    $lock = [IO.File]::Open((Join-Path $gitDir 'gsadus-setup.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
    try {
        pull-env -RepoPath $context.Root -MissingOnly
        $lockfile = Join-Path $context.Root 'package-lock.json'
        $modules = Join-Path $context.Root 'node_modules'
        $pending = Join-Path $gitDir 'gsadus-npm-pending'
        if ((Test-Path -LiteralPath $lockfile) -and
            ((Test-Path -LiteralPath $pending) -or -not (Test-Path -LiteralPath $modules))) {
            if ((Test-Path -LiteralPath $modules) -and
                ((Get-Item -LiteralPath $modules).Attributes -band [IO.FileAttributes]::ReparsePoint)) {
                throw 'Refusing npm ci through a node_modules junction or symlink.'
            }
            git -C $context.Root check-ignore -q -- $modules
            if ($LASTEXITCODE -ne 0) { throw 'node_modules must be gitignored.' }
            # Profile restores the existing managed npm auth without rendering other repos.
            Push-Location $context.Root
            try {
                Set-Content -LiteralPath $pending -Value 'npm ci incomplete; retry required'
                Write-Host 'Installing worktree dependencies from package-lock.json...'
                $installOutput = & npm.cmd ci --no-audit --no-fund 2>&1
                if ($LASTEXITCODE -ne 0) {
                    # Do not echo npm/lifecycle output: it can include environment values.
                    throw 'npm ci failed. Check registry authentication and lockfile compatibility; rerun init-worktree after resolving.'
                }
                Remove-Item -LiteralPath $pending
            } finally { Pop-Location }
        }
        Write-Host "Worktree ready: $($context.Root)"
    } finally { $lock.Dispose() }
} catch {
    Write-Error $_.Exception.Message -ErrorAction Continue
    exit 1
}
