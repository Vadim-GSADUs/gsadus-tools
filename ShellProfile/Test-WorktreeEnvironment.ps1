# Offline regression checks. Uses fake secrets and disposable Git repositories only.
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'profile.ps1')
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('gsadus-env-test-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixture | Out-Null
$originalRoot = $GSADUsRoot
$GSADUsRoot = $fixture
$repo = Join-Path $fixture 'WebApp'
$worktree = Join-Path $fixture 'agent checkout'
$script:downloadFails = $false
function doppler {
    if ($args[0] -eq 'me') { $global:LASTEXITCODE = 0; return }
    if ($script:downloadFails) { $global:LASTEXITCODE = 1; return }
    $global:LASTEXITCODE = 0
    'EXAMPLE_VALUE=fake-test-value'
}
function Assert($condition, $message) { if (-not $condition) { throw $message } }
try {
    git init -q $repo
    Set-Content (Join-Path $repo '.gitignore') ".env*`nnode_modules/"
    git -C $repo add .gitignore
    git -C $repo -c user.name=Test -c user.email=test@example.invalid commit -qm fixture
    git -C $repo worktree add -q --detach $worktree
    Assert ($LASTEXITCODE -eq 0) 'Fixture creation failed'
    $context = Get-GSADUsWorktreeContext -Path $worktree
    Assert ($context.Main -eq $repo) 'Wrong main repo resolution'
    pull-env -RepoPath $worktree
    Assert (Test-Path $context.EnvPath) 'Missing rendered env'
    Assert (-not (Test-Path (Join-Path $repo '.env.local'))) 'Main env was changed'
    $before = (Get-FileHash $context.EnvPath).Hash
    $script:downloadFails = $true
    pull-env -RepoPath $worktree -MissingOnly
    $failed = $false
    try { pull-env -RepoPath $worktree } catch { $failed = $true }
    Assert $failed 'Download failure must propagate'
    Assert ((Get-FileHash $context.EnvPath).Hash -eq $before) 'Failure replaced old env'
    $script:downloadFails = $false
    Set-Content (Join-Path $worktree '.gitignore') 'node_modules/'
    $failed = $false
    try { pull-env -RepoPath $worktree } catch { $failed = $true }
    Assert $failed 'Unignored secret was accepted'
    Set-Content (Join-Path $worktree '.gitignore') ".env*`nnode_modules/"
    git -C $worktree add -f .env.local
    $failed = $false
    try { pull-env -RepoPath $worktree } catch { $failed = $true }
    Assert $failed 'Tracked secret was accepted'
    Assert ((Get-FileHash $context.EnvPath).Hash -eq $before) 'Unsafe target was changed'
    Write-Host 'PASS: linked checkout routing, main preservation, missing-only, failure preservation, ignored and tracked destination guards.'
} finally {
    $GSADUsRoot = $originalRoot
    # Only the exact unique fixture created above is removed; no user repo is targeted.
    $resolved = [IO.Path]::GetFullPath($fixture)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($tempRoot) -or (Split-Path $resolved -Leaf) -notlike 'gsadus-env-test-*') {
        throw 'Unsafe fixture cleanup path'
    }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}
