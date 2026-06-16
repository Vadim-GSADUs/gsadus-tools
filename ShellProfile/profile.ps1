# GSADUs shell profile — per-machine-branch wip / unwip sync
# ===========================================================================
# This file is TRACKED in the gsadus-tools repo (Tools\ShellProfile\profile.ps1).
# Each machine's $PROFILE is a thin shim that dot-sources this file (see
# Install-Profile.ps1), so logic changes propagate via git/unwip-all — there is
# no per-machine copy to keep in sync. Do NOT scp this around; edit + commit it.
#
# Model (per-machine branches — no force-push race, no silent data loss):
#   * Real work lives on the normal branch (usually 'main'); pushed/pulled
#     normally, never forced.
#   * Each machine snapshots its in-progress working tree to its OWN remote
#     branch:  wip/<hostname>.  Only that machine writes that ref, so a
#     force-push can never clobber another machine's work.
#   * 'unwip' fast-forwards the real branch, then adopts the NEWEST wip/* branch
#     belonging to a DIFFERENT machine as uncommitted changes.
#   * The autostash is only discarded after we confirm its content is already
#     safe on this machine's own wip/<host> ref; otherwise it is kept.

$GSADUsRoot     = "C:\GSADUs"
$GSADUsProfile  = "$GSADUsRoot\Tools\ShellProfile\profile.ps1"   # canonical path of this file

function Get-WipHost {
    # Stable, ref-safe machine id: lowercase computer name, non-alphanumerics -> '-'
    (($env:COMPUTERNAME).ToLower() -replace '[^a-z0-9]+', '-').Trim('-')
}

function Get-WipRepos {
    # Only include repos whose origin points at Vadim-GSADUs (skips third-party forks)
    $candidates = @()
    if (Test-Path "$GSADUsRoot\.git") { $candidates += $GSADUsRoot }
    Get-ChildItem $GSADUsRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        if (Test-Path "$($_.FullName)\.git") { $candidates += $_.FullName }
        Get-ChildItem $_.FullName -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            if (Test-Path "$($_.FullName)\.git") { $candidates += $_.FullName }
        }
    }
    $repos = @()
    foreach ($c in $candidates) {
        Push-Location $c
        $url = git remote get-url origin 2>$null
        if ($url -match "Vadim-GSADUs") { $repos += $c }
        Pop-Location
    }
    $repos
}

# -- cross-machine sync report (collected per run, shown loudly at the end) ----
# A single missed WARN line in a wall of per-repo output could let a repo's work
# silently stay behind. Every repo that can't be saved/synced cleanly is recorded
# here and reported as a summary so it can't be overlooked before switching PCs.
$global:GSADUsWipReport = @()
function Reset-WipReport   { $global:GSADUsWipReport = @() }
function Add-WipAttention  { param([string]$Msg) $global:GSADUsWipReport += $Msg }
function Show-WipReport {
    param([string]$Verb = "saved")
    if ($global:GSADUsWipReport.Count -gt 0) {
        Write-Host ""
        Write-Host ("  !! {0} repo(s) NOT {1} — need attention before switching machines:" -f $global:GSADUsWipReport.Count, $Verb) -ForegroundColor Red
        foreach ($m in $global:GSADUsWipReport) { Write-Host "       - $m" -ForegroundColor Red }
        Write-Host "    Run 'unwip-all' here, resolve, then re-run." -ForegroundColor DarkGray
    }
}

# -- core per-repo helpers (operate on the current directory) ------------------

function Save-RepoWip {
    param([string]$Label = ".")
    $me     = Get-WipHost
    $branch = git rev-parse --abbrev-ref HEAD 2>$null
    if (-not $branch) { Write-Host "  skip $Label (not a git repo)" -ForegroundColor DarkGray; return }

    git fetch -q --prune origin 2>$null

    # 1. Push any real commits first (never force). If this fails the remote has
    #    real commits we don't have -> unwip first instead of clobbering.
    $pushedReal = $false
    if (git log "origin/$branch..HEAD" --oneline 2>$null) {
        git push -q origin $branch 2>$null
        if ($LASTEXITCODE -ne 0) {
            Add-WipAttention "$Label — local commits on '$branch' diverge from remote; run unwip-all and reconcile"
            Write-Host "  WARN $Label — '$branch' diverged from remote; left unchanged (see summary)" -ForegroundColor Red
            return
        }
        $pushedReal = $true
    }

    # 1b. If the real branch is behind remote, AUTO-RESOLVE before snapshotting so
    #     we never snapshot on a STALE base (the cause of phantom unwip conflicts).
    #     "Behind" = real commits landed on origin we haven't pulled. We do the
    #     *pull half* of unwip — fast-forward 'main' and replay our uncommitted
    #     edits on top — but deliberately NOT the wip-adoption half (that belongs
    #     to unwip on arrival, not to a save). Clean replay -> continue and snapshot
    #     a fresh, current base. If our edits genuinely clash with the incoming
    #     commits, roll the repo back to its EXACT original state and flag it —
    #     never hand the user a half-merged tree at save time.
    $behindRaw = git rev-list "HEAD..origin/$branch" --count 2>$null
    if ($behindRaw -and [int]$behindRaw -gt 0) {
        $orig  = git rev-parse HEAD 2>$null
        $dirty = [bool](git status --porcelain)
        if ($dirty) { git stash push -u -q -m "wip-autoresolve" 2>$null }

        git merge --ff-only -q "origin/$branch" 2>$null
        if ($LASTEXITCODE -ne 0) {
            if ($dirty) { git stash pop -q 2>$null }   # ff impossible (local commits) — undo stash, bail
            Add-WipAttention "$Label — '$branch' diverged from remote; run unwip-all and reconcile"
            Write-Host "  WARN $Label — '$branch' diverged; left unchanged (see summary)" -ForegroundColor Red
            return
        }

        if ($dirty) {
            git stash pop 2>&1 | Out-Null              # replay our edits onto current 'main'
            if ($LASTEXITCODE -ne 0) {
                # Genuine overlap between our edits and the incoming commits. Restore
                # the repo to exactly where it started: reset 'main' back, then reapply
                # the still-present stash onto its own base (always clean), and flag it.
                git reset -q --hard $orig 2>$null
                git stash pop -q 2>$null
                Add-WipAttention "$Label — your edits clash with incoming commits; run unwip-all here and resolve"
                Write-Host "  WARN $Label — could not auto-resolve; left unchanged (see summary)" -ForegroundColor Red
                return
            }
        }
        Write-Host "  sync $Label (caught up $behindRaw commit(s) on '$branch')" -ForegroundColor DarkGray
    }

    # 2. Snapshot the working tree (tracked + untracked) onto wip/<me>.
    git add -A
    if (-not (git status --porcelain)) {
        git reset -q
        if ($pushedReal) { Write-Host "  push $Label (commits)" -ForegroundColor Yellow }
        else { Write-Host "  skip $Label (clean)" -ForegroundColor DarkGray }
        return
    }
    git commit --no-verify --no-gpg-sign -q -m "wip: $me $(Get-Date -Format 'yyyyMMdd-HHmm')"
    git push -q --force origin "HEAD:refs/heads/wip/$me"
    $pushOk = $LASTEXITCODE -eq 0
    # 3. Restore the dirty working tree; the real branch is left untouched.
    git reset -q --mixed HEAD~1
    if ($pushOk) { Write-Host "  wip  $Label -> wip/$me" -ForegroundColor Cyan }
    else { Write-Host "  ERROR $Label push failed (work is safe locally)" -ForegroundColor Red }
}

function Restore-RepoWip {
    param([string]$Label = ".")
    $me     = Get-WipHost
    $branch = git rev-parse --abbrev-ref HEAD 2>$null
    if (-not $branch) { Write-Host "  skip $Label (not a git repo)" -ForegroundColor DarkGray; return }

    git fetch -q --prune origin 2>$null

    # Newest wip branch belonging to ANOTHER machine.
    $other = git for-each-ref --sort=-committerdate --format="%(refname:short)" "refs/remotes/origin/wip/*" 2>$null |
        Where-Object { $_ -and $_ -ne "origin/wip/$me" } | Select-Object -First 1
    $otherSha = if ($other) { git rev-parse $other 2>$null } else { $null }
    $last     = git config --local --get wip.lastApplied 2>$null
    $incoming = $other -and ($otherSha -ne $last)

    $behindRaw = git rev-list "HEAD..origin/$branch" --count 2>$null
    $behind    = if ($behindRaw) { [int]$behindRaw } else { 0 }

    if (-not $incoming -and $behind -eq 0) {
        Write-Host "  skip $Label (up to date)" -ForegroundColor DarkGray
        return
    }

    # Protect local state by stashing to a clean tree. Decide first whether that
    # local state is already safely captured on our own wip/<me> ref.
    $stashed  = $false
    $unsynced = $false
    if (git status --porcelain) {
        $haveMyWip = git rev-parse --verify --quiet "origin/wip/$me" 2>$null
        if ($haveMyWip) {
            git add -A
            $delta = git diff --cached "origin/wip/$me" 2>$null
            git reset -q
            if ($delta) { $unsynced = $true }   # local work not yet on my wip branch
        } else {
            $unsynced = $true
        }
        git stash push -u -q -m "unwip-autostash" 2>$null
        $stashed = $true
    }

    # 1. Fast-forward the real branch.
    if ($behind -gt 0) {
        git merge --ff-only -q "origin/$branch" 2>$null
        if ($LASTEXITCODE -ne 0) {
            Add-WipAttention "$Label — '$branch' diverged from remote; resolve manually"
            Write-Host "  WARN $Label — '$branch' diverged from remote; resolve manually" -ForegroundColor Red
            if ($stashed) { git stash pop -q 2>$null }
            return
        }
    }

    # 2. Adopt the incoming snapshot as uncommitted changes (onto a clean base).
    if ($incoming) {
        git cherry-pick --no-commit $other 2>$null
        if ($LASTEXITCODE -ne 0) {
            # A conflicted '--no-commit' cherry-pick leaves an unmerged index and
            # conflict markers in the worktree; 'cherry-pick --abort' clears the
            # sequencer flag but NOT the half-merged tree. Hard-reset to the
            # (already fast-forwarded) branch tip so 'left unchanged' is literally
            # true and markers can't pile up across runs. The user's own work is
            # safe in the autostash below — never touched by this reset.
            git cherry-pick --abort 2>$null
            git reset -q --hard HEAD 2>$null
            Add-WipAttention "$Label — incoming wip conflicts with '$branch'; resolve manually"
            Write-Host "  WARN $Label — incoming wip conflicts with '$branch'; left unchanged" -ForegroundColor Red
            if ($stashed) { Write-Host "       your local changes are saved in 'git stash'." -ForegroundColor DarkGray }
            return
        }
        git reset -q
        git config --local wip.lastApplied $otherSha
    }

    # 3. Resolve the autostash without ever losing work.
    if ($stashed) {
        if ($incoming) {
            if ($unsynced) {
                Write-Host "  unwip $Label (your unsynced edits are kept in 'git stash')" -ForegroundColor Yellow
            } else {
                git stash drop -q 2>$null   # content already safe on origin/wip/$me
                Write-Host "  unwip $Label <- $other" -ForegroundColor Cyan
            }
        } else {
            git stash pop 2>&1 | Out-Null   # reapply local edits on top of updated branch
            if ($LASTEXITCODE -ne 0) {
                Add-WipAttention "$Label — local edits conflict after update; kept in 'git stash'"
                Write-Host "  WARN $Label — local edits conflict after update; kept in 'git stash'" -ForegroundColor Red
            } else {
                Write-Host "  pull $Label (commits + local edits restored)" -ForegroundColor Cyan
            }
        }
    } elseif ($incoming) {
        Write-Host "  unwip $Label <- $other" -ForegroundColor Cyan
    } else {
        Write-Host "  pull $Label (commits)" -ForegroundColor Cyan
    }
}

# -- single-repo commands (run from inside a repo) ----------------------------

function wip   { Reset-WipReport; Save-RepoWip    -Label (Split-Path -Leaf (Get-Location)); Show-WipReport }
function unwip { Reset-WipReport; Restore-RepoWip -Label (Split-Path -Leaf (Get-Location)); Show-WipReport -Verb "synced" }

# -- all-repo commands (run from anywhere) ------------------------------------

function wip-all {
    Reset-WipReport
    foreach ($repo in Get-WipRepos) {
        Push-Location $repo
        $rel = $repo.Replace($GSADUsRoot, "").TrimStart("\")
        if (-not $rel) { $rel = "." }
        Save-RepoWip -Label $rel
        Pop-Location
    }
    Show-WipReport
}

function unwip-all {
    Reset-WipReport
    # Sync the workspace root first so setup.ps1 is current before cloning missing repos.
    Push-Location $GSADUsRoot
    Restore-RepoWip -Label "."
    Pop-Location

    # Clone any repos listed in setup.ps1 that aren't present locally.
    & "$GSADUsRoot\setup.ps1" -NonInteractive -CloneOnly

    foreach ($repo in Get-WipRepos) {
        if ($repo -eq $GSADUsRoot) { continue }
        Push-Location $repo
        $rel = $repo.Replace($GSADUsRoot, "").TrimStart("\")
        if (-not $rel) { $rel = "." }
        Restore-RepoWip -Label $rel
        Pop-Location
    }
    Show-WipReport -Verb "synced"
}

function end-day {
    Write-Host ""
    Write-Host "Saving work across all repos..." -ForegroundColor Cyan
    wip-all
    Write-Host ""
    Write-Host "Locking screen." -ForegroundColor Yellow
    rundll32.exe user32.dll,LockWorkStation
}

# -- startup task helpers -----------------------------------------------------
function Register-StartupUnwip {
    # The scheduled task dot-sources the tracked profile directly (not $PROFILE),
    # so it works even if the shim is ever missing.
    $scriptPath = $GSADUsProfile
    # 60s delay lets network, SSH agent, and OneDrive settle after logon.
    # Window stays open on failure so the user sees something went wrong.
    $cmd = @"
Start-Sleep -Seconds 60; `
. '$scriptPath'; `
unwip-all *>&1 | Tee-Object -Variable _out; `
if (`$_out -match 'ERROR|WARN') { `
  Write-Host ''; `
  Write-Host '  unwip-all finished with errors — review above and press Enter to close.' -ForegroundColor Red; `
  Read-Host `
} else { `
  Write-Host ''; `
  Write-Host '  All repos synced.' -ForegroundColor Green; `
  Start-Sleep -Seconds 3 `
}
"@
    $action   = New-ScheduledTaskAction -Execute "pwsh" `
        -Argument "-NoProfile -Command $cmd" `
        -WorkingDirectory $GSADUsRoot
    $trigger  = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
    $settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
        -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
    Register-ScheduledTask -TaskName "GSADUs-unwip-all" -Action $action `
        -Trigger $trigger -Settings $settings -RunLevel Highest -Force | Out-Null
    Write-Host "Startup unwip registered (60s delay, window stays open on failure)." -ForegroundColor Green
}

function Unregister-StartupUnwip {
    Unregister-ScheduledTask -TaskName "GSADUs-unwip-all" -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Startup unwip removed." -ForegroundColor Yellow
}
