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
#   * A kept autostash self-prunes on the next unwip: once its content is
#     reproducible from the working tree or committed history it is dropped, so
#     redundant stashes can never pile up across runs (see Remove-RedundantWipStash).

$GSADUsRoot     = "C:\GSADUs"
$GSADUsProfile  = "$GSADUsRoot\Tools\ShellProfile\profile.ps1"   # canonical path of this file
$GSADUsWipConflictReport = "$GSADUsRoot\.wip-conflict.md"

function Get-WipHost {
    # Stable, ref-safe machine id: lowercase computer name, non-alphanumerics -> '-'
    (($env:COMPUTERNAME).ToLower() -replace '[^a-z0-9]+', '-').Trim('-')
}

# Retired repos kept on disk as read-only reference. Their GitHub remotes are either
# archived (pushes rejected) or deleted outright (every remote op fails), so wip/unwip
# must skip them. Paths relative to $GSADUsRoot.
$GSADUsRetiredRepos = @(
    'PostProcess\DigitalDarkroom'   # archived 2026-07-07; superseded by PNGTools darkroom
    'PostProcess\Darkroom'          # archived 2026-07-07; stalled web console, PNGTools outgrew it
    'SiteCheck'                     # retired 2026-08-06, GitHub repo DELETED 2026-08-11 (no remote); module ships from WebApp
)

function Get-WipRepos {
    # Only include repos whose origin points at Vadim-GSADUs (skips third-party forks)
    $retired = $GSADUsRetiredRepos | ForEach-Object { Join-Path $GSADUsRoot $_ }
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
        if ($retired -contains $c) { continue }
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
function Reset-WipReport   {
    $global:GSADUsWipReport = @()
    $global:GSADUsWipConflictReportStarted = $false
}
function Add-WipAttention  { param([string]$Msg) $global:GSADUsWipReport += $Msg }
function Write-WipConflictReport {
    param(
        [Parameter(Mandatory)][string]$Label,
        [Parameter(Mandatory)][string]$Reason,
        [string]$IncomingRef,
        [string]$IncomingSha,
        [string]$LastApplied,
        [string[]]$ConflictedFiles = @(),
        [string[]]$IncomingFiles = @()
    )

    $repoPath = (Get-Location).Path
    $branch = git rev-parse --abbrev-ref HEAD 2>$null
    $head = git rev-parse HEAD 2>$null
    $status = @(git status --short 2>$null)
    $when = Get-Date -Format "yyyy-MM-dd HH:mm:ss zzz"

    $lines = @()
    $firstReportThisRun = -not $global:GSADUsWipConflictReportStarted
    if ($firstReportThisRun) {
        $lines += "# GSADUs WIP Conflict Report"
        $lines += ""
        $lines += "Generated: $when"
        $lines += ""
        $lines += 'If this file exists, the last `unwip`/`unwip-all` run found a conflict that needs an agent to reconcile.'
        $lines += 'Do not ask the human to manage this. Resolve the repo state, update `wip.lastApplied` when the incoming wip has been accounted for, then rerun `unwip-all`.'
        $lines += ""
        $global:GSADUsWipConflictReportStarted = $true
    }

    $lines += "## $Label"
    $lines += ""
    $lines += ('- Repo: `{0}`' -f $repoPath)
    $lines += ('- Branch: `{0}`' -f $branch)
    $lines += ('- HEAD: `{0}`' -f $head)
    if ($IncomingRef) { $lines += ('- Incoming wip: `{0}`' -f $IncomingRef) }
    if ($IncomingSha) { $lines += ('- Incoming SHA: `{0}`' -f $IncomingSha) }
    if ($LastApplied) { $lines += ('- Previous `wip.lastApplied`: `{0}`' -f $LastApplied) }
    else { $lines += '- Previous `wip.lastApplied`: `(none)`' }
    $lines += "- Reason: $Reason"
    $lines += ""

    if ($ConflictedFiles.Count -gt 0) {
        $lines += "### Conflicted files"
        $lines += ""
        foreach ($f in $ConflictedFiles) { $lines += ('- `{0}`' -f $f) }
        $lines += ""
    }

    if ($IncomingFiles.Count -gt 0) {
        $lines += "### Incoming wip file changes"
        $lines += ""
        $lines += '```text'
        $lines += $IncomingFiles
        $lines += '```'
        $lines += ""
    }

    if ($status.Count -gt 0) {
        $lines += "### Current status"
        $lines += ""
        $lines += '```text'
        $lines += $status
        $lines += '```'
        $lines += ""
    }

    $lines += "### Agent resolution checklist"
    $lines += ""
    $lines += "1. Inspect the repo and incoming wip:"
    $lines += ('   - `git -C "{0}" status --short --branch`' -f $repoPath)
    if ($IncomingRef) { $lines += ('   - `git -C "{0}" diff HEAD..{1}`' -f $repoPath, $IncomingRef) }
    $lines += "2. If the incoming wip content is already represented on the current branch, do not edit files just to satisfy the merge."
    $lines += "3. If real incoming work is missing, apply it intentionally and keep only the root-cause/current path."
    if ($IncomingSha) { $lines += ('4. After the incoming wip has been accounted for, mark it applied: `git -C "{0}" config --local wip.lastApplied {1}`' -f $repoPath, $IncomingSha) }
    else { $lines += '4. After the incoming wip has been accounted for, mark the exact incoming SHA as applied with `git config --local wip.lastApplied <sha>`.' }
    $lines += '5. Rerun `unwip-all`. Resolution is complete only when it reports no warnings and this file is removed automatically.'
    $lines += ""
    $lines += "---"
    $lines += ""

    if ($firstReportThisRun) {
        Set-Content -LiteralPath $GSADUsWipConflictReport -Value $lines -Encoding utf8
    } else {
        Add-Content -LiteralPath $GSADUsWipConflictReport -Value $lines -Encoding utf8
    }
}
function Clear-WipConflictReportIfClean {
    if ($global:GSADUsWipReport.Count -eq 0 -and (Test-Path -LiteralPath $GSADUsWipConflictReport)) {
        Remove-Item -LiteralPath $GSADUsWipConflictReport -Force
    }
}
function Show-WipReport {
    param([string]$Verb = "saved")
    if ($global:GSADUsWipReport.Count -gt 0) {
        Write-Host ""
        Write-Host ("  !! {0} repo(s) NOT {1} — need attention before switching machines:" -f $global:GSADUsWipReport.Count, $Verb) -ForegroundColor Red
        foreach ($m in $global:GSADUsWipReport) { Write-Host "       - $m" -ForegroundColor Red }
        Write-Host "    Run 'unwip-all' here, resolve, then re-run." -ForegroundColor DarkGray
        if (Test-Path -LiteralPath $GSADUsWipConflictReport) {
            Write-Host "    Agent conflict report: $GSADUsWipConflictReport" -ForegroundColor Yellow
            Invoke-WipConflictAgent
        }
    } elseif ($Verb -eq "synced") {
        Clear-WipConflictReportIfClean
    }
}

# The conflict report is written FOR an agent ("Do not ask the human to manage
# this") — this is the hook that actually offers to spawn one. Interactive-only:
# ask first, then open a seeded Claude session in the workspace root. Set
# GSADUS_WIP_AGENT=0 to suppress the offer entirely.
function Invoke-WipConflictAgent {
    if ($env:GSADUS_WIP_AGENT -eq '0') { return }
    if (-not (Get-Command claude -ErrorAction SilentlyContinue)) { return }
    $ans = $null
    try { $ans = Read-Host "    Launch Claude to resolve the conflict now? [Y/n]" }
    catch { return }   # non-interactive host (scheduled/redirected) — skip
    if ($ans -ne '' -and $ans -notmatch '^[Yy]') { return }
    Push-Location $GSADUsRoot
    try {
        # Pinned setup for the conflict-resolution session (decision 2026-07-14):
        # Opus 4.8 (not the account default), acceptEdits so file resolution in the
        # conflicted sub-repos never prompts, and git pre-allowed in both shells —
        # the whole checklist is git surgery, so per-command approval is pure noise.
        claude --model claude-opus-4-8 --permission-mode acceptEdits `
            --allowedTools "Bash(git:*)" "PowerShell(git:*)" "Edit" "Write" `
            ("A GSADUs wip sync conflict was just detected. Read {0} and resolve each repo section by following its 'Agent resolution checklist'. Rules: prefer the root-cause/current path (workspace CLAUDE.md rule 6); if the incoming wip content is already represented on the current branch, do not edit files just to satisfy the merge; after accounting for incoming work set wip.lastApplied to the incoming SHA; finish by rerunning unwip-all and confirming the report file is removed automatically." -f $GSADUsWipConflictReport)
    } finally { Pop-Location }
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

    # Secret guard: `git add -A` stages only non-ignored files, so this lists
    # exactly what would be committed. If a known secret file is no longer
    # ignored (e.g. .gitignore got weakened/synced away), refuse to wip this
    # repo instead of snapshotting the secret to the remote wip/<host> branch.
    $leak = git ls-files --others --exclude-standard 2>$null | Where-Object {
        ($_ -match '(^|/)\.env($|\.)' -and $_ -notmatch '\.env\.example$') -or
        ($_ -match 'serviceaccount.*\.json$') -or
        ($_ -match '(^|/)gcs-.*\.json$')
    }
    if ($leak) {
        Add-WipAttention "$Label — NOT wipped: secret file not ignored ($($leak -join ', ')); fix .gitignore"
        Write-Host "  BLOCK $Label — secret file would be committed; left unchanged (see summary)" -ForegroundColor Red
        return
    }

    # 2. Snapshot the working tree (tracked + untracked) onto wip/<me>.
    git add -A
    if (-not (git status --porcelain)) {
        git reset -q
        if ($pushedReal) { Write-Host "  push $Label (commits)" -ForegroundColor Yellow }
        else { Write-Host "  skip $Label (clean)" -ForegroundColor DarkGray }
        return
    }
    # Re-snapshotting an UNCHANGED dirty tree would mint a new SHA over
    # identical content; the other machine judges "new incoming wip" by SHA, so
    # every such snapshot triggers a phantom re-unwip there (and its own
    # re-snapshot back — a daily ping-pong of the same bytes). If the snapshot
    # tree matches what wip/<me> already holds, keep the existing ref.
    $newTree = git write-tree 2>$null
    $wipTree = git rev-parse --verify --quiet "origin/wip/$me^{tree}" 2>$null
    if ($newTree -and $wipTree -and $newTree -eq $wipTree) {
        git reset -q
        Write-Host "  skip $Label (wip/$me already current)" -ForegroundColor DarkGray
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

function Test-WipStashRedundant {
    # A 'unwip-autostash' entry is redundant when its ENTIRE content is already
    # reproducible from the current repo — tracked changes identical to the
    # working tree AND every captured untracked file present with identical
    # content. In that case the stash is a pure duplicate (the work survives in
    # the tree, or is committed) and can be dropped losslessly. The check is
    # deliberately conservative: any doubt -> not redundant -> the stash is kept.
    param([string]$Ref)
    # Compare tracked content only over the paths the stash itself modified
    # ($Ref^1 is the stash's base commit). A whole-tree `git diff $Ref` also
    # flags files touched by commits landed AFTER the stash was taken, so a
    # stash based on an older parent looked non-redundant forever even when
    # every change it captured was already in HEAD. NUL-delimited so paths
    # with spaces (e.g. "GSADUs Tools.tab") survive intact.
    $rawPaths = git diff --name-only -z "$Ref^1" "$Ref" 2>$null
    if ($LASTEXITCODE -ne 0) { return $false }        # can't enumerate -> keep
    $paths = @((@($rawPaths) -join '') -split "`0" | Where-Object { $_ })
    if ($paths.Count -gt 0) {
        if (git diff $Ref -- $paths 2>$null) { return $false }   # tracked content differs
        if ($LASTEXITCODE -ne 0) { return $false }
    }
    # `git stash -u` records untracked files under the stash's 3rd parent.
    if (git rev-parse --verify --quiet "$Ref^3" 2>$null) {
        foreach ($f in (git ls-tree -r --name-only "$Ref^3" 2>$null)) {
            if (-not (Test-Path -LiteralPath $f)) { return $false }
            if ((git rev-parse "${Ref}^3:$f" 2>$null) -ne (git hash-object -- $f 2>$null)) { return $false }
        }
    }
    return $true
}

function Remove-RedundantWipStash {
    # Prune any pre-existing 'unwip-autostash' entries that are now redundant, so
    # they can't snowball across runs (3+ byte-identical stashes was the original
    # symptom). Re-enumerate after each drop because indices shift; bail as soon
    # as a pass finds nothing left to remove. Only ever drops stashes proven safe
    # by Test-WipStashRedundant — genuine unsynced work is always left untouched.
    while ($true) {
        $entries = @(git stash list 2>$null)
        $dropped = $false
        for ($i = 0; $i -lt $entries.Count; $i++) {
            if ($entries[$i] -notmatch 'unwip-autostash') { continue }
            if (Test-WipStashRedundant "stash@{$i}") {
                git stash drop -q "stash@{$i}" 2>$null
                $dropped = $true
                break
            }
        }
        if (-not $dropped) { break }
    }
}

function Restore-RepoWip {
    param([string]$Label = ".")
    $me     = Get-WipHost
    $branch = git rev-parse --abbrev-ref HEAD 2>$null
    if (-not $branch) { Write-Host "  skip $Label (not a git repo)" -ForegroundColor DarkGray; return }

    git fetch -q --prune origin 2>$null

    # Self-clean: discard leftover autostashes whose content is already safe in
    # the tree/history before doing anything else, so a kept stash never outlives
    # its usefulness and pile up over successive unwip runs.
    Remove-RedundantWipStash

    # Newest wip branch belonging to ANOTHER machine.
    $other = git for-each-ref --sort=-committerdate --format="%(refname:short)" "refs/remotes/origin/wip/*" 2>$null |
        Where-Object { $_ -and $_ -ne "origin/wip/$me" } | Select-Object -First 1
    $otherSha = if ($other) { git rev-parse $other 2>$null } else { $null }
    $last     = git config --local --get wip.lastApplied 2>$null
    $incoming = $other -and ($otherSha -ne $last)

    # A new SHA does not always mean new content: re-snapshots and ping-ponged
    # adoptions mint fresh SHAs over identical trees. If the incoming snapshot's
    # tree matches what this repo already holds (committed + uncommitted), there
    # is nothing to adopt — record it as applied instead of re-cherry-picking.
    if ($incoming) {
        git add -A 2>$null
        $localTree = git write-tree 2>$null
        git reset -q 2>$null
        if ($localTree -and $localTree -eq (git rev-parse "$other^{tree}" 2>$null)) {
            git config --local wip.lastApplied $otherSha
            $incoming = $false
        }
    }

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
            Write-WipConflictReport `
                -Label $Label `
                -Reason "'$branch' diverged from remote during unwip fast-forward" `
                -LastApplied $last
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
            $conflictedFiles = @(git diff --name-only --diff-filter=U 2>$null)
            $incomingFiles = @(git diff --name-status "HEAD..$other" 2>$null)
            Write-WipConflictReport `
                -Label $Label `
                -Reason "incoming wip conflicts with '$branch'" `
                -IncomingRef $other `
                -IncomingSha $otherSha `
                -LastApplied $last `
                -ConflictedFiles $conflictedFiles `
                -IncomingFiles $incomingFiles
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
                # A pop can fail merely because the stashed content already
                # landed via the pulled commits (e.g. an untracked file that is
                # now tracked with identical content). That stash is a pure
                # duplicate — drop it and finish clean instead of warning.
                if (Test-WipStashRedundant 'stash@{0}') {
                    git stash drop -q 2>$null
                    Write-Host "  pull $Label (commits; local edits already included)" -ForegroundColor Cyan
                    return
                }
                $conflictedFiles = @(git diff --name-only --diff-filter=U 2>$null)
                Write-WipConflictReport `
                    -Label $Label `
                    -Reason "local edits conflict after update; autostash kept" `
                    -LastApplied $last `
                    -ConflictedFiles $conflictedFiles
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
    # Arriving at a PC stays one command: after a CLEAN sync, refresh the rendered
    # .env files from Doppler too (see the secret rendering section below). Set
    # GSADUS_UNWIP_PULLENV=0 to skip (same opt-out pattern as GSADUS_ENDDAY_SCAN).
    if ($global:GSADUsWipReport.Count -eq 0 -and $env:GSADUS_UNWIP_PULLENV -ne '0') {
        Write-Host ""
        pull-env
    }
}

function end-day {
    Write-Host ""
    Write-Host "Saving work across all repos..." -ForegroundColor Cyan
    wip-all
    # Light vault scan runs headless in the background (survives the screen lock)
    # and pushes its own vault commit — see Vault\scripts\wiki-scan.ps1.
    $scan = "$GSADUsRoot\Vault\scripts\wiki-scan.ps1"
    if ((Test-Path $scan) -and $env:GSADUS_ENDDAY_SCAN -ne '0') {
        Write-Host ""
        Write-Host "Launching background vault scan (light, today's repos only)..." -ForegroundColor Cyan
        Start-Process pwsh -ArgumentList '-NoProfile','-File',$scan,'-Light' -WindowStyle Hidden
    }
    Write-Host ""
    Write-Host "Locking screen." -ForegroundColor Yellow
    rundll32.exe user32.dll,LockWorkStation
}

# -- secret (.env) rendering from Doppler --------------------------------------
# .env files are gitignored (web-catalog has an external collaborator), so they
# can't ride git like everything else. Since the 2026-08-25 cutover, Doppler is
# the single source of truth for every application secret (Vault wiki:
# secrets-management) and the local files are rendered artifacts:
#
#   pull-env   render each repo's env file from Doppler, then bridge PM's
#              NODE_AUTH_TOKEN into npm auth (see below); also chained onto the
#              end of a clean unwip-all (set GSADUS_UNWIP_PULLENV=0 to skip)
#   push-env   tombstone — there is no push. Edit the secret in Doppler
#              (dashboard or 'doppler secrets set'), then pull-env per machine.
#
# npm auth seam: the @gsadus scope (@gsadus/pipedrive and friends) lives on
# GitHub Packages, and each consuming repo's committed .npmrc reads its token
# from ${NODE_AUTH_TOKEN} — which npm takes from the PROCESS environment and
# never from a .env.local. pull-env therefore mirrors PM's rendered
# NODE_AUTH_TOKEN into one MANAGED line in the user-level ~/.npmrc (rewritten
# every run) and exports it for the current session, so a fresh pull-env is all
# a machine needs before 'npm install'. Every LATER shell gets the variable back
# from that managed line at profile load (Restore-GSADUsNpmAuthEnv — the repo
# .npmrc's own ${NODE_AUTH_TOKEN} line outranks user config, so an unset
# variable would shadow the token with an empty one). Rotation stays
# Doppler -> pull-env.
#
# One-time machine enrollment: winget install doppler.doppler; doppler login.
# Rotation = change the value once in Doppler; every consumer picks it up.
# The old scp push/pull sync and its rollback guard are retired; they survive
# in gsadus-tools git history if ever needed.

# Render table: one Doppler config -> one gitignored local file (paths relative
# to $GSADUsRoot). Mirrors the render-target table in the vault page.
$GSADUsDopplerRenders = @(
    @{ Project = 'webapp';     Config = 'dev'; Target = 'WebApp\.env.local' }
    @{ Project = 'pm';         Config = 'dev'; Target = 'PM\.env.local' }
    @{ Project = 'webcatalog'; Config = 'prd'; Target = 'WebCatalog\pipeline\.env' }
    @{ Project = 'pngtools';   Config = 'prd'; Target = 'PostProcess\PNGTools\.env' }
)

# The single ~/.npmrc line pull-env owns. The marker comment is what makes the
# rewrite idempotent: both it and any bare auth line for the registry are
# dropped before the fresh pair is appended, so the file never accumulates
# stale tokens. ASCII only — this text lands in an ini file npm parses.
$GSADUsNpmAuthMarker = '# managed by pull-env (GSADUs): @gsadus scope on GitHub Packages - rotate in Doppler'
$GSADUsNpmAuthKey    = '//npm.pkg.github.com/:_authToken='

function Sync-GSADUsNpmAuth {
    # Bridge PM's rendered NODE_AUTH_TOKEN to npm: a managed ~/.npmrc auth line
    # (npm reads no .env file) plus a session export, so 'npm install' resolves
    # the private @gsadus scope immediately after a pull-env. The value is only
    # ever read into memory and written to $HOME — never echoed, never copied
    # into a repo. A missing token is not an error: only PM renders one today.
    $envFile = Join-Path $GSADUsRoot 'PM\.env.local'
    if (-not (Test-Path -LiteralPath $envFile)) { return }
    $match = Select-String -LiteralPath $envFile -Pattern '^\s*NODE_AUTH_TOKEN\s*=' | Select-Object -First 1
    if (-not $match) { return }
    $token = ($match.Line -replace '^\s*NODE_AUTH_TOKEN\s*=\s*', '').Trim().Trim('"').Trim("'")
    if (-not $token) { return }

    $env:NODE_AUTH_TOKEN = $token

    $npmrc = Join-Path $HOME '.npmrc'
    $kept = @()
    if (Test-Path -LiteralPath $npmrc) {
        $kept = @(Get-Content -LiteralPath $npmrc | Where-Object {
            $_ -notlike "$GSADUsNpmAuthMarker*" -and $_.Trim() -notlike "$GSADUsNpmAuthKey*"
        })
    }
    try {
        Set-Content -LiteralPath $npmrc -Value @($kept + @($GSADUsNpmAuthMarker, "$GSADUsNpmAuthKey$token")) -Encoding utf8 -ErrorAction Stop
        Write-Host "  npmauth ~/.npmrc <- NODE_AUTH_TOKEN (@gsadus scope; exported for this session)" -ForegroundColor Cyan
    } catch {
        Write-Host "  WARN  ~/.npmrc — write failed ($($_.Exception.Message)); npm auth left as it was" -ForegroundColor Red
    }
}

function pull-env {
    # Render every target in $GSADUsDopplerRenders. Doppler output is CAPTURED —
    # values never hit the console — written to a temp file BESIDE the target,
    # then swapped in with Move-Item so the live file is replaced atomically. A
    # failed or empty render leaves the existing file untouched.
    if (-not (Get-Command doppler -ErrorAction SilentlyContinue)) {
        Write-Host "  ERROR doppler CLI not found — install and authenticate first:" -ForegroundColor Red
        Write-Host "        winget install doppler.doppler" -ForegroundColor DarkGray
        Write-Host "        doppler login" -ForegroundColor DarkGray
        return
    }
    doppler me *> $null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  ERROR doppler CLI not ready ('doppler me' failed) — run: doppler login" -ForegroundColor Red
        return
    }

    Write-Host "Rendering .env files from Doppler..." -ForegroundColor Cyan
    foreach ($render in $GSADUsDopplerRenders) {
        $project = $render.Project
        $config  = $render.Config
        $rel     = $render.Target
        $target  = Join-Path $GSADUsRoot $rel
        $dir     = Split-Path -Parent $target
        if (-not (Test-Path $dir)) {
            Write-Host "  WARN  $rel — folder missing (repo not cloned here?); skipped" -ForegroundColor Yellow
            continue
        }

        # env-no-quotes, not env: the quoted format backslash-escapes " and \n, which
        # node's dotenv (@next/env) mis-decodes for inline-JSON values (python-dotenv
        # differs again). Raw unquoted lines are the format these files always used.
        # Constraint: no rendered config may hold values with newlines, '#', or edge
        # whitespace — multi-line secrets (e.g. NPM_RC) live only in Vercel-sync configs.
        $rendered = doppler secrets download --project $project --config $config --no-file --format env-no-quotes 2>$null
        if ($LASTEXITCODE -ne 0 -or -not $rendered) {
            Write-Host "  WARN  $rel — render failed ($project/$config); existing file left untouched" -ForegroundColor Red
            continue
        }

        # The temp name keeps the '.env.' prefix so repo .gitignores (and the wip
        # secret-leak guard) still cover it even if a crash ever left it behind.
        $tmp = Join-Path $dir ".env.doppler-tmp-$PID"
        try {
            Set-Content -LiteralPath $tmp -Value $rendered -Encoding utf8 -ErrorAction Stop
            Move-Item -LiteralPath $tmp -Destination $target -Force -ErrorAction Stop
            Write-Host "  render $rel <- $project/$config" -ForegroundColor Cyan
        } catch {
            Write-Host "  WARN  $rel — write failed ($($_.Exception.Message)); existing file left untouched" -ForegroundColor Red
        } finally {
            if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
        }
    }

    Sync-GSADUsNpmAuth
}

function Restore-GSADUsNpmAuthEnv {
    # Every shell needs NODE_AUTH_TOKEN, not just the one that ran pull-env: a
    # consuming repo's committed .npmrc sets
    # '//npm.pkg.github.com/:_authToken=${NODE_AUTH_TOKEN}', and project config
    # OUTRANKS user config — so with the variable unset that empty expansion
    # SHADOWS the managed ~/.npmrc line and npm install fails 401 (measured
    # 2026-08-31). Rehydrating the variable from that same managed line at
    # profile load keeps one owner for the value (Doppler -> pull-env ->
    # ~/.npmrc) while making every new shell install-ready. Silent by design;
    # never overrides a variable the session already set (CI, an ad-hoc token).
    if ($env:NODE_AUTH_TOKEN) { return }
    $npmrc = Join-Path $HOME '.npmrc'
    if (-not (Test-Path -LiteralPath $npmrc)) { return }
    $line = Get-Content -LiteralPath $npmrc -ErrorAction SilentlyContinue |
        Where-Object { $_.Trim() -like "$GSADUsNpmAuthKey*" } | Select-Object -First 1
    if (-not $line) { return }
    $token = $line.Trim().Substring($GSADUsNpmAuthKey.Length).Trim()
    if ($token) { $env:NODE_AUTH_TOKEN = $token }
}

Restore-GSADUsNpmAuthEnv

function push-env {
    # Tombstone — the scp push direction retired with the Doppler cutover
    # (2026-08-25). Local files are rendered FROM Doppler, never pushed back.
    # throw (not Write-Error): callers and 'pwsh -Command' must see a real failure.
    throw ("push-env is retired: secrets are edited in Doppler (dashboard or " +
        "'doppler secrets set --project <p> --config <c> NAME'), then rendered " +
        "locally with 'pull-env' on each machine.")
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

# -- vault wiki-agent scan (Phase 2, activated 2026-07-10) ---------------------
# Deep scan weekly via scheduled task; light scan rides end-day. Register the
# weekly task on ONE machine only — the scan pushes vault main when done.

function vault-scan {
    param([switch]$Light, [switch]$DryRun)
    & "$GSADUsRoot\Vault\scripts\wiki-scan.ps1" -Light:$Light -DryRun:$DryRun
}

function Register-WikiScan {
    $action   = New-ScheduledTaskAction -Execute "pwsh" `
        -Argument "-NoProfile -File `"$GSADUsRoot\Vault\scripts\wiki-scan.ps1`"" `
        -WorkingDirectory $GSADUsRoot
    $trigger  = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 18:00
    $settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 45) `
        -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
    Register-ScheduledTask -TaskName "GSADUs-wiki-scan" -Action $action `
        -Trigger $trigger -Settings $settings -Force | Out-Null
    Write-Host "Weekly vault wiki-scan registered (Sundays 18:00; register on one machine only)." -ForegroundColor Green
}

function Unregister-WikiScan {
    Unregister-ScheduledTask -TaskName "GSADUs-wiki-scan" -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Weekly vault wiki-scan removed." -ForegroundColor Yellow
}
