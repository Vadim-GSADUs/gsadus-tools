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

# Archived repos kept on disk as read-only reference. Their GitHub remotes are
# archived (pushes rejected), so wip/unwip must skip them. Paths relative to $GSADUsRoot.
$GSADUsRetiredRepos = @(
    'PostProcess\DigitalDarkroom'   # archived 2026-07-07; superseded by PNGTools darkroom
    'PostProcess\Darkroom'          # archived 2026-07-07; stalled web console, PNGTools outgrew it
    'SiteCheck'                     # archived 2026-08-06; spec+spike only, module ships from WebApp
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

# -- cross-PC secret (.env) sync ----------------------------------------------
# .env files are gitignored (web-catalog has an external collaborator), so they
# can't ride git like everything else. Instead they hop directly between PCs over
# the existing Tailscale/OpenSSH link (see Vault wiki: remote-access, wip-sync).
#
#   push-env   copy THIS machine's .env files to the other PC
#   pull-env   copy the other PC's .env files to THIS machine
#   -Force     override the rollback guard (see below)
#
# Run these at the machine's OWN console. Don't drive them through a nested SSH
# session (e.g. ssh OTHER-PC "... push-env") — the double remote-pwsh hop hangs.
# To sync from afar, just run the opposite verb locally: to fetch the other PC's
# secrets, run pull-env here; to send yours, run push-env here.
#
# Rollback guard: a file is only overwritten when the source copy DIFFERS AND is
# NEWER than the destination. If the destination is newer (e.g. you just rolled a
# key there), the transfer is BLOCKED so a stale secret can't clobber a fresh one.
# -Force overrides it, but first stashes the newer destination copy to
# %TEMP%\gsadus-env-backups (kept out of any repo so it can't trip the wip leak
# guard), so nothing is ever truly lost. scp -p preserves mtimes so the
# newer/older comparison stays honest across machines.

# SSH target of each machine, keyed by COMPUTERNAME (Tailscale MagicDNS names).
$GSADUsSshTargets = @{
    'VG-HOME'      = 'User@vg-home'
    'GSADUS-VADIM' = 'Vadim@gsadus-vadim'
}

# .env files that live outside git, relative to $GSADUsRoot.
$GSADUsEnvFiles = @(
    'WebCatalog\pipeline\.env'
    'PostProcess\PNGTools\.env'
    'SiteCheck\spike\config.js'
    'PM\.env.local'
)

function Get-GSPeer {
    $me = $env:COMPUTERNAME.ToUpper()
    if (-not $GSADUsSshTargets.ContainsKey($me)) {
        Write-Host "  ERROR unknown machine '$me' — add it to `$GSADUsSshTargets in profile.ps1" -ForegroundColor Red
        return $null
    }
    $GSADUsSshTargets.Keys | Where-Object { $_ -ne $me } |
        ForEach-Object { $GSADUsSshTargets[$_] } | Select-Object -First 1
}

function Invoke-GSRemote {
    # Run a pwsh snippet on the peer. -EncodedCommand sidesteps every SSH/pwsh
    # quoting pitfall for the small scripts we send.
    #   -n            : never read local stdin (good hygiene; harmless locally)
    #   BatchMode     : fail fast instead of blocking on any auth/host-key prompt
    #   ConnectTimeout: don't hang if the peer is offline
    # These run push/pull-env LOCALLY (one ssh hop to the peer). Driving the whole
    # command through a nested SSH session is unsupported — see header note.
    param([string]$Peer, [string]$Script)
    $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Script))
    ssh -n -o BatchMode=yes -o ConnectTimeout=10 $Peer "pwsh -NoProfile -EncodedCommand $enc"
}

function Format-EnvTime {
    param([long]$Ticks)
    if ($Ticks -le 0) { return '(none)' }
    [datetime]::new($Ticks, [DateTimeKind]::Utc).ToLocalTime().ToString('MM-dd HH:mm')
}

function Sync-OneEnv {
    param(
        [Parameter(Mandatory)][string]$Rel,
        [Parameter(Mandatory)][ValidateSet('push','pull')][string]$Direction,
        [switch]$Force
    )
    $peer = Get-GSPeer
    if (-not $peer) { return 'error' }

    $local     = Join-Path $GSADUsRoot $Rel
    $remoteFs  = "C:/GSADUs/$($Rel -replace '\\','/')"
    $remoteScp = "${peer}:$remoteFs"

    # local meta
    if (Test-Path $local) {
        $lHash  = (Get-FileHash $local -Algorithm SHA256).Hash
        $lTicks = (Get-Item $local).LastWriteTimeUtc.Ticks
    } else { $lHash = $null; $lTicks = 0 }

    # remote meta (hash + UTC ticks, or MISSING) in one round trip
    $metaScript = @"
`$p = '$remoteFs'
if (Test-Path `$p) { (Get-FileHash `$p -Algorithm SHA256).Hash + '|' + (Get-Item `$p).LastWriteTimeUtc.Ticks } else { 'MISSING' }
"@
    $raw = (Invoke-GSRemote $peer $metaScript | Out-String).Trim()
    if (-not $raw) {
        Write-Host "  ERROR $Rel — no response from $peer (online? SSH up?)" -ForegroundColor Red
        return 'error'
    }
    if ($raw -eq 'MISSING') { $rHash = $null; $rTicks = 0 }
    else { $parts = $raw -split '\|', 2; $rHash = $parts[0]; $rTicks = [long]$parts[1] }

    if ($Direction -eq 'push') { $srcHash=$lHash; $srcTicks=$lTicks; $dstTicks=$rTicks; $dstLabel='remote'; $arrow='->' }
    else                       { $srcHash=$rHash; $srcTicks=$rTicks; $dstTicks=$lTicks; $dstLabel='local';  $arrow='<-' }

    if (-not $srcHash) {
        Write-Host "  skip  $Rel (no source file to $Direction)" -ForegroundColor DarkGray
        return 'skip'
    }
    if ($lHash -and $rHash -and $lHash -eq $rHash) {
        Write-Host "  ok    $Rel (in sync)" -ForegroundColor DarkGray
        return 'insync'
    }

    # content differs (or dest missing) — block if we'd overwrite a NEWER dest
    if ($dstTicks -gt $srcTicks -and -not $Force) {
        Write-Host ("  BLOCK {0} — {1} is NEWER (src {2} < dst {3}); refusing to roll back. Use -Force to override." `
            -f $Rel, $dstLabel, (Format-EnvTime $srcTicks), (Format-EnvTime $dstTicks)) -ForegroundColor Red
        return 'blocked'
    }

    # overwriting a newer dest under -Force: stash the newer copy first
    if ($Force -and $dstTicks -gt $srcTicks) {
        $bakDir = Join-Path $env:TEMP 'gsadus-env-backups'
        New-Item -ItemType Directory -Force $bakDir | Out-Null
        $bak = Join-Path $bakDir ('{0}.{1}.bak' -f ($Rel -replace '[\\/]','_'), (Get-Date -Format 'yyyyMMdd-HHmmss'))
        if ($Direction -eq 'push') { scp -p -q -o BatchMode=yes -o ConnectTimeout=10 "$remoteScp" "$bak" } else { Copy-Item $local $bak -Force }
        Write-Host "       stashed newer $dstLabel copy -> $bak" -ForegroundColor DarkYellow
    }

    if ($Direction -eq 'push') { scp -p -q -o BatchMode=yes -o ConnectTimeout=10 "$local" "$remoteScp" } else { scp -p -q -o BatchMode=yes -o ConnectTimeout=10 "$remoteScp" "$local" }
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  sync  $Rel ($dstLabel updated, $arrow $peer)" -ForegroundColor Cyan
        return 'synced'
    }
    Write-Host "  ERROR $Rel — scp failed (exit $LASTEXITCODE)" -ForegroundColor Red
    return 'error'
}

function Invoke-EnvSync {
    param([ValidateSet('push','pull')][string]$Direction, [switch]$Force)
    $peer = Get-GSPeer
    if (-not $peer) { return }
    Write-Host ("{0}ing .env files {1} {2}" -f $Direction, ($Direction -eq 'push' ? 'to' : 'from'), $peer) -ForegroundColor Cyan
    $blocked = 0; $synced = 0
    foreach ($f in $GSADUsEnvFiles) {
        switch (Sync-OneEnv -Rel $f -Direction $Direction -Force:$Force) {
            'blocked' { $blocked++ }
            'synced'  { $synced++ }
        }
    }
    if ($blocked -gt 0) {
        Write-Host ("  !! {0} file(s) BLOCKED to avoid a rollback. Re-run '{1}-env -Force' only if you're sure." -f $blocked, $Direction) -ForegroundColor Red
    } elseif ($synced -eq 0) {
        Write-Host "  nothing to do — all in sync." -ForegroundColor DarkGray
    }
}

function push-env { param([switch]$Force) Invoke-EnvSync -Direction push -Force:$Force }
function pull-env { param([switch]$Force) Invoke-EnvSync -Direction pull -Force:$Force }

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
