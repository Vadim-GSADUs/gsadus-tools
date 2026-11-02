# Lists the node dev servers listening on the GSADUs port lanes (AGENTS.md → Dev Port Map, plus
# the Site's 3019), then stops each one with its whole node process tree after one yes.
# Anything else on those ports (Docker, for one) is shown and left alone.
$lanes = 3009, 3010, 3019, 3111, 3200, 3209, 3211, 3300, 3309, 3311
$procs = @{}
Get-CimInstance Win32_Process | ForEach-Object { $procs[[int]$_.ProcessId] = $_ }

function Get-TreeRoot([int]$id) {
    # Climb while the parent is node too (npm -> next dev -> server), so the kill takes the
    # launcher with it and nothing respawns the server; the terminal shell above stays.
    while ($procs[$id] -and $procs[[int]$procs[$id].ParentProcessId].Name -eq 'node.exe') {
        $id = [int]$procs[$id].ParentProcessId
    }
    $id
}

$servers = Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue |
    Where-Object LocalPort -in $lanes | Sort-Object LocalPort, OwningProcess -Unique |
    ForEach-Object {
        $p = $procs[[int]$_.OwningProcess]
        $where = if ($p.CommandLine -match 'GSADUs\\([^\\]+)(\\\.claude\\worktrees\\([^\\]+))?') {
            if ($Matches[3]) { "$($Matches[1]) (worktree $($Matches[3]))" } else { $Matches[1] }
        } else { '?' }
        [pscustomobject]@{ Port = $_.LocalPort; Repo = $where; Process = $p.Name; Pid = [int]$_.OwningProcess }
    }

if (-not $servers) { Write-Host 'No dev servers are running.'; Start-Sleep 3; return }
$servers | Format-Table -AutoSize | Out-Host
$node = @($servers | Where-Object Process -eq 'node.exe')
if (-not $node) { Write-Host 'Nothing here is a node dev server, so nothing to stop.'; Start-Sleep 5; return }

try { $answer = Read-Host "Stop the $($node.Count) node server(s)? [Y/n]" } catch { $answer = 'n' }
if ($answer -notin '', 'y', 'yes') { Write-Host 'Left running.'; Start-Sleep 2; return }
foreach ($root in ($node | ForEach-Object { Get-TreeRoot $_.Pid } | Sort-Object -Unique)) {
    taskkill /PID $root /T /F | Out-Null
}
Write-Host 'Stopped.'; Start-Sleep 2
