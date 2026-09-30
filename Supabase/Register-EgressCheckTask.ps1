<#
.SYNOPSIS
    Register (or remove) the daily Task Scheduler task that runs the Supabase egress check:
    GSADUs\supabase-egress-check. No admin rights needed. Register it on ONE PC only (the owner
    chose the office PC, gsadus-vadim, 2026-09-29).

.DESCRIPTION
    Action:  pwsh -NoProfile -NonInteractive -WindowStyle Hidden -Command
             node egress-probe.mjs check, output appended to .state\check.log (gitignored)
    Trigger: daily at 08:00; a run missed while the PC was off starts when it is next available.
    Runs interactively as the current user (the Doppler login lives in the user profile),
    limited privileges, no stored password.

    Why daily: each check saves the snapshot the next one measures against, so the log shows a
    24 h rate, and a regression shows within a day instead of in Supabase's overage email. One
    run fetches about 15 KB, roughly 0.5 MB per billing cycle.

    Usage:   pwsh -File C:\GSADUs\Tools\Supabase\Register-EgressCheckTask.ps1 [-Unregister]
    Check:   Get-ScheduledTask -TaskPath \GSADUs\ -TaskName supabase-egress-check | Get-ScheduledTaskInfo
             (LastTaskResult 2 = over the limit; the details are the tail of .state\check.log)
    Run now: Start-ScheduledTask -TaskPath \GSADUs\ -TaskName supabase-egress-check
#>
[CmdletBinding()]
param([switch]$Unregister)

$taskPath = '\GSADUs\'
$taskName = 'supabase-egress-check'
$probe    = Join-Path $PSScriptRoot 'egress-probe.mjs'
$log      = Join-Path $PSScriptRoot '.state\check.log'

if ($Unregister) {
    Unregister-ScheduledTask -TaskPath $taskPath -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
    "removed $taskPath$taskName (if it existed)"
    exit 0
}

if (-not (Test-Path (Join-Path $PSScriptRoot 'node_modules\pg'))) {
    npm ci --prefix $PSScriptRoot --no-audit --no-fund --silent
    if ($LASTEXITCODE -ne 0) { throw "npm ci failed in $PSScriptRoot" }
}
New-Item -ItemType Directory -Force (Split-Path $log) | Out-Null

$pwsh = (Get-Command pwsh.exe -ErrorAction Stop).Source
$node = (Get-Command node.exe -ErrorAction Stop).Source
# node writes UTF-8; without the first statement pwsh decodes it as the OEM code page.
$command = "[Console]::OutputEncoding = [Text.UTF8Encoding]::new(); & '$node' '$probe' check *>&1 | Out-File -Append -Encoding utf8 '$log'; exit `$LASTEXITCODE"
$action  = New-ScheduledTaskAction -Execute $pwsh -Argument "-NoProfile -NonInteractive -WindowStyle Hidden -Command `"$command`""
$trigger = New-ScheduledTaskTrigger -Daily -At '08:00'
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -MultipleInstances IgnoreNew
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskPath $taskPath -TaskName $taskName -Action $action -Trigger $trigger `
    -Settings $settings -Principal $principal -Force `
    -Description 'Daily Supabase egress check for the shared gsadus-web-catalog project (read-only; exit 2 = over the limit). Log: C:\GSADUs\Tools\Supabase\.state\check.log. Script: C:\GSADUs\Tools\Supabase\Register-EgressCheckTask.ps1' | Out-Null

Get-ScheduledTask -TaskPath $taskPath -TaskName $taskName | Select-Object TaskPath, TaskName, State
