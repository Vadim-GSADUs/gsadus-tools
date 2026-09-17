#requires -Version 7.0
param([string]$Scenario)
$ErrorActionPreference = 'Stop'
$helper = Join-Path $PSScriptRoot 'Wait-Deployment.ps1'

if ($Scenario) {
    # Child-process CLI double: no credentials, network, deployments or sleep longer than 1s.
    $script:listCalls = 0
    function global:vercel {
        param([Parameter(ValueFromRemainingArguments)][string[]]$CliArgs)
        $global:LASTEXITCODE = 0
        if ($CliArgs -notcontains '--non-interactive' -or
            $CliArgs -notcontains 'vadim-7430s-projects') { throw 'Missing scope/noninteractive flag' }
        if ($Scenario -eq 'cli-error') { $global:LASTEXITCODE = 1; return }
        if ($Scenario -eq 'invalid-json') { return 'not JSON' }
        $sha = 'a' * 40
        $url = 'gsadus-unique.vercel.app'
        if ($CliArgs[0] -eq 'list') {
            $script:listCalls++
            if ($CliArgs -notcontains "githubCommitSha=$sha") { throw 'Missing SHA filter' }
            if ($Scenario -eq 'missing' -or ($Scenario -eq 'delayed' -and $script:listCalls -eq 1)) {
                return '{"deployments":[]}'
            }
            if ($Scenario -eq 'wrong-commit') { $sha = 'b' * 40 }
            $name = if ($Scenario -eq 'pm-preview') { 'gsadus-pm' } else { 'gsadus' }
            $target = if ($Scenario -eq 'pm-preview') { $null } else { 'production' }
            return @{ deployments = @(@{ url = $url; name = $name; target = $target;
                state = 'BUILDING'; createdAt = 1; meta = @{ githubCommitSha = $sha } }) } |
                ConvertTo-Json -Depth 5 -Compress
        }
        if ($CliArgs[0] -ne 'inspect' -or $CliArgs[1] -ne $url -or $CliArgs -notcontains '--wait') {
            throw 'Must inspect the immutable URL with --wait'
        }
        $state = switch ($Scenario) { 'error' { 'ERROR' }; 'canceled' { 'CANCELED' }; default { 'READY' } }
        if ($Scenario -eq 'wrong-url') { $url = 'other.vercel.app' }
        return @{ id = 'dpl_test'; url = $url; readyState = $state } | ConvertTo-Json -Compress
    }
    $testProject = if ($Scenario -eq 'pm-preview') { 'PM' } else { 'WebApp' }
    $testEnvironment = if ($Scenario -eq 'pm-preview') { 'preview' } else { 'production' }
    & $helper -Project $testProject -Commit ('a' * 40) -Environment $testEnvironment `
        -DiscoveryTimeoutSeconds 1 -BuildTimeoutSeconds 1 -PollSeconds 1
    exit $LASTEXITCODE
}

$cases = [ordered]@{
    ready = 0; delayed = 0; 'pm-preview' = 0; missing = 1; 'wrong-commit' = 1
    error = 1; canceled = 1; 'cli-error' = 1; 'invalid-json' = 1; 'wrong-url' = 1
}
foreach ($case in $cases.GetEnumerator()) {
    $output = & pwsh -NoProfile -File $PSCommandPath -Scenario $case.Key 2>&1
    $code = $LASTEXITCODE
    if ($code -ne $case.Value) { throw "$($case.Key): expected $($case.Value), got $code`n$output" }
    if ($code -eq 0 -and ($output -join "`n") -notmatch '"state": "READY"') {
        throw "$($case.Key): success must include READY result"
    }
    Write-Host "PASS $($case.Key)"
}
Write-Host "$($cases.Count) deployment-wait tests passed."
