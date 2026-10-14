#requires -Version 7.0
param([string]$Scenario)
$ErrorActionPreference = 'Stop'
$helper = Join-Path $PSScriptRoot 'Wait-Deployment.ps1'

if ($Scenario) {
    # Child-process CLI double: no credentials, network, deployments or sleep longer than 1s.
    # Like the real CLI, inspect --wait exits 1 on any end state but READY and still prints it.
    $global:listCalls = @{}
    $global:ends = @{
        'site-skipped' = @{ 'gsadus-site' = 'CANCELED' }; 'app-skipped' = @{ gsadus = 'CANCELED' }
        'site-error' = @{ 'gsadus-site' = 'ERROR' }; error = @{ gsadus = 'ERROR' }; building = @{ gsadus = 'BUILDING' }
    }
    function global:vercel {
        param([Parameter(ValueFromRemainingArguments)][string[]]$CliArgs)
        $global:LASTEXITCODE = 0
        if ($CliArgs -notcontains '--non-interactive' -or
            $CliArgs -notcontains 'vadim-7430s-projects') { throw 'Missing scope/noninteractive flag' }
        if ($Scenario -eq 'cli-error') { $global:LASTEXITCODE = 1; return }
        if ($Scenario -eq 'invalid-json') { return 'not JSON' }
        $sha = 'a' * 40
        if ($CliArgs[0] -eq 'list') {
            $name = $CliArgs[1]
            $global:listCalls[$name] = 1 + [int]$global:listCalls[$name]
            if ($CliArgs -notcontains "githubCommitSha=$sha") { throw 'Missing SHA filter' }
            if ($Scenario -eq 'missing' -or ($Scenario -eq 'site-missing' -and $name -eq 'gsadus-site') -or
                ($Scenario -eq 'delayed' -and $global:listCalls[$name] -eq 1)) {
                return '{"deployments":[]}'
            }
            if ($Scenario -eq 'wrong-commit') { $sha = 'b' * 40 }
            $target = if ($Scenario -eq 'pm-preview') { $null } else { 'production' }
            return @{ deployments = @(@{ url = "$name-unique.vercel.app"; name = $name; target = $target;
                state = 'BUILDING'; createdAt = 1; meta = @{ githubCommitSha = $sha } }) } |
                ConvertTo-Json -Depth 5 -Compress
        }
        if ($CliArgs[0] -ne 'inspect' -or $CliArgs[1] -notmatch '-unique\.vercel\.app$' -or $CliArgs -notcontains '--wait') {
            throw 'Must inspect the immutable URL with --wait'
        }
        $url = $CliArgs[1]
        $state = ($global:ends[$Scenario] ?? @{})[$url -replace '-unique\.vercel\.app$'] ?? 'READY'
        if ($state -ne 'READY') { $global:LASTEXITCODE = 1 }
        if ($Scenario -eq 'wrong-url') { $url = 'other.vercel.app' }
        return @{ id = 'dpl_test'; url = $url; readyState = $state } | ConvertTo-Json -Compress
    }
    $testProject = if ($Scenario -eq 'pm-preview') { 'PM' } else { 'WebApp' }
    $testEnvironment = if ($Scenario -eq 'pm-preview') { 'preview' } else { 'production' }
    & $helper -Project $testProject -Commit ('a' * 40) -Environment $testEnvironment `
        -DiscoveryTimeoutSeconds 1 -BuildTimeoutSeconds 1 -PollSeconds 1
    exit $LASTEXITCODE
}

# Scenario = expected exit code, and a pattern its combined output must match.
$cases = [ordered]@{
    ready = 0, '"project": "gsadus",[^}]*"state": "READY"[\s\S]*"project": "gsadus-site",[^}]*"state": "READY"'
    'site-skipped' = 0, '"project": "gsadus-site",[^}]*"state": "CANCELED"'
    'app-skipped' = 0, '"project": "gsadus",[^}]*"state": "CANCELED"'
    delayed = 0, '"project": "gsadus-site",[^}]*"state": "READY"'
    'pm-preview' = 0, '"project": "gsadus-pm",[^}]*"state": "READY"'
    'site-error' = 1, "gsadus-site ended in state 'ERROR'"
    error = 1, "gsadus ended in state 'ERROR'"
    building = 1, "gsadus ended in state 'BUILDING'"
    'site-missing' = 1, 'No gsadus-site deployment found'
    missing = 1, 'No gsadus, gsadus-site deployment found'
    'wrong-commit' = 1, 'No gsadus, gsadus-site deployment found'
    'cli-error' = 1, 'Vercel CLI failed'
    'invalid-json' = 1, 'printed no JSON'
    'wrong-url' = 1, 'identity changed'
}
foreach ($case in $cases.GetEnumerator()) {
    $expectedCode, $pattern = $case.Value
    $output = (& pwsh -NoProfile -File $PSCommandPath -Scenario $case.Key 2>&1) -join "`n"
    $code = $LASTEXITCODE
    if ($code -ne $expectedCode) { throw "$($case.Key): expected exit $expectedCode, got $code`n$output" }
    if ($output -notmatch $pattern) { throw "$($case.Key): output does not match /$pattern/`n$output" }
    Write-Host "PASS $($case.Key)"
}
Write-Host "$($cases.Count) deployment-wait tests passed."
