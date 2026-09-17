#requires -Version 7.0
<#
.SYNOPSIS
Wait for one pushed WebApp/PM commit through the authenticated Vercel CLI (read-only).
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet('WebApp', 'PM')][string]$Project,
    [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F]{40}$')][string]$Commit,
    [ValidateSet('production', 'preview')][string]$Environment = 'production',
    [ValidateRange(1, 3600)][int]$DiscoveryTimeoutSeconds = 120,
    [ValidateRange(1, 3600)][int]$BuildTimeoutSeconds = 600,
    [ValidateRange(1, 30)][int]$PollSeconds = 3
)

$ErrorActionPreference = 'Stop'
$scope = 'vadim-7430s-projects'
$projectName = @{ WebApp = 'gsadus'; PM = 'gsadus-pm' }[$Project]
$Commit = $Commit.ToLowerInvariant()

function Invoke-VercelJson([string[]]$CliArgs) {
    # Keep stderr separate: Vercel writes progress there and JSON to stdout.
    $raw = & vercel @CliArgs --scope $scope --non-interactive --format json
    if ($LASTEXITCODE -ne 0) {
        throw "Vercel CLI failed (exit $LASTEXITCODE). Resolve the reported auth/API/build error; do not keep polling."
    }
    return ($raw -join "`n" | ConvertFrom-Json -ErrorAction Stop)
}

try {
    $null = Get-Command vercel -ErrorAction Stop
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $deployment = $null
    Write-Host "Finding $projectName / $Environment / $Commit (discovery every ${PollSeconds}s)."
    do {
        $result = Invoke-VercelJson @('list', $projectName, '--environment', $Environment,
            '--meta', "githubCommitSha=$Commit")
        if ($null -eq $result.deployments) { throw 'Vercel list returned no deployments field.' }
        # Revalidate the response; never accept another commit or a stale production alias.
        $deployment = $result.deployments |
            Where-Object { $_.meta.githubCommitSha -eq $Commit -and $_.name -eq $projectName -and
                ($_.target -eq $Environment -or ($Environment -eq 'preview' -and $null -eq $_.target)) } |
            Sort-Object createdAt -Descending | Select-Object -First 1
        if ($deployment) { break }
        $remaining = $DiscoveryTimeoutSeconds - $watch.Elapsed.TotalSeconds
        if ($remaining -le 0) {
            throw "No deployment found for $Commit within ${DiscoveryTimeoutSeconds}s. Check that this commit was pushed and Vercel's Git integration/ignored-build rules accepted it."
        }
        Start-Sleep -Milliseconds ([int](1000 * [Math]::Min($PollSeconds, $remaining)))
    } while ($true)

    if (-not $deployment.url) { throw 'Matched deployment has no immutable URL.' }
    Write-Host "Following https://$($deployment.url) [$($deployment.state)] with vercel inspect --wait."
    $info = Invoke-VercelJson @('inspect', $deployment.url, '--wait', '--timeout', "${BuildTimeoutSeconds}s")
    if ($info.readyState -ne 'READY') {
        throw "Deployment ended in state '$($info.readyState)' ($($info.id)). Inspect its build logs; it is not Ready."
    }
    # inspect JSON omits Git metadata; list validated the SHA before pinning this URL.
    if ($info.url -ne $deployment.url -or -not $info.id) {
        throw 'Deployment identity changed or its ID is missing; refusing to report success.'
    }
    [ordered]@{
        project = $Project
        environment = $Environment
        commit = $Commit
        deploymentId = $info.id
        url = "https://$($info.url)"
        state = $info.readyState
        elapsedSeconds = [Math]::Round($watch.Elapsed.TotalSeconds, 1)
    } | ConvertTo-Json
    exit 0
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
