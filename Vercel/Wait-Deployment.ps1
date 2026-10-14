#requires -Version 7.0
<#
.SYNOPSIS
Wait for one pushed WebApp/PM commit on every Vercel project that repo deploys, through the
authenticated Vercel CLI (read-only).
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
# Every push of WebApp deploys two projects: the App (repo root) and the Site (apps/site).
# Each runs its own ignoreCommand and cancels a push that changed nothing it builds.
$projectNames = @{ WebApp = @('gsadus', 'gsadus-site'); PM = @('gsadus-pm') }[$Project]
$Commit = $Commit.ToLowerInvariant()

function Invoke-VercelJson([string[]]$CliArgs, [switch]$EndState) {
    # Keep stderr separate: Vercel writes progress there and JSON to stdout.
    $raw = & vercel @CliArgs --scope $scope --non-interactive --format json
    $code = $LASTEXITCODE
    $json = try { $raw -join "`n" | ConvertFrom-Json -ErrorAction Stop } catch { $null }
    # inspect --wait exits 1 when the deployment ends ERROR or CANCELED but still prints it.
    if ($code -ne 0 -and -not ($EndState -and $json.readyState)) {
        throw "Vercel CLI failed (exit $code). Resolve the reported auth/API/build error; do not keep polling."
    }
    if ($null -eq $json) { throw 'Vercel CLI printed no JSON.' }
    return $json
}

try {
    $null = Get-Command vercel -ErrorAction Stop
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $found = [ordered]@{}
    Write-Host "Finding $($projectNames -join ' + ') / $Environment / $Commit (discovery every ${PollSeconds}s)."
    do {
        foreach ($name in @($projectNames | Where-Object { -not $found.Contains($_) })) {
            $result = Invoke-VercelJson @('list', $name, '--environment', $Environment,
                '--meta', "githubCommitSha=$Commit")
            if ($null -eq $result.deployments) { throw 'Vercel list returned no deployments field.' }
            # Revalidate the response; never accept another commit or a stale production alias.
            $match = $result.deployments |
                Where-Object { $_.meta.githubCommitSha -eq $Commit -and $_.name -eq $name -and
                    ($_.target -eq $Environment -or ($Environment -eq 'preview' -and $null -eq $_.target)) } |
                Sort-Object createdAt -Descending | Select-Object -First 1
            if ($match) { $found[$name] = $match }
        }
        if ($found.Count -eq $projectNames.Count) { break }
        $remaining = $DiscoveryTimeoutSeconds - $watch.Elapsed.TotalSeconds
        if ($remaining -le 0) {
            $missing = @($projectNames | Where-Object { -not $found.Contains($_) }) -join ', '
            throw "No $missing deployment found for $Commit within ${DiscoveryTimeoutSeconds}s. Check that this commit was pushed and that Vercel's Git integration created its deployments."
        }
        Start-Sleep -Milliseconds ([int](1000 * [Math]::Min($PollSeconds, $remaining)))
    } while ($true)

    # One build budget for all of them: Vercel builds them in parallel or queued, not after us.
    $builds = [Diagnostics.Stopwatch]::StartNew()
    $results = foreach ($name in $projectNames) {
        $deployment = $found[$name]
        if (-not $deployment.url) { throw "Matched $name deployment has no immutable URL." }
        $left = [Math]::Max(1, [int][Math]::Ceiling($BuildTimeoutSeconds - $builds.Elapsed.TotalSeconds))
        Write-Host "Following https://$($deployment.url) [$($deployment.state)] with vercel inspect --wait."
        $info = Invoke-VercelJson @('inspect', $deployment.url, '--wait', '--timeout', "${left}s") -EndState
        # inspect JSON omits Git metadata; list validated the SHA before pinning this URL.
        if ($info.url -ne $deployment.url -or -not $info.id) {
            throw "$name deployment identity changed or its ID is missing; refusing to report it."
        }
        [ordered]@{ project = $name; deploymentId = $info.id; url = "https://$($info.url)"; state = $info.readyState }
    }
    [ordered]@{
        project = $Project
        environment = $Environment
        commit = $Commit
        deployments = @($results)
        elapsedSeconds = [Math]::Round($watch.Elapsed.TotalSeconds, 1)
    } | ConvertTo-Json -Depth 4
    # READY is live. CANCELED was not built (its ignoreCommand skipped this push, or someone
    # canceled it): production keeps its previous build. Every other end state failed.
    $failed = @($results | Where-Object { $_.state -notin 'READY', 'CANCELED' })
    foreach ($f in $failed) {
        [Console]::Error.WriteLine("$($f.project) ended in state '$($f.state)' ($($f.deploymentId)); it is not live. Build logs: vercel inspect $($f.url) --logs")
    }
    exit ([int]($failed.Count -gt 0))
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
