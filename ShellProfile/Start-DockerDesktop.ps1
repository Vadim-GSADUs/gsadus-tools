# Start Docker Desktop outside every agent app, then wait for its Linux engine.
#
# Claude and Codex desktop are MSIX apps. A program an agent starts from its tool
# shell (Start-Process, &, `docker desktop start`) joins that app's job and file-system
# container. Windows kills it whenever the app updates or restarts, AppData folders it
# creates land in the app's private package store, and inside that container Docker
# cannot open, remove or rename its own AF_UNIX socket files (error 1920), so every
# later start fails (SETUP-VERIFICATION.md -> Docker repair). Explorer starts the
# program as its own child instead, outside every app container, where Docker clears
# its leftover sockets itself. Stop Docker with `docker desktop stop`.
[CmdletBinding()]
param([int]$TimeoutSeconds = 240)
$ErrorActionPreference = 'Stop'
$dockerExe = Join-Path $env:ProgramFiles 'Docker\Docker\Docker Desktop.exe'
$backendLog = Join-Path $env:LOCALAPPDATA 'Docker\log\host\com.docker.backend.exe.log'
if (-not (Test-Path -LiteralPath $dockerExe)) { throw "Docker Desktop is not installed at $dockerExe" }

function Get-EngineVersion {
    $version = docker --context desktop-linux version --format '{{.Server.Version}}' 2>$null
    if ($LASTEXITCODE -eq 0 -and $version) { $version }
}

# The first backend process; the rest are its own children.
function Get-RootBackend {
    $backends = @(Get-CimInstance Win32_Process -Filter "Name='com.docker.backend.exe'")
    $backends | Where-Object { $backends.ProcessId -notcontains $_.ParentProcessId } | Select-Object -First 1
}

$root = Get-RootBackend
if ($root) {
    # Inside an app container, the Secrets Engine socket Docker creates at startup lands
    # in that app's package store instead of %LOCALAPPDATA%\docker-secrets-engine.
    $hosts = Resolve-Path (Join-Path $env:LOCALAPPDATA 'Packages\*\LocalCache\Local\docker-secrets-engine') -ErrorAction SilentlyContinue |
        ForEach-Object { Get-ChildItem -LiteralPath $_.Path -Force } |
        Where-Object { $_.Name -eq 'engine.sock' -and $_.LastWriteTime -ge $root.CreationDate.AddSeconds(-5) } |
        ForEach-Object { ($_.FullName -split '\\Packages\\')[1].Split('\')[0] }
    if ($hosts) {
        Write-Warning ("Docker Desktop was started from an agent shell and runs inside the $($hosts -join ', ') " +
            "app. Windows kills it when that app closes or updates, and it cannot clear its sockets there. " +
            "When nothing needs Docker, run 'docker desktop stop' and start it again with this script.")
    }
    $since = $root.CreationDate.ToUniversalTime()
} else {
    $since = (Get-Date).ToUniversalTime()
    Start-Process explorer.exe -ArgumentList ('"{0}"' -f $dockerExe)
    Write-Host 'Launched Docker Desktop through Explorer.'
}

$sinceStamp = $since.ToString('yyyy-MM-ddTHH:mm:ss')
$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
while ((Get-Date) -lt $deadline) {
    $version = Get-EngineVersion
    if ($version) {
        Write-Host "Docker Desktop engine ready: $version"
        return
    }
    $failure = Get-Content -LiteralPath $backendLog -Tail 400 -ErrorAction SilentlyContinue |
        Where-Object { $_ -match '^\[(\S+?)Z\]\[com\.docker\.backend\.exe\] backend cancelling with error: ' -and
            $Matches[1] -ge $sinceStamp } |
        Select-Object -Last 1
    if ($failure) {
        throw ("Docker Desktop backend failed to start. Do not force-kill Docker; follow " +
            "SETUP-VERIFICATION.md -> Docker repair.`n$failure")
    }
    Start-Sleep -Seconds 3
}
throw "Docker Desktop engine was not ready after $TimeoutSeconds seconds."
