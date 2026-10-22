# Offline checks for 'remote' (profile.ps1): its Sunshine log reading, status text, menu
# and actions, with every Moonlight launch, SSH call, process kill and service restart
# mocked. It starts and ends no stream. The only real network use is ssh to a TEST-NET
# address and an .invalid name, which must fail fast.
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'profile.ps1')
function Assert($condition, $message) { if (-not $condition) { throw $message } }
function Start-Sleep { }   # the actions poll mocked state; nothing to wait for

$sampleLog = @(
    '[2026-10-21 08:20:00.100]: Info: Sunshine version: 2026.914.233613 commit: 63d35f7'
    '[2026-10-21 08:23:02.091]: Info: CLIENT CONNECTED'
    '[2026-10-21 16:54:38.992]: Info: CLIENT DISCONNECTED'
    '[2026-10-21 17:31:31.353]: Info: CLIENT CONNECTED'
)
$liveLog = $sampleLog                 # a client is streaming since 17:31
$idleLog = $sampleLog[0..2]           # the last stream ended at 16:54
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('gsadus-remote-test-' + [guid]::NewGuid().ToString('N') + '.log')
$savedSsh = $env:SSH_CONNECTION       # this test may itself run over SSH
$env:SSH_CONNECTION = $null

function New-TestState {
    # A Get-RemoteState result as seen from VG-HOME.
    param([int]$Moonlight = 0, [bool]$Reached = $true, [int]$PeerMoonlight = 0,
        [bool]$OutLive = $false, [bool]$InLive = $false, [string]$InService = 'Running')
    [pscustomobject]@{
        Here          = 'VG-HOME'
        Peer          = Get-StreamPeer -ComputerName 'VG-HOME'
        Moonlight     = @(for ($i = 0; $i -lt $Moonlight; $i++) { [pscustomobject]@{ Id = 100 + $i } })
        Incoming      = ConvertFrom-SunshineLog -Lines $(if ($InLive) { $liveLog } else { $idleLog }) -Service $InService
        PeerReached   = $Reached
        PeerError     = if ($Reached) { $null } else { 'ssh: connect to host gsadus-vadim port 22: Connection timed out' }
        PeerMoonlight = @(for ($i = 0; $i -lt $PeerMoonlight; $i++) { Get-Date })
        Outgoing      = if ($Reached) { ConvertFrom-SunshineLog -Lines $(if ($OutLive) { $liveLog } else { $idleLog }) } else { $null }
    }
}
function Get-Keys($State) { (@(Get-RemoteActions $State) | ForEach-Object { $_.Key }) -join ',' }
function Get-Status($State) { (Get-RemoteStatusLines $State | ForEach-Object { $_.Text }) -join "`n" }

try {
    # -- the pair ------------------------------------------------------------------
    Assert ((Get-StreamPeer -ComputerName 'VG-HOME').Name -eq 'GSADUS-VADIM') 'VG-HOME must pair with GSADUS-VADIM'
    $p = Get-StreamPeer -ComputerName 'gsadus-vadim'
    Assert ($p.Name -eq 'VG-HOME' -and $p.Moonlight -eq 'VG-Home' -and $p.SshUser -eq 'User') 'GSADUS-VADIM must reach VG-HOME under its paired name'
    $failed = $false; try { Get-StreamPeer -ComputerName 'LAPTOP' } catch { $failed = $true }
    Assert $failed 'An unknown PC must throw'

    # -- reading sunshine.log ------------------------------------------------------
    $s = ConvertFrom-SunshineLog -Lines $liveLog
    Assert ($s.Live -and $s.Since -eq [datetime]'2026-10-21 17:31:31' -and $s.Ended -eq [datetime]'2026-10-21 16:54:38') 'The last CONNECTED must make it live'
    $s = ConvertFrom-SunshineLog -Lines $idleLog
    Assert (-not $s.Live -and $s.Ended -eq [datetime]'2026-10-21 16:54:38' -and $s.Started -eq [datetime]'2026-10-21 08:20:00') 'A DISCONNECTED must end it'
    $s = ConvertFrom-SunshineLog -Lines ($liveLog + '[2026-10-21 18:00:00.000]: Info: Sunshine version: x')
    Assert (-not $s.Live -and -not $s.Since -and $s.Started -eq [datetime]'2026-10-21 18:00:00') 'A Sunshine start must reset it'
    Assert (-not (ConvertFrom-SunshineLog -Lines $liveLog -Service 'Stopped').Live) 'A stopped service must not be live'
    $s = ConvertFrom-SunshineLog -Lines @('Max Luminance      : 301.849 nits', '[2026-10-21 08:23:01.378]: Info: Executing [Desktop]')
    Assert (-not $s.Live -and -not $s.Since -and -not $s.Started) 'Other log lines must be ignored'

    Set-Content -LiteralPath $fixture -Value ($liveLog + 'Device Description : NVIDIA GeForce RTX 4060')
    $s = Get-SunshineStream -Log $fixture -ServiceName 'GSADUs-no-such-service'
    Assert ($s.Service -eq 'Missing' -and -not $s.Live -and $s.Since -eq [datetime]'2026-10-21 17:31:31') 'The file must be read; a missing service is not live'
    $s = Get-SunshineStream -Log $fixture -ServiceName 'Winmgmt'   # always running: stands in for Sunshine
    Assert ($s.Live -and $s.Service -eq 'Running') 'A running service with a CONNECTED log must be live'

    # -- the SSH probe, run by a fresh pwsh exactly as the far side runs it ---------
    & {
        $GSADUsSunshineLog = $fixture; $GSADUsSunshineService = 'Winmgmt'
        function Invoke-StreamPeer {
            param([string]$Script, [int]$TimeoutSec)
            $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Script))
            $out = @(pwsh -NoProfile -NonInteractive -EncodedCommand $enc 2>$null)
            [pscustomobject]@{ Ok = $LASTEXITCODE -eq 0; Output = $out; Error = $null }
        }
        function Get-StreamMoonlight { @([pscustomobject]@{ Id = 1 }) }
        $state = Get-RemoteState
        Assert ($state.PeerReached -and -not $state.PeerError) 'The probe must answer with readable JSON'
        Assert ($state.Outgoing.Live -and ([datetime]$state.Outgoing.Since) -eq [datetime]'2026-10-21 17:31:31') 'The probe must carry the far Sunshine state'
        Assert ($state.Incoming.Live -and $state.Moonlight.Count -eq 1) 'The local half must be read here'
        Assert ((Get-Keys $state) -like 'disconnect,*') 'Streaming out (and in, in this fixture) must offer disconnect first'
    }
    & {
        function Invoke-StreamPeer { [pscustomobject]@{ Ok = $false; Output = @(); Error = 'ssh: connect to host gsadus-vadim port 22: Connection timed out' } }
        function Get-StreamMoonlight { @() }
        function Get-SunshineStream { ConvertFrom-SunshineLog -Lines $liveLog }
        $state = Get-RemoteState
        Assert (-not $state.PeerReached -and $state.PeerError -like '*timed out') 'An SSH failure must be reported, not thrown'
        Assert ((Get-Keys $state) -eq 'connect,restart') 'SSH down with a stream into this PC must offer the Sunshine restart'
        Assert ((Get-Status $state) -like '*over SSH: ssh: connect*') 'The status must say why the peer is unknown'
    }

    # -- the menu and the status, state by state -----------------------------------
    $st = New-TestState
    Assert ((Get-Keys $st) -eq 'connect') 'Idle must offer only connect'
    Assert ((Get-Status $st) -like '*VG-HOME -> GSADUS-VADIM*none (last ended *16:54)*GSADUS-VADIM -> VG-HOME*none (last ended *16:54)*') 'Idle status'
    $st = New-TestState -Moonlight 1 -OutLive $true
    Assert ((Get-Keys $st) -eq 'disconnect' -and (Get-RemoteActions $st).Label -eq 'Disconnect my stream to GSADUS-VADIM') 'Streaming out'
    Assert ((Get-Status $st) -like '*VG-HOME -> GSADUS-VADIM*live since *17:31*') 'Streaming-out status'
    $st = New-TestState -Moonlight 1
    Assert ((Get-RemoteActions $st).Label -like 'Close Moonlight here*') 'Moonlight open but not streaming'
    $st = New-TestState -Moonlight 1 -Reached $false
    Assert ((Get-RemoteActions $st).Label -like 'Disconnect my stream*' -and (Get-Status $st) -like '*unconfirmed*') 'Moonlight open, peer unknown'
    $st = New-TestState -InLive $true -PeerMoonlight 1
    Assert ((Get-Keys $st) -eq 'connect,kick') 'A stream into this PC must offer kick'
    Assert ((Get-Status $st) -like '*GSADUS-VADIM -> VG-HOME*live since *17:31*') 'Incoming status'
    $st = New-TestState -InLive $true
    Assert ((Get-Keys $st) -eq 'connect,restart' -and (Get-Status $st) -like "*still holds a session*Moonlight is closed*") 'A stale Sunshine session must offer the restart'
    $st = New-TestState -InLive $true -InService 'Stopped' -PeerMoonlight 1
    Assert ((Get-Keys $st) -eq 'connect' -and (Get-Status $st) -like '*Sunshine is stopped here*') 'A stopped Sunshine has nothing to end'
    $st = New-TestState -OutLive $true
    Assert ((Get-Status $st) -like "*none from here; GSADUS-VADIM's Sunshine has another client*") 'A foreign client on the peer'

    # -- the menu dispatches the listed action ---------------------------------------
    & {
        $script:calls = @()
        function Get-RemoteState { New-TestState -InLive $true -PeerMoonlight 1 }
        function Connect-RemoteStream { $script:calls += 'connect' }
        function Stop-RemoteIncoming { $script:calls += 'kick' }
        function Read-Host { '2' }
        remote 6>$null
        Assert (($script:calls -join ',') -eq 'kick') 'Choice 2 must run the second listed action'
        function Read-Host { '' }
        remote 6>$null
        function Read-Host { '7' }
        remote 6>$null
        function Read-Host { throw 'non-interactive' }
        remote 6>$null
        Assert (($script:calls -join ',') -eq 'kick') 'Enter, an unlisted number or no console must run nothing'
    }

    # -- connect ---------------------------------------------------------------------
    & {
        $GSADUsMoonlightExe = $fixture   # any file that exists
        $script:launched = @()
        function Start-StreamMoonlight { param([string[]]$Arguments) $script:launched += ($Arguments -join ' '); 4242 }
        function Get-StreamMoonlight { @([pscustomobject]@{ Id = 1 }) }
        remote connect 6>$null
        Assert ($script:launched.Count -eq 0) 'Connect must not start a second Moonlight'
        function Get-StreamMoonlight { @() }
        $env:SSH_CONNECTION = '100.92.227.106 50000 100.119.245.116 22'
        remote connect 6>$null
        $env:SSH_CONNECTION = $null
        Assert ($script:launched.Count -eq 0) 'Connect must refuse over SSH'
        remote connect 6>$null
        Assert (($script:launched -join '|') -eq "stream $((Get-StreamPeer).Moonlight) Desktop") 'Connect must stream the Desktop of the paired peer'
        $GSADUsMoonlightExe = 'C:\GSADUs-no-such-dir\Moonlight.exe'
        $out = remote connect 6>&1 | Out-String
        Assert ($script:launched.Count -eq 1 -and $out -like '*Moonlight not found*') 'A missing Moonlight must be reported'
    }

    # -- disconnect -----------------------------------------------------------------
    & {
        $script:running = @([pscustomobject]@{ Id = 11 })
        $script:launched = @(); $script:stopped = 0
        function Get-StreamMoonlight { $script:running }
        function Start-StreamMoonlight { param([string[]]$Arguments) $script:launched += ($Arguments -join ' '); 4242 }
        function Get-Process { [CmdletBinding()] param([int]$Id) }   # the quit run already exited
        function Stop-Process {
            [CmdletBinding()] param([Parameter(ValueFromPipeline)]$InputObject, [switch]$Force)
            process { $script:stopped++ } end { $script:running = @() }
        }
        $out = remote disconnect 6>&1 | Out-String
        Assert (($script:launched -join '|') -eq "quit $((Get-StreamPeer).Moonlight)") 'Disconnect must quit cleanly first'
        Assert ($script:stopped -eq 1 -and $out -like '*Moonlight closed here*') 'Disconnect must close Moonlight here'
        $script:running = @([pscustomobject]@{ Id = 12 }); $script:launched = @(); $script:stopped = 0
        $env:SSH_CONNECTION = 'x'
        remote disconnect 6>$null
        $env:SSH_CONNECTION = $null
        Assert ($script:launched.Count -eq 0 -and $script:stopped -eq 1) 'Over SSH, disconnect must only close Moonlight'
        $script:stopped = 0
        remote disconnect 6>$null
        Assert ($script:stopped -eq 0) 'Nothing to close when Moonlight is not running'
    }

    # -- kick -----------------------------------------------------------------------
    & {
        $script:reads = 0
        function Get-SunshineStream {
            $script:reads++
            if ($script:reads -le 2) { ConvertFrom-SunshineLog -Lines $liveLog }
            else { ConvertFrom-SunshineLog -Lines ($liveLog + '[2026-10-21 19:00:00.000]: Info: CLIENT DISCONNECTED') }
        }
        $script:sent = $null
        function Invoke-StreamPeer { param([string]$Script, [int]$TimeoutSec) $script:sent = $Script; [pscustomobject]@{ Ok = $true; Output = @('1'); Error = $null } }
        $out = remote kick 6>&1 | Out-String
        Assert ($script:sent -like '*Get-Process -Name Moonlight*Stop-Process -Force*') "Kick must close the peer's Moonlight"
        Assert ($out -like '*closed 1 Moonlight*' -and $out -like '*has ended*') 'Kick must wait for Sunshine to drop the stream'

        $script:restarts = 0
        function Restart-RemoteSunshine { $script:restarts++ }
        function Invoke-StreamPeer { [pscustomobject]@{ Ok = $false; Output = @(); Error = 'ssh: connect to host gsadus-vadim port 22: Connection timed out' } }
        function Read-Host { 'y' }
        $script:reads = 0
        remote kick 6>$null
        Assert ($script:restarts -eq 1) 'SSH down: yes must restart Sunshine here'
        function Read-Host { '' }
        $script:reads = 0
        remote kick 6>$null
        Assert ($script:restarts -eq 1) 'SSH down: Enter must leave Sunshine alone'

        function Get-SunshineStream { ConvertFrom-SunshineLog -Lines $idleLog }
        $script:sent = $null
        function Invoke-StreamPeer { param([string]$Script, [int]$TimeoutSec) $script:sent = $Script; [pscustomobject]@{ Ok = $true; Output = @('0') } }
        remote kick 6>$null
        Assert ($null -eq $script:sent) 'No stream into this PC: kick must not touch the peer'
    }

    # -- restart --------------------------------------------------------------------
    & {
        function Start-Sleep { [Threading.Thread]::Sleep(25) }   # the no-fresh-start case waits out its 15 s deadline
        $script:did = @()
        function Start-Process {
            [CmdletBinding()] param($FilePath, $Verb, $ArgumentList, [switch]$Wait, $WindowStyle)
            $script:did += "$Verb $($ArgumentList -join ' ')"
        }
        function Restart-Service { [CmdletBinding()] param($Name, [switch]$Force) $script:did += "direct $Name" }
        $script:reads = 0
        function Get-SunshineStream {
            $script:reads++
            if ($script:reads -eq 1) { ConvertFrom-SunshineLog -Lines $liveLog }
            else { ConvertFrom-SunshineLog -Lines @('[2026-10-21 19:30:00.000]: Info: Sunshine version: x') }
        }
        function Test-RemoteAdmin { $false }
        $out = remote restart 6>&1 | Out-String
        Assert ($script:did[0] -like 'RunAs *Restart-Service -Name SunshineService -Force') 'Not elevated: restart through one UAC prompt'
        Assert ($out -like '*Sunshine restarted*') 'Restart must confirm a fresh Sunshine'
        function Test-RemoteAdmin { $true }
        $script:reads = 0
        remote restart 6>$null
        Assert ($script:did[-1] -eq 'direct SunshineService') 'Elevated: restart the service directly'
        function Test-RemoteAdmin { $false }
        function Start-Process { [CmdletBinding()] param($FilePath, $Verb, $ArgumentList, [switch]$Wait, $WindowStyle) throw 'The operation was canceled by the user.' }
        $script:reads = 0
        $out = remote restart 6>&1 | Out-String
        Assert ($out -like '*not restarted*canceled by the user*') 'A declined UAC prompt must be reported'
        function Get-SunshineStream { ConvertFrom-SunshineLog -Lines $liveLog }
        function Test-RemoteAdmin { $true }
        $out = remote restart 6>&1 | Out-String
        Assert ($out -like '*WARN*no fresh start*') 'A restart that never shows a fresh log must warn'
    }

    # -- SSH fails fast --------------------------------------------------------------
    & {
        function Get-StreamPeer { @{ Name = 'TEST'; Moonlight = 'TEST'; SshUser = 'nobody'; SshHost = '192.0.2.1' } }   # TEST-NET-1
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $r = Invoke-StreamPeer -Script '1' -TimeoutSec 2
        Assert (-not $r.Ok -and $r.Error -and $sw.Elapsed.TotalSeconds -lt 5) "An unreachable peer must fail by the deadline (took $($sw.Elapsed.TotalSeconds) s)"
        $sw.Restart()
        $r = Invoke-StreamPeer -Script '1'
        Assert (-not $r.Ok -and $r.Error -and $sw.Elapsed.TotalSeconds -lt 9) "ssh's own connect timeout must apply (took $($sw.Elapsed.TotalSeconds) s)"
        function Get-StreamPeer { @{ Name = 'TEST'; Moonlight = 'TEST'; SshUser = 'nobody'; SshHost = 'gsadus-remote-test.invalid' } }
        $r = Invoke-StreamPeer -Script '1'
        Assert (-not $r.Ok -and $r.Error -like '*resolve*') "A bad host name must fail with ssh's reason (got: $($r.Error))"
    }

    Write-Host 'PASS: pairing, sunshine.log reading, the SSH probe round trip, status text, menu per state and dispatch, connect/disconnect/kick/restart with mocks, SSH fail-fast.'
} finally {
    $env:SSH_CONNECTION = $savedSsh
    # Only the exact unique fixture created above is removed.
    $resolved = [IO.Path]::GetFullPath($fixture)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($tempRoot) -or (Split-Path $resolved -Leaf) -notlike 'gsadus-remote-test-*.log') {
        throw 'Unsafe fixture cleanup path'
    }
    if (Test-Path -LiteralPath $resolved) { Remove-Item -LiteralPath $resolved -Force }
}
