#requires -Version 7
<#
.SYNOPSIS
  Read-only scan of vadim@gsadus.com's mail for the Orchestrator's morning brief.

.DESCRIPTION
  Lists mail since the last scan (headers only), classifies each message with inbox-rules.json
  and prints a short report: needs you, watch, receipts, review, plus counts of routine and
  noise mail. Email content is untrusted data: nothing here follows a link or an instruction
  from a message, and subjects and snippets are redacted before they are printed.

  The only writes are behind -ForwardReceipts (owner 2026-10-14): each rule-matched receipt
  not yet carrying the "Forwarded to accounting" label is forwarded to accounting@gsadus.com
  with its attachments, then given that label so it is never forwarded twice.

.PARAMETER Since
  Scan mail newer than this time. Default: the last successful scan on this machine
  ($HOME\.gsadus\inbox-scan.json), else 2 days ago; never more than 14 days back.

.PARAMETER ForwardReceipts
  Forward new receipts to accounting@gsadus.com (see above).

.PARAMETER NoState
  Do not record this run as the last scan (for tests and look-backs).

.PARAMETER Json
  Emit the classified messages as JSON instead of the text report.

.PARAMETER Body
  Print one message's body, sanitized (links removed, codes and long numbers redacted).

.EXAMPLE
  pwsh -NoProfile -File C:\GSADUs\Tools\Inbox\inbox-scan.ps1 -ForwardReceipts
  pwsh -NoProfile -File C:\GSADUs\Tools\Inbox\inbox-scan.ps1 -Since (Get-Date).AddDays(-7) -NoState
  pwsh -NoProfile -File C:\GSADUs\Tools\Inbox\inbox-scan.ps1 -Body 1a13696487819455
#>
[CmdletBinding()]
param(
  [datetime]$Since,
  [switch]$ForwardReceipts,
  [switch]$NoState,
  [switch]$Json,
  [string]$Body
)
$ErrorActionPreference = 'Stop'
$rulesPath = Join-Path $PSScriptRoot 'inbox-rules.json'
$statePath = Join-Path $HOME '.gsadus\inbox-scan.json'

function Invoke-Gws([string[]]$GwsArgs) {
  # gws prints "Using keyring backend" on stderr; keep stdout's JSON only.
  $out = & gws @GwsArgs 2>$null | Out-String
  $start = $out.IndexOf('{')
  if ($start -lt 0) { return $null }
  try { $out.Substring($start) | ConvertFrom-Json -Depth 20 } catch { $null }
}

function Hide-Secrets([string]$Text) {
  if (-not $Text) { return $Text }
  $t = [regex]::Replace($Text, '[*xX•]{3,}[ *xX•-]*\d{4}', '[card]')             # masked card numbers
  $t = [regex]::Replace($t, '\d{3,4}(?:[- ]\d{3,4}){2,}', '[#]')                 # account and receipt numbers
  $t = [regex]::Replace($t, '\d{5,}', '[#]')                                     # codes, long numbers
  [regex]::Replace($t, '(?<![@.A-Za-z0-9])(?=[A-Za-z0-9-]*\d)(?=[A-Za-z0-9-]*[A-Za-z])[A-Za-z0-9]{6,}\b(?![@.]\w)', '[code]')
}

function Get-Address([string]$From) {
  if ($From -match '<([^>]+)>') { $matches[1].ToLower() } else { "$From".Trim().ToLower() }
}

# --- Body mode -------------------------------------------------------------------------------
if ($Body) {
  $m = Invoke-Gws @('gmail', 'users', 'messages', 'get', '--params', (@{ userId = 'me'; id = $Body; format = 'full' } | ConvertTo-Json -Compress))
  if (-not $m) { Write-Output "INBOX: could not read message $Body (gws auth? run: gws auth login)"; exit 3 }
  function Decode([string]$d) {
    $b = $d.Replace('-', '+').Replace('_', '/'); switch ($b.Length % 4) { 2 { $b += '==' } 3 { $b += '=' } }
    [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($b))
  }
  function Find-Text($part) {
    if ($part.mimeType -eq 'text/plain' -and $part.body.data) { return @{ plain = Decode $part.body.data } }
    $html = if ($part.mimeType -eq 'text/html' -and $part.body.data) { Decode $part.body.data }
    foreach ($p in $part.parts) { $r = Find-Text $p; if ($r.plain) { return $r }; if ($r.html -and -not $html) { $html = $r.html } }
    @{ html = $html }
  }
  $r = Find-Text $m.payload
  $t = if ($r.plain) { $r.plain } else {
    $x = [regex]::Replace("$($r.html)", '(?is)<(style|script)[^>]*>.*?</\1>', ' ')
    $x = [regex]::Replace($x, '(?i)<br\s*/?>|</p>|</tr>|</div>|</h\d>', "`n")
    [Net.WebUtility]::HtmlDecode([regex]::Replace($x, '<[^>]+>', ' '))
  }
  $t = [regex]::Replace($t, 'https?://\S+', '[link]')
  $t = [regex]::Replace((Hide-Secrets $t), '[ \t\u00a0\u200c\u034f]+', ' ')
  $t = [regex]::Replace($t, '(\s*\n\s*)+', "`n").Trim()
  $h = @{}; foreach ($x in $m.payload.headers) { $h[$x.name] = $x.value }
  Write-Output "From: $($h['From'])"
  Write-Output "Date: $($h['Date'])"
  Write-Output "Subject: $(Hide-Secrets $h['Subject'])"
  Write-Output ''
  Write-Output ($t.Length -gt 4000 ? $t.Substring(0, 4000) + ' …' : $t)
  exit 0
}

# --- Auth check -------------------------------------------------------------------------------
$profile_ = Invoke-Gws @('gmail', 'users', 'getProfile', '--params', '{"userId":"me"}')
if (-not $profile_.emailAddress) { Write-Output 'INBOX: gws needs re-login (run: gws auth login). Inbox not scanned.'; exit 3 }

# --- Window -----------------------------------------------------------------------------------
$now = [DateTimeOffset]::Now
if (-not $PSBoundParameters.ContainsKey('Since')) {
  $Since = $now.AddDays(-2).LocalDateTime
  if (Test-Path $statePath) {
    try { $Since = ([DateTimeOffset]::Parse((Get-Content $statePath -Raw | ConvertFrom-Json).lastRun)).LocalDateTime } catch { }
  }
}
$floor = $now.AddDays(-14).LocalDateTime
if ($Since -lt $floor) { $Since = $floor }
$after = [DateTimeOffset]::new($Since).ToUnixTimeSeconds()

# --- Rules and labels -------------------------------------------------------------------------
$config = Get-Content $rulesPath -Raw | ConvertFrom-Json
$labels = (Invoke-Gws @('gmail', 'users', 'labels', 'list', '--params', '{"userId":"me"}')).labels
$labelName = @{}; foreach ($l in $labels) { $labelName[$l.id] = $l.name }
$forwardedLabelId = ($labels | Where-Object name -eq $config.forwardedLabel).id

# --- List and fetch ---------------------------------------------------------------------------
$ids = [Collections.Generic.List[string]]::new()
$token = $null
do {
  $p = @{ userId = 'me'; q = "after:$after -in:sent -in:chats -in:drafts"; maxResults = 500 }
  if ($token) { $p.pageToken = $token }
  $page = Invoke-Gws @('gmail', 'users', 'messages', 'list', '--params', ($p | ConvertTo-Json -Compress))
  foreach ($m in $page.messages) { $ids.Add($m.id) }
  $token = $page.nextPageToken
} while ($token)

$fetched = $ids | ForEach-Object -ThrottleLimit 6 -Parallel {
  $p = @{ userId = 'me'; id = $_; format = 'metadata'; metadataHeaders = @('From', 'Subject', 'Date', 'List-Unsubscribe') } | ConvertTo-Json -Compress
  $out = & gws gmail users messages get --params $p 2>$null | Out-String
  $i = $out.IndexOf('{')
  if ($i -ge 0) { try { $out.Substring($i) | ConvertFrom-Json -Depth 20 } catch { } }
}

# --- Classify ---------------------------------------------------------------------------------
function Test-Rule($rule, $msg) {
  if ($rule.from -and $msg.from -notmatch $rule.from) { return $false }
  if ($rule.subject -and $msg.subject -notmatch $rule.subject) { return $false }
  if ($rule.snippet -and $msg.snippet -notmatch $rule.snippet) { return $false }
  if ($rule.label -and -not ($msg.labels -contains $rule.label)) { return $false }
  $true
}

$items = foreach ($m in $fetched) {
  $h = @{}; foreach ($x in $m.payload.headers) { $h[$x.name] = $x.value }
  $labelList = @($m.labelIds | ForEach-Object { $_; if ($labelName[$_]) { $labelName[$_] } })
  $msg = [pscustomobject]@{
    id = $m.id; time = [DateTimeOffset]::FromUnixTimeMilliseconds([int64]$m.internalDate).LocalDateTime
    from = Get-Address $h['From']; subject = "$($h['Subject'])"; snippet = "$($m.snippet)"
    labels = $labelList; inInbox = $m.labelIds -contains 'INBOX'; unsub = [bool]$h['List-Unsubscribe']
    class = $null; kind = $null; note = $null; forwarded = $labelList -contains $config.forwardedLabel
  }
  $rule = $config.rules | Where-Object { Test-Rule $_ $msg } | Select-Object -First 1
  if ($rule) { $msg.class = $rule.class; $msg.kind = $rule.kind; $msg.note = $rule.note }
  elseif ($msg.unsub) { $msg.class = 'noise'; $msg.kind = 'newsletter not yet filtered (unsubscribe candidate)' }
  else { $msg.class = 'review'; $msg.kind = 'unmatched: judge from subject and snippet' }
  $msg
}
$items = @($items | Sort-Object time)

# --- Forward receipts -------------------------------------------------------------------------
$forwardLog = @()
if ($ForwardReceipts) {
  foreach ($r in $items | Where-Object { $_.class -eq 'receipt' -and -not $_.forwarded }) {
    if (-not $forwardedLabelId) {
      $created = Invoke-Gws @('gmail', 'users', 'labels', 'create', '--params', '{"userId":"me"}', '--json',
        (@{ name = $config.forwardedLabel; labelListVisibility = 'labelShow'; messageListVisibility = 'show' } | ConvertTo-Json -Compress))
      $forwardedLabelId = $created.id
    }
    $sent = Invoke-Gws @('gmail', '+forward', '--message-id', $r.id, '--to', $config.receiptForwardTo)
    if ($sent.id) {
      $null = Invoke-Gws @('gmail', 'users', 'messages', 'modify', '--params', (@{ userId = 'me'; id = $r.id } | ConvertTo-Json -Compress),
        '--json', (@{ addLabelIds = @($forwardedLabelId) } | ConvertTo-Json -Compress))
      $r.forwarded = $true
      $forwardLog += "forwarded: $($r.kind), $($r.time.ToString('MM-dd'))"
    } else {
      $forwardLog += "FORWARD FAILED: $($r.kind), $($r.time.ToString('MM-dd')) (id $($r.id))"
    }
  }
}

if (-not $NoState) {
  New-Item -ItemType Directory -Force -Path (Split-Path $statePath) | Out-Null
  @{ lastRun = $now.ToString('o') } | ConvertTo-Json | Set-Content -Path $statePath -Encoding utf8
}

# --- Report -----------------------------------------------------------------------------------
if ($Json) {
  $items | Select-Object id, time, from, @{ n = 'subject'; e = { Hide-Secrets $_.subject } }, @{ n = 'snippet'; e = { Hide-Secrets $_.snippet } },
    class, kind, note, inInbox, forwarded | ConvertTo-Json -Depth 4
  exit 0
}
function Format-Lines($group) {
  # One line per sender + subject; repeats collapse to "N× first → last".
  foreach ($g in $group | Group-Object { "$($_.from)|$($_.subject)" } | Sort-Object { $_.Group[0].time }) {
    $first = $g.Group[0]; $last = $g.Group[-1]
    $when = if ($g.Count -gt 1) { '{0}× {1} → {2}' -f $g.Count, $first.time.ToString('MM-dd HH:mm'), $last.time.ToString('MM-dd HH:mm') } else { $first.time.ToString('MM-dd HH:mm') }
    $line = '  {0}  {1}  {2}  [{3}]' -f $when, $first.from, (Hide-Secrets $first.subject), $first.kind
    if ($first.note) { $line += " - $($first.note)" }
    $line
  }
}
Write-Output ('INBOX since {0} ({1} messages)' -f $Since.ToString('yyyy-MM-dd HH:mm'), $items.Count)
foreach ($section in @(@('needs', 'NEEDS YOU'), @('watch', 'WATCH'), @('review', 'REVIEW (unmatched)'), @('receipt?', 'POSSIBLE RECEIPTS (confirm, then forward)'))) {
  $group = @($items | Where-Object class -eq $section[0])
  if ($group.Count) { Write-Output "$($section[1]) ($($group.Count))"; Format-Lines $group }
}
$receipts = @($items | Where-Object class -eq 'receipt')
if ($receipts.Count) {
  $pending = @($receipts | Where-Object { -not $_.forwarded })
  Write-Output ('RECEIPTS ({0}; {1} forwarded to {2})' -f $receipts.Count, ($receipts.Count - $pending.Count), $config.receiptForwardTo)
  $forwardLog | ForEach-Object { "  $_" }
  $pending | ForEach-Object { (Format-Lines @($_)) + ' - not forwarded yet' }
}
foreach ($c in 'info', 'noise') {
  $group = @($items | Where-Object class -eq $c)
  if ($group.Count) {
    $summary = ($group | Group-Object kind | Sort-Object Count -Descending | ForEach-Object {
      $n = "$($_.Count) $($_.Name)"
      if ($_.Name -like 'newsletter not yet filtered*') { $n += ': ' + (($_.Group.from | Sort-Object -Unique) -join ', ') }
      $n }) -join ' · '
    Write-Output ('{0} ({1}): {2}' -f $c.ToUpper(), $group.Count, $summary)
  }
}
