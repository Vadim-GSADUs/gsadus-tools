# Owner confirmation for `helpdesk approve` / `helpdesk reject`: a topmost desktop window the
# owner clicks, so an approval is always a deliberate owner act. Claude Code and Codex permission
# rules match command text as typed, so they can't guard this tool; a window works the same in
# both harnesses and in every permission mode. It stops accidental and unprompted approvals, but
# it is not a security boundary: an agent set on going around it could automate the UI or write
# to the database directly. The helpdesk skill forbids both. The hard boundary comes with the
# Chat Approve button, which is checked against the owner's Google account (spec, "Later").
# The prompt text arrives in $env:HELPDESK_CONFIRM_TEXT, never on the command line.
# Exit codes: 0 = confirmed, 1 = declined, 2 = timed out.
param(
    [Parameter(Mandatory)][ValidateSet('Approve', 'Reject')][string]$Action,
    [int]$TimeoutSeconds = 240
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "GSADUs helpdesk: ${Action}?"
$form.TopMost = $true
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.ClientSize = New-Object System.Drawing.Size(600, 340)

$body = New-Object System.Windows.Forms.TextBox
$body.Multiline = $true
$body.ReadOnly = $true
$body.ScrollBars = 'Vertical'
$body.Text = ($env:HELPDESK_CONFIRM_TEXT -replace "`r?`n", "`r`n")
$body.Location = New-Object System.Drawing.Point(12, 12)
$body.Size = New-Object System.Drawing.Size(576, 270)
$body.TabStop = $false
$form.Controls.Add($body)

$yes = New-Object System.Windows.Forms.Button
$yes.Text = $Action
$yes.Location = New-Object System.Drawing.Point(388, 296)
$yes.Size = New-Object System.Drawing.Size(96, 30)
$yes.DialogResult = [System.Windows.Forms.DialogResult]::Yes
$form.Controls.Add($yes)

$no = New-Object System.Windows.Forms.Button
$no.Text = 'Cancel'
$no.Location = New-Object System.Drawing.Point(492, 296)
$no.Size = New-Object System.Drawing.Size(96, 30)
$no.DialogResult = [System.Windows.Forms.DialogResult]::No
$form.Controls.Add($no)

# No AcceptButton: a stray Enter never confirms. Esc and the default focus both mean Cancel.
$form.CancelButton = $no
$form.Add_Shown({ $form.Activate(); $no.Focus() })

$timedOut = $false
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = [Math]::Max(1, $TimeoutSeconds) * 1000
$timer.Add_Tick({ $script:timedOut = $true; $timer.Stop(); $form.Close() })
$timer.Start()

$result = $form.ShowDialog()
$timer.Dispose()
$form.Dispose()
if ($timedOut) { exit 2 }
if ($result -eq [System.Windows.Forms.DialogResult]::Yes) { exit 0 }
exit 1
