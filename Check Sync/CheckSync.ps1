# Define the remote PC details
$remotePC = "GSADUS-Vadim" # or use the IP address
$user = "Vadim"

Write-Host "--- Checking Python Sync Status ---" -ForegroundColor Cyan

# 1. Get local version
$localVer = (python --version 2>&1).ToString().Trim()
Write-Host "Local PC (VG-Home): $localVer"

# 2. Get remote version via SSH
# This assumes you have SSH keys set up; otherwise, it will prompt for a password
$remoteVer = (ssh "$user@$remotePC" "python --version" 2>$null).Trim()

if ([string]::IsNullOrWhiteSpace($remoteVer)) {
    Write-Host "Error: Could not connect to $remotePC" -ForegroundColor Red
} else {
    Write-Host "Remote PC ($remotePC): $remoteVer"

    # 3. Compare
    if ($localVer -eq $remoteVer) {
        Write-Host "SUCCESS: Environments are perfectly synced!" -ForegroundColor Green
    } else {
        Write-Host "WARNING: Version mismatch detected!" -ForegroundColor Yellow
        Write-Host "Action required: Update one of the machines to match $localVer"
    }
}