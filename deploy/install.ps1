# Installs the Travel Accommodation Portal on this Windows server (10.10.30.90).
# Run from an elevated PowerShell in the project folder:
#   powershell -ExecutionPolicy Bypass -File .\deploy\install.ps1
# Safe to re-run: it updates packages and re-registers the task.

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$TaskName = 'TravelPortal'
Set-Location $Root

$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $admin) { throw 'Run this script from an elevated (Run as administrator) PowerShell.' }

# --- Prerequisites ---------------------------------------------------------
$python = Get-Command python -ErrorAction SilentlyContinue
if (-not $python) { throw 'Python 3.10+ not found. Install it from python.org (tick "Add to PATH", install for all users).' }
& python -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)"
if ($LASTEXITCODE -ne 0) { throw 'Python 3.10 or newer is required.' }

if (-not (Get-OdbcDriver -Name 'ODBC Driver 18 for SQL Server' -ErrorAction SilentlyContinue)) {
    throw 'ODBC Driver 18 for SQL Server is not installed: https://learn.microsoft.com/sql/connect/odbc/download-odbc-driver-for-sql-server'
}

if (-not (Test-Path "$Root\.env")) { throw ".env not found in $Root. Copy it here and fill in the placeholders." }
$placeholders = Select-String -Path "$Root\.env" -Pattern '^[A-Z_]+=.*<[A-Z_]+' | ForEach-Object { ($_.Line -split '=')[0] }
if ($placeholders) { throw ".env still has placeholders: $($placeholders -join ', ')" }

$port = 8080
$portLine = Select-String -Path "$Root\.env" -Pattern '^PORT=(\d+)' | Select-Object -First 1
if ($portLine) { $port = [int]$portLine.Matches[0].Groups[1].Value }

# --- Python environment ----------------------------------------------------
if (-not (Test-Path "$Root\.venv\Scripts\python.exe")) {
    Write-Host 'Creating virtual environment...'
    & python -m venv "$Root\.venv"
}
Write-Host 'Installing packages...'
& "$Root\.venv\Scripts\python.exe" -m pip install --disable-pip-version-check -q -r "$Root\requirements.txt"
if ($LASTEXITCODE -ne 0) { throw 'pip install failed.' }

Write-Host 'Checking database connection...'
& "$Root\.venv\Scripts\python.exe" -c "import app; c = app.get_conn(); c.cursor().execute('SELECT TOP 1 1 FROM TravelRequests'); print('Database OK')"
if ($LASTEXITCODE -ne 0) { throw 'Database check failed (see the error above). Check DB_* values in .env.' }

# --- Firewall: allow the internal network to reach the portal --------------
$rule = 'Travel Portal (TCP ' + $port + ')'
Get-NetFirewallRule -DisplayName 'Travel Portal (TCP *' -ErrorAction SilentlyContinue | Remove-NetFirewallRule
New-NetFirewallRule -DisplayName $rule -Direction Inbound -Protocol TCP -LocalPort $port `
    -Action Allow -Profile Domain, Private | Out-Null
Write-Host "Firewall rule added: $rule (Domain/Private networks)"

# --- Scheduled task: start at boot, restart on failure ---------------------
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}
# Stopping the task does not end the Python process it started, so free the port explicitly
Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue |
    ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }
$action = New-ScheduledTaskAction -Execute "$Root\deploy\run.cmd" -WorkingDirectory $Root
$trigger = New-ScheduledTaskTrigger -AtStartup
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1)
$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
    -Principal $principal -Description 'Travel Accommodation Portal (Waitress)' | Out-Null
Start-ScheduledTask -TaskName $TaskName

# --- Verify -----------------------------------------------------------------
Start-Sleep -Seconds 5
try {
    $r = Invoke-WebRequest -UseBasicParsing "http://127.0.0.1:$port/healthz" -TimeoutSec 10
    Write-Host "Portal is running: http://10.10.30.90:$port  (health: $($r.Content))" -ForegroundColor Green
} catch {
    Write-Warning "Portal did not answer on port $port. Check $Root\logs\portal.log"
}
