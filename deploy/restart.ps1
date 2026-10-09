# Restarts the Travel Portal (run elevated). Use after editing .env or updating the code.
#   powershell -ExecutionPolicy Bypass -File .\deploy\restart.ps1
param([int]$Port = 8080)

Stop-ScheduledTask -TaskName 'TravelPortal' -ErrorAction SilentlyContinue
# Stopping the task ends run.cmd but not the Python process it started, so free the port explicitly
Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue |
    ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }
Start-Sleep -Seconds 1
Start-ScheduledTask -TaskName 'TravelPortal'
Start-Sleep -Seconds 4
try {
    $r = Invoke-WebRequest -UseBasicParsing "http://127.0.0.1:$Port/healthz" -TimeoutSec 10
    Write-Host "Portal restarted (health: $($r.Content))" -ForegroundColor Green
} catch {
    Write-Warning "Portal did not answer on port $Port. Check logs\portal.log"
}
