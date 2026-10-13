# Adds the IIS reverse-proxy rule that sends http://<host>:9000/dashboard/travel/... to the Travel Portal.
# Run on the server (10.10.30.90) in an elevated PowerShell, from the project folder, AFTER deploy\install.ps1:
#   powershell -ExecutionPolicy Bypass -File .\deploy\add_iis_rule.ps1
# Optional: -PublicPort 9000 (the IIS site's port).  Safe to re-run: the rule is replaced, not duplicated.
param([int]$PublicPort = 9000)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$RuleName = 'Travel Portal'

function Get-EnvValue([string]$Name, [string]$Default = '') {
    $line = Select-String -Path "$Root\.env" -Pattern "^$Name=(.*)$" | Select-Object -First 1
    if ($line) { return $line.Matches[0].Groups[1].Value.Trim() } else { return $Default }
}

$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $admin) { throw 'Run this script from an elevated (Run as administrator) PowerShell.' }

$portalPort = [int](Get-EnvValue 'PORT' '8090')
$prefix = (Get-EnvValue 'APP_PREFIX' '/dashboard/travel').Trim('/')
if (-not $prefix) { throw 'APP_PREFIX is empty in .env - set it to /dashboard/travel' }
$matchUrl = '^' + [regex]::Escape($prefix).Replace('\/', '/') + '(/.*)?$'
$target = "http://127.0.0.1:$portalPort/$prefix{R:1}"

Import-Module WebAdministration

# --- Prerequisites: URL Rewrite + ARR proxy -------------------------------------
if (-not (Get-WebConfiguration -PSPath 'MACHINE/WEBROOT/APPHOST' -Filter 'system.webServer/rewrite/rules' -ErrorAction SilentlyContinue)) {
    throw 'IIS URL Rewrite module is not installed (https://www.iis.net/downloads/microsoft/url-rewrite).'
}
$proxy = Get-WebConfiguration -PSPath 'MACHINE/WEBROOT/APPHOST' -Filter 'system.webServer/proxy' -ErrorAction SilentlyContinue
if (-not $proxy) { throw 'IIS Application Request Routing (ARR) is not installed.' }
if (-not $proxy.enabled) {
    throw ('The ARR proxy is disabled server-wide. Enable it in IIS Manager > server > Application Request Routing ' +
           'Cache > Server Proxy Settings > "Enable proxy" (check with IT first), then re-run this script.')
}
# Server-wide ARR settings are shared with the Mingalaba app, so they are only READ here, never changed.
# The portal needs ARR to pass the visitor's IP (X-Forwarded-For) for User Management's office-network check.
$xff = (Get-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' -Filter 'system.webServer/proxy' -Name xForwardedForHeaderName -ErrorAction SilentlyContinue).Value

# --- Find the site on the public port --------------------------------------------
$site = Get-Website | Where-Object { $_.Bindings.Collection | Where-Object { $_.bindingInformation -match ":$PublicPort`:" } } |
    Select-Object -First 1
if (-not $site) { throw "No IIS site is bound to port $PublicPort." }
Write-Host "IIS site on port ${PublicPort}: '$($site.Name)'"

# --- Backup ---------------------------------------------------------------------
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = "TravelPortal-before-rule-$stamp"
Backup-WebConfiguration -Name $backup | Out-Null           # applicationHost.config (restore: Restore-WebConfiguration -Name ...)
$siteConfig = Join-Path ([Environment]::ExpandEnvironmentVariables($site.physicalPath)) 'web.config'
if (Test-Path $siteConfig) { Copy-Item $siteConfig "$siteConfig.bak-$stamp" }
Write-Host "Backed up IIS configuration ($backup) and the site's web.config."

# --- Where do the existing proxy rules live? --------------------------------------
# Global (server-wide) rules run BEFORE site rules, so if Mingalaba is forwarded by a global rule, ours must be global too.
$globalRules = @(Get-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' -Filter 'system.webServer/rewrite/globalRules' -Name '.' -ErrorAction SilentlyContinue).Collection |
    Where-Object { $_.name -ne $RuleName }
$useGlobal = [bool]$globalRules
if ($useGlobal) {
    $psPath = 'MACHINE/WEBROOT/APPHOST'
    $filter = 'system.webServer/rewrite/globalRules'
    Write-Host "Server-wide rewrite rules found ($($globalRules.Count)); adding '$RuleName' at the top of the global rules."
} else {
    $psPath = "IIS:\Sites\$($site.Name)"
    $filter = 'system.webServer/rewrite/rules'
    Write-Host "Adding '$RuleName' at the top of the site's rewrite rules."
}

# --- Replace any earlier copy, then insert as the FIRST rule ------------------------
$existing = @(Get-WebConfigurationProperty -PSPath $psPath -Filter $filter -Name '.' -ErrorAction SilentlyContinue).Collection |
    Where-Object { $_.name -eq $RuleName }
if ($existing) {
    Remove-WebConfigurationProperty -PSPath $psPath -Filter $filter -Name '.' -AtElement @{ name = $RuleName }
}
Add-WebConfigurationProperty -PSPath $psPath -Filter $filter -Name '.' -AtIndex 0 -Value @{ name = $RuleName; stopProcessing = 'True' }
$rule = "$filter/rule[@name='$RuleName']"
Set-WebConfigurationProperty -PSPath $psPath -Filter "$rule/match" -Name url -Value $matchUrl
Set-WebConfigurationProperty -PSPath $psPath -Filter "$rule/action" -Name type -Value 'Rewrite'
Set-WebConfigurationProperty -PSPath $psPath -Filter "$rule/action" -Name url -Value $target
if ($useGlobal) {
    # Global rules see every site: only act on requests to the public port
    Add-WebConfigurationProperty -PSPath $psPath -Filter "$rule/conditions" -Name '.' -Value @{ input = '{SERVER_PORT}'; pattern = "^$PublicPort$" }
}

Write-Host ""
Write-Host "Rule '$RuleName' added: /$prefix/...  ->  $target" -ForegroundColor Green
if (-not $xff) {
    Write-Warning ('ARR is not sending X-Forwarded-For, so the portal cannot see visitor IPs and User Management will ' +
        'refuse everyone. In IIS Manager > server > Application Request Routing Cache > Server Proxy Settings, tick ' +
        '"Preserve client IP in the following header" (X-Forwarded-For). Check with IT first: this setting is server-wide.')
}

# --- Verify through IIS -----------------------------------------------------------
try {
    $r = Invoke-WebRequest -UseBasicParsing "http://localhost:$PublicPort/$prefix/healthz" -TimeoutSec 10 -MaximumRedirection 0
    Write-Host "Check through IIS: http://localhost:$PublicPort/$prefix/healthz -> $($r.Content)" -ForegroundColor Green
} catch {
    Write-Warning "IIS did not return the portal's health check. Is the TravelPortal task running (deploy\install.ps1)? Details: $($_.Exception.Message)"
}
