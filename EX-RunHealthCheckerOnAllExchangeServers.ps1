# Define Health Checker script path
$HealthCheckerScriptPath = $env:TEMP + "\HealthChecker.ps1"  # Update with actual path
Set-Location $env:TEMP
# Download HealthChecker script if not present
if (-not (Test-Path $HealthCheckerScriptPath)) {
    $HealthCheckerScriptUrl = "https://github.com/microsoft/CSS-Exchange/releases/latest/download/HealthChecker.ps1"
    Invoke-WebRequest -Uri $HealthCheckerScriptUrl -OutFile $HealthCheckerScriptPath
}

# Add Exchange snap-in if not already loaded
if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}
# Fetch Exchange Servers
$ExchangeServers = Get-ExchangeServer | Select-Object -ExpandProperty Name

# Run Health Checker on all Exchange servers
.\HealthChecker.ps1 -Server $ExchangeServers

# Optional: Build HTML report for all servers
.\HealthChecker.ps1 -BuildHtmlServersReport
