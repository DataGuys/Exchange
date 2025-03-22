#Exchange Migration Script From Go.
#Prepare Admin Station

#Install Exchange Management Tools Pre-Req
Enable-WindowsOptionalFeature -Online -FeatureName IIS-ManagementScriptingTools,IIS-ManagementScriptingTools,IIS-IIS6ManagementCompatibility,IIS-LegacySnapIn,IIS-ManagementConsole,IIS-Metabase,IIS-WebServerManagementTools,IIS-WebServerRole
#Install Visual C++ 2012
https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe
#Install Exchange Management Tools
#Elevated CMD
setup.exe /mode:install /role:managementtools /IAcceptExchangeServerLicenseTerms_DiagnosticDataOn

#PowerShell ISE as Administrator
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
#As long as the user running the PowerShell ISE window (Admin account) has Organization Management Role 
#no need to remote connect to Exchange as long as the Management Tools are installed and imported using the Add-PSSnapin command above.

Get-ExchangeServer

#Get All Exchange Server SSL Certificates Report. Import the CSV files into the Dashboard Workbook.
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

$ExchangeServers = Get-ExchangeServer
$Results = @()
foreach ($server in $ExchangeServers) {
    Write-Host "Checking server: $($server.Name)" -ForegroundColor Cyan

    # Retrieve Exchange SSL Certificates for the current server
    $Certificates = Get-ExchangeCertificate -Server $Server.Name

    $Results += $Certificates
}
cd c:\temp
$Results | Export-Csv .\EXCHSSLCerts.csv -NoTypeInformation
