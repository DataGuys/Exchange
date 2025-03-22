# Import Exchange Online Management module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
$UserCredential = Get-Credential
Connect-ExchangeOnline -Credential $UserCredential

# Set date range for the report (last 180 days)
$EndDate = Get-Date
$StartDate = $EndDate.AddDays(-180)

# Get admin activity logs from the Unified Audit Log
$AdminLogs = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -RecordType ExchangeAdmin
$AdminLogs | ForEach-Object { 
$_.Auditdata | ConvertFrom-Json } | Export-CSV .\ELG-ExchangeAdminActivityReport5-7-2023-180days-AuditDataJSONConverted.csv -NoTypeInformation

# Export the results to a CSV file
$AdminLogs | Export-Csv -Path ".\ExchangeAdminActivityReport.csv" -NoTypeInformation

# Disconnect from Exchange Online
Disconnect-ExchangeOnline

# Display a message that the export is complete
Write-Host "Admin activity report generated successfully and saved as ExchangeAdminActivityReport.csv"
