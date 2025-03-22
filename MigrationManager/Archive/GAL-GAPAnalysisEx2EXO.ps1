# Define the output path
$outputPath = "C:\ExchangeMigration"
#Start-Transcript -OutputDirectory $outputPath
#Install-Module -Name SharePointPnPPowerShellOnline -Scope AllUsers
Import-Module ExchangeOnlineManagement
Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue
Import-Module CredentialManager
$certThumbprint = "5E79521C7F31B61A5870F20BE06CF78E3947655C"
$appid = '0609911f-6d2a-47a7-a49c-d41535b2d4a3'
$Tenantdomain = 'polsinelli.onmicrosoft.com'

Connect-ExchangeOnline -AppId $appid -Organization $Tenantdomain -CertificateThumbprint $certThumbprint
# Ensure the Exchange Management PowerShell snap-in is loaded for on-premises Exchangeac
if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}

# Collect Global Address List (GAL) from on-premises Exchange
$OnPremGAL = Get-Recipient -ResultSize Unlimited | Select-Object DisplayName,PrimarySmtpAddress,RecipientType

# Connect to Exchange Online using certificate-based authentication
Connect-ExchangeOnline -AppId $appid -Organization $Tenantdomain -CertificateThumbprint $certThumbprint

# Collect Global Address List (GAL) from Exchange Online
$EXOGAL = Get-Recipient -ResultSize Unlimited | Select-Object DisplayName,PrimarySmtpAddress,RecipientType

# Add a custom property to each object to indicate the source environment
$OnPremGAL = $OnPremGAL | ForEach-Object {
    $_ | Add-Member -NotePropertyName "Environment" -NotePropertyValue "On-Premises" -PassThru
}

$EXOGAL = $EXOGAL | ForEach-Object {
    $_ | Add-Member -NotePropertyName "Environment" -NotePropertyValue "Exchange Online" -PassThru
}

# Display the recipient counts for both environments
Write-Host "On-premises recipient count = $($OnPremGAL.Count)" -ForegroundColor Yellow
Write-Host "Exchange Online recipient count = $($EXOGAL.Count)" -ForegroundColor Yellow

# Compare the objects, including the Environment property, to find differences
$Differences = Compare-Object -ReferenceObject $OnPremGAL -DifferenceObject $EXOGAL -Property PrimarySmtpAddress

# Process and mark the differences with appropriate location indicators
$ActualDifferences = $Differences | ForEach-Object {
    $side = if ($_.SideIndicator -eq "<=") { "Only in On-Premises" } else { "Only in Exchange Online" }
    $_ | Add-Member -NotePropertyName "LocationDifference" -NotePropertyValue $side -PassThru
}

# Export the findings to a CSV file
$ActualDifferences | Export-Csv -Path "$outputPath\GALGapAnalysisReport.csv" -NoTypeInformation

# Output the actual differences variable for display or further processing
$ActualDifferences
