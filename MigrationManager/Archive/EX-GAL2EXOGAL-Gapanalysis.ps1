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

if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}
# Run this on your on-premises Exchange Management Shell
$OnPremGAL = Get-Recipient -ResultSize Unlimited | Select-Object DisplayName,PrimarySmtpAddress,RecipientType

Connect-ExchangeOnline -AppId $appid -Organization $Tenantdomain -CertificateThumbprint $certThumbprint

# Connect to Exchange Online PowerShell session first
$EXOGAL = Get-Recipient -ResultSize Unlimited | Select-Object DisplayName,PrimarySmtpAddress,RecipientType
# Assuming $OnPremGAL and $EXOGAL have been collected as before

# Adding a custom property to each object to indicate the source environment
$OnPremGAL = $OnPremGAL | ForEach-Object { Add-Member -InputObject $_ -NotePropertyName "Environment" -NotePropertyValue "On-Premises" -PassThru }
$EXOGAL = $EXOGAL | ForEach-Object { Add-Member -InputObject $_ -NotePropertyName "Environment" -NotePropertyValue "Exchange Online" -PassThru }

# Adding a custom property to each object to indicate the source environment
$OnPremGAL = $OnPremGAL | ForEach-Object { Add-Member -InputObject $_ -NotePropertyName "Environment" -NotePropertyValue "On-Premises" -PassThru }
$EXOGAL = $EXOGAL | ForEach-Object { Add-Member -InputObject $_ -NotePropertyName "Environment" -NotePropertyValue "Exchange Online" -PassThru }

Write-Host = "Onpremises recipient count = $onPremGal.Count" -ForeGroundColor Yellow
Write-Host = "Exchange Online recipient count = $EXOGAL.Count" -ForeGroundColor Yellow

# Now perform the comparison including the Environment property
$Differences = Compare-Object -ReferenceObject $OnPremGAL -DifferenceObject $EXOGAL -Property PrimarySmtpAddress

# Filter out only the differences
$ActualDifferences = $Differences | ForEach-Object {
    if ($_.SideIndicator -eq "<=") {
        $side = "Only in On-Premises"
    } else {
        $side = "Only in Exchange Online"
    }
    # Add or modify properties as needed to include the side indication
    $_ | Add-Member -NotePropertyName "LocationDifference" -NotePropertyValue $side -PassThru
}

# Export the findings to a CSV file
Set-Location $outputPath
$ActualDifferences | Export-Csv -Path "GALGapAnalysisReport.csv" -NoTypeInformation

$actualDifferences