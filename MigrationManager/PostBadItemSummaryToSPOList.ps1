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
$tenantid = 'c824b048-0922-4f96-87fc-bf920e1a756d'
$creds = Get-StoredCredential -Target "Batcher"
#Scan SharePoint List for Updates
Connect-ExchangeOnline -AppId $appid -Organization $Tenantdomain -CertificateThumbprint $certThumbprint
$MoveRequests = Get-MoveRequest -ResultSize Unlimited

$MoveReport = foreach ($Move in $MoveRequests) {
    $Statistics = Get-MoveRequestStatistics -Identity $Move.Identity -IncludeReport
    foreach ($item in $Statistics.Report.BadItems) {
        [PSCustomObject]@{
            User = $Statistics.Identity.ToString() # Convert to string for consistency
            Kind = $item.Kind
            Folder = $item.Folder
        }
    }
}

# Group by User, then summarize
$GroupedByUser = $MoveReport | Group-Object -Property User

Connect-PnPOnline -Url "https://polsinelli.sharepoint.com/sites/ExchangeOnlineProject" -ClientId $appid -Thumbprint $certThumbprint -Tenant $tenantId -WarningAction SilentlyContinue

foreach ($group in $GroupedByUser) {
    $user = $group.Name
    $BadItemsSummaryArray = $group.Group | Group-Object -Property Kind | ForEach-Object {
        "$($_.Name): $($_.Group.Count)"
    }
    $BadItemsSummary = $BadItemsSummaryArray -join ', '

    $ExchangeGuid = (Get-MailUser -Identity $user).ExchangeGuid
    $item = Get-PnPListItem -List "OnPremMailboxes" | Where-Object {$_.FieldValues.Title -eq $ExchangeGuid}

    if ($item) {
        Set-PnPListItem -List "OnPremMailboxes" -Identity $item.Id -Values @{"BadItemSummary" = $BadItemsSummary}
    }
}

