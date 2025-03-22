# Define the output path
#$outputPath = "C:\ExchangeMigration"
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

Connect-ExchangeOnline -AppId $appid -Organization $Tenantdomain -CertificateThumbprint $certThumbprint
#Get Move Statistics and email updates to the migration team
# Retrieve move request statistics and enrich them with PrimarySmtpAddress
$MoveStats = Get-MoveRequest
Get-MigrationUser
$MoveStats | ForEach-Object {
    
    $mailbox = Get-Mailuser -Identity $_.DisplayName
    $stats = Get-MoveRequestStatistics -Identity $mailbox.PrimarySmtpAddress
    
    
    # Create a custom PSObject to include the necessary properties along with PrimarySmtpAddress
    $output = New-Object PSObject -Property @{
        DisplayName = $stats.DisplayName
        BatchName = $stats.BatchName
        Status = $stats.Status
        PercentComplete = $stats.PercentComplete
        DataConsistencyScore = $stats.DataConsistencyScore
        BadItemsEncountered = $stats.BadItemsEncountered
        LargeItemsEncountered = $stats.LargeItemsEncountered
        PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
        ExchangeGuid = $mailbox.ExchangeGuid
    }

    # Output the custom object
    $output
}

# Now $MoveStats will include PrimarySmtpAddress for each mailbox

$MoveStatsHTML = $output | ConvertTo-Html -As Table | Out-String
$output
Disconnect-ExchangeOnline -Confirm:$false

if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}
# Define the mailboxes to scan
$mailboxesToScan = @()
$mailboxesToScan = $MoveStats | Where-Object {$_.DataConsistencyScore -like "Investigate"}
# Array to hold the report
$invalidPermissionsReport = @()
foreach ($mailboxToScan in $mailboxesToScan) {
$userStats = Get-MoverequestStatistics -Identity $mailboxToScan.DisplayName
$userStats.SkippedItems | ft -a Subject, Sender, DateSent, ScoringClassifications
# Get mailbox permissions
$fullAccessPermissions = Get-MailboxPermission -Owner $mailboxToScan.DisplayName | Where-Object { $_.IsInherited -eq $false -and $_.User -like "NT AUTHORITY\SELF" -eq $false }
$sendAsPermissions = Get-ADUser -Filter {Enabled -eq $true} -Properties * | Where-Object { $_.msExchDelegateListLink -ne $null }
$sendOnBehalfPermissions = (Get-Mailbox $mailboxToScan.DisplayName | fl ).GrantSendOnBehalfTo

# Check each permission type
foreach ($permission in $fullAccessPermissions) {
    $user = $permission.User.ToString()
    try {
        $resolvedUser = Get-ADUser $user -ErrorAction Stop
    } catch {
        $invalidPermissionsReport += "Full Access permission to '$user' on mailbox '$mailboxToScan' is invalid or the user is disabled."
    }
}

foreach ($user in $sendAsPermissions) {
    try {
        $resolvedUser = Get-ADUser $user -ErrorAction Stop
    } catch {
        $invalidPermissionsReport += "Send As permission for '$user' is invalid or the user is disabled."
    }
}

foreach ($user in $sendOnBehalfPermissions) {
    try {
        $resolvedUser = Get-ADUser $user -ErrorAction Stop
    } catch {
        $invalidPermissionsReport += "Send On Behalf permission for '$user' is invalid or the user is disabled."
    }
}
}

# Output the report
$invalidPermissionsReport | Export-csv C:\Scripts\InvalidPermissionReports.csv -NoTypeInformation


