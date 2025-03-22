# Import required modules
Import-Module ExchangeOnlineManagement
Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue
Import-Module CredentialManager

# Initialize variables
$testMode = $true # Failsafe! Only test users get emails, change to $false when ready
$testEmailAddresses = @("ghall@helient.com") # Add more as needed
$today = Get-Date -Format "MM/dd/yyyy"
$creds = Get-StoredCredential -Target "Batcher"
$outputPath = "C:\ExchangeMigration"
$certThumbprint = "5E79521C7F31B61A5870F20BE06CF78E3947655C"
$appId = '0609911f-6d2a-47a7-a49c-d41535b2d4a3'
$tenantDomain = 'polsinelli.onmicrosoft.com'
$tenantId = 'c824b048-0922-4f96-87fc-bf920e1a756d'
$smtpServer = "dc1-p-hex-01"
$smtpFrom = "migrationteam@polsinelli.com"
$smtpCredential = $creds


# Function to send emails with an option for test mode
function Send-TminusEmail {
    param (
        [string]$EmailAddress,
        [string]$Subject,
        [string]$Body
            )
    $recipients = { $EmailAddress }
    foreach ($recipient in $recipients) {
        Send-MailMessage -SmtpServer $smtpServer -From $smtpFrom -To $recipient -Subject $Subject -Body $Body -BodyAsHtml -Credential $smtpCredential
    }
}

# Connect to SharePoint and retrieve items
Connect-PnPOnline -Url "https://polsinelli.sharepoint.com/sites/ExchangeOnlineProject" -ClientId $appId -Thumbprint $certThumbprint -Tenant $tenantId -WarningAction SilentlyContinue
$items = Get-PnPListItem -List "OnPremMailboxes"

# Construct table view from SharePoint list items
$tableView = $items | Select-Object @{Name='Display Name';Expression={$_.FieldValues.field_0}},
    @{Name='Alias';Expression={$_.FieldValues.field_1}},
    @{Name='UserPrincipalName';Expression={$_.FieldValues.field_2}},
    @{Name='PrimarySmtpAddress';Expression={$_.FieldValues.field_3}},
    @{Name='OUPath';Expression={$_.FieldValues.field_5}},
    @{Name='TotalItemSizeMB';Expression={$_.FieldValues.field_6}},
    @{Name='TotalItemCount';Expression={$_.FieldValues.field_7}},
    @{Name='ReceipientType';Expression={$_.FieldValues.field_8}},
    @{Name='Tminus14';Expression={if ($_.FieldValues.T_x002d_14) {Get-Date $_.FieldValues.T_x002d_14 -Format "MM/dd/yyyy"}}},
    @{Name='Tminus7';Expression={if ($_.FieldValues.T_x002d_7) {Get-Date $_.FieldValues.T_x002d_7 -Format "MM/dd/yyyy"}}},
    @{Name='Tminus3';Expression={if ($_.FieldValues.T_x002d_3) {Get-Date $_.FieldValues.T_x002d_3 -Format "MM/dd/yyyy"}}},
    @{Name='Tminus0';Expression={if ($_.FieldValues.T_x002d_0) {Get-Date $_.FieldValues.T_x002d_0 -Format "MM/dd/yyyy"}}},
    @{Name='Title';Expression={$_.FieldValues.Title}}

# Filter users based on Tminus14 not being null
$batchedUsers = $tableView | Where-Object {$_.Tminus14 -ne $null}
$batchedUsers | Select-Object UserPrincipalName, Tminus14, Tminus0 | Format-Table
Disconnect-PnPOnline


if ($testMode -eq $true){
foreach ($testuser in $testEmailAddresses) {
# Determine days to migration
    $daysToMigration = @("14","7","3","0")
  foreach ($day in $daysToMigration) {
        # Construct the file path using the correct daysToMigration value
        $emailTemplatePath = "C:\Scripts\EmailTemplate_" + $day + ".html"
        $emailBody = Get-Content -Path $emailTemplatePath -Raw
        # Send the email
           Send-MailMessage -SmtpServer $smtpServer -From $smtpFrom -To $testuser -Subject "Your migration is in $day days" -Body $emailBody -BodyAsHtml -Credential $smtpCredential
        } else {
           Write-Warning "The email template for $daysToMigration days to migration does not exist at path: $emailTemplatePath"
        }
    }
    Exit
    }

foreach ($user in $batchedUsers) {
    # Determine days to migration
    $daysToMigration = $null
    if ($today -eq $user.Tminus14) { $daysToMigration = "14" }
    elseif ($today -eq $user.Tminus7) { $daysToMigration = "7" }
    elseif ($today -eq $user.Tminus3) { $daysToMigration = "3" }
    elseif ($today -eq $user.Tminus0) { $daysToMigration = "0" }

    if ($daysToMigration -ne $null) {
        # Construct the file path using the correct daysToMigration value
        $emailTemplatePath = "C:\Scripts\EmailTemplate_" + $daysToMigration + ".html"
        
        # Check if the template file exists to avoid errors
        if (Test-Path $emailTemplatePath) {
            $emailBody = Get-Content -Path $emailTemplatePath -Raw

            # Send the email
            # Ensure $smtpServer, $smtpFrom, and other email details are correctly set
            Send-MailMessage -SmtpServer $smtpServer -From $smtpFrom -To $user.PrimarySmtpAddress -Subject "Your migration is in $daysToMigration days" -Body $emailBody -BodyAsHtml -Credential $smtpCredential
        } else {
            Write-Warning "The email template for $daysToMigration days to migration does not exist at path: $emailTemplatePath"
        }
    }
}


