# User defined Variables 
# Define the array of recipient email addresses
$emailRecipients = @("ghall@helient.com")
#Modules necessary for use
#Install-Module -Name SharePointPnPPowerShellOnline -Scope AllUsers
Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue

# Connection settings for the API in Entra ID - This is setup during onboarding
$certThumbprint = "5E79521C7F31B61A5870F20BE06CF78E3947655C"
$appid = '0609911f-6d2a-47a7-a49c-d41535b2d4a3'
$Tenantdomain = 'polsinelli.onmicrosoft.com'
$tenantid = 'c824b048-0922-4f96-87fc-bf920e1a756d'
# Connect to the SharePoint Online site for the Exchange Online Project
Connect-PnPOnline -Url "https://polsinelli.sharepoint.com/sites/ExchangeOnlineProject" `
                  -ClientId $appid -Thumbprint $certThumbprint `
                  -Tenant $tenantId -WarningAction SilentlyContinue

# Retrieve items from the "OnPremMailboxes" list
$items = Get-PnPListItem -List "OnPremMailboxes"

# Format the retrieved items for easier viewing
$TableView = $items | Select-Object @{Name='DisplayName'; Expression={$_.FieldValues.field_0}},
                                        @{Name='Alias'; Expression={$_.FieldValues.field_1}},
                                        @{Name='UserPrincipalName'; Expression={$_.FieldValues.field_2}},
                                        @{Name='PrimarySmtpAddress'; Expression={$_.FieldValues.field_3}},
                                        @{Name='RepairMailbox'; Expression={$_.FieldValues.RepairMailbox}},
                                        @{Name='RepairMailboxResults'; Expression={$_.FieldValues.RepairMailboxResults}},
                                        @{Name='LastSync'; Expression={$_.FieldValues.LastSync}},
                                        @{Name='BatchName'; Expression={$_.FieldValues.BatchName}}

# Filter for mailboxes marked for repair
$Repairmailboxes = $TableView | Where-Object {$_.RepairMailbox -eq $true}

# Ensure the Exchange Management Snap-in is loaded for on-premises commands
if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}


foreach ($Mailbox in $Repairmailboxes) {
    $MailboxStats = Get-MailboxStatistics $Mailbox.PrimarySmtpAddress

    # Retrieve the SharePoint list item for the mailbox
    $item = Get-PnPListItem -List "OnPremMailboxes" | Where-Object {$_.FieldValues.Title -eq $Mailbox.DisplayName}

    # Check for existing repair requests in the last 30 days
    $existingRepairRequests = Get-MailboxRepairRequest -StoreMailbox $MailboxStats.MailboxGuid -Database $MailboxStats.Database |
                              Where-Object { $_.CreationTime -gt (Get-Date).AddDays(-30) }

    if ($existingRepairRequests) {
        # Update SharePoint list indicating no new repair needed due to recent request
        Set-PnPListItem -List "OnPremMailboxes" -Identity $item.Id -Values @{
            "RepairMailboxResults" = "No new repair initiated; a request was already made within the last 30 days."
            "RepairMailbox" = $false
        }
    } else {
        # Issue a new repair request
        New-MailboxRepairRequest -Mailbox $MailboxStats.MailboxGuid -CorruptionType FolderACL,ProvisionedFolder,SearchFolder,AggregateCounts,Folderview

        # Update SharePoint list with the initiation of a new repair request and clear the repair flag
        Set-PnPListItem -List "OnPremMailboxes" -Identity $item.Id -Values @{
            "RepairMailboxResults" = "New repair request initiated."
            "RepairMailbox" = $false
        }
    }
}



Disconnect-PnPOnline
Exit