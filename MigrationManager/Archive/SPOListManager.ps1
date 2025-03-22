# Modules necessary for use
Import-Module ExchangeOnlineManagement
Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue

# Connection settings for the API in Entra ID - This is setup during onboarding
$certThumbprint = "5E79521C7F31B61A5870F20BE06CF78E3947655C"
$appid = '0609911f-6d2a-47a7-a49c-d41535b2d4a3'
$tenantid = 'c824b048-0922-4f96-87fc-bf920e1a756d'

# Scan SharePoint List for T-14 Updates
Connect-PnPOnline -Url "https://polsinelli.sharepoint.com/sites/ExchangeOnlineProject" -ClientId $appid -Thumbprint $certThumbprint -Tenant $tenantId -WarningAction SilentlyContinue
$spItems = Get-PnPListItem -List "OnPremMailboxes" 

# Extract Title (Exchange Guid) from SharePoint list for comparison
$spGuids = $spItems | ForEach-Object { $_.FieldValues.Title }

Disconnect-PnPOnline -Confirm:$false

# Ensure the Exchange Management Snap-in is loaded for on-premises commands
if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}

# Retrieve on-premises mailboxes
$OnPremMailboxes = Get-Mailbox -ResultSize Unlimited

# Find new on-premises mailboxes not in SharePoint list
$NewOnPremMailboxes = $OnPremMailboxes | Where-Object { $_.ExchangeGuid -notin $spGuids }

if ($NewOnPremMailboxes.Count -gt 0) {
    # Reconnect to SharePoint Online to add new items
    Connect-PnPOnline -Url "https://polsinelli.sharepoint.com/sites/ExchangeOnlineProject" -ClientId $appid -Thumbprint $certThumbprint -Tenant $tenantId -WarningAction SilentlyContinue

    # If there are new mailboxes, add them to the SharePoint list
    foreach ($mailbox in $NewOnPremMailboxes) {
        # Gather mailbox statistics for TotalItemSize and TotalItemCount
        $mailboxStats = Get-MailboxStatistics $mailbox.Identity

        # Example data to add, adjust according to your list's schema
        $itemValues = @{
            "Title" = $mailbox.ExchangeGuid
            "DisplayName" = $mailbox.DisplayName
            "UserPrincipalName" = $mailbox.UserPrincipalName
            "PrimarySmtpAddress" = $mailbox.PrimarySmtpAddress
            "TotalItemSizeMB" = [Math]::Round(($mailboxStats.TotalItemSize.Value.ToMB()), 2)
            "TotalItemCount" = $mailboxStats.ItemCount
            "RecipientTypeDetail" = $mailbox.RecipientTypeDetails.ToString()
            # Add other necessary fields here
        }
        Add-PnPListItem -List "OnPremMailboxes" -Values $itemValues
    }

    # Disconnect after operations are completed
    Disconnect-PnPOnline -Confirm:$false
} else {
    Write-Output "No new on-premises mailboxes to add to SharePoint list."
}


