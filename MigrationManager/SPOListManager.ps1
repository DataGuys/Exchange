# Ensure the latest PnP PowerShell module is installed
# Install-Module PnP.PowerShell -Scope CurrentUser

# Import the PnP PowerShell module
Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue

# Connection settings for the API
$certThumbprint = "5E79521C7F31B61A5870F20BE06CF78E3947655C"
$appid = '0609911f-6d2a-47a7-a49c-d41535b2d4a3'
$tenantid = 'c824b048-0922-4f96-87fc-bf920e1a756d'

if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}

# Connect to SharePoint Online
$siteUrl = "https://polsinelli.sharepoint.com/sites/ExchangeOnlineProject"
Connect-PnPOnline -Url $siteUrl -ClientId $appid -Thumbprint $certThumbprint -Tenant $tenantid -WarningAction SilentlyContinue

# Retrieve items from the SharePoint list
$spItems = Get-PnPListItem -List "OnPremMailboxes"

# Extract Title (Exchange Guid) for comparison
$spGuids = $spItems | ForEach-Object { $_.FieldValues.Title }

# Retrieve on-premises mailboxes excluding specific system mailboxes
$OnPremMailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object { $_.Name -notlike "DiscoverySearchMailbox*" }

# Determine new mailboxes not already in the SharePoint list
$NewOnPremMailboxes = $OnPremMailboxes | Where-Object { $_.ExchangeGuid -notin $spGuids }

foreach ($mailbox in $NewOnPremMailboxes) {
    $mailboxStats = Get-MailboxStatistics $mailbox.Identity
    $sizeString = $mailboxStats.TotalItemSize.ToString()

    if ($sizeString -match '^(?<Size>\d+(\.\d+)?)\s*(?<Unit>KB|MB|GB)') {
        $size = [float]$Matches['Size']
        $unit = $Matches['Unit']

        # Convert size to GB based on the unit
        switch ($unit) {
            'KB' { $gbValue = $size / 1MB / 1GB }
            'MB' { $gbValue = $size / 1GB }
            'GB' { $gbValue = $size }
        }

        # Round the GB value to 2 decimal places
        $field_6 = [Math]::Round($gbValue, 2)
    } else {
        Write-Error "Failed to parse size from '$sizeString'"
    }

    # Add new mailbox information to SharePoint list
    Add-PnPListItem -List "OnPremMailboxes" -Values @{
        "Title" = $mailbox.ExchangeGuid;
        "field_0" = $mailbox.DisplayName;
        "field_1" = $mailbox.Alias;
        "field_2" = $mailbox.UserPrincipalName;
        "field_3" = $mailbox.PrimarySmtpAddress;
        "field_6" = $field_6;
        "field_7" = $mailboxStats.ItemCount;
        "field_8" = $($mailbox.RecipientTypeDetails).ToString()
    }
}

Disconnect-PnPOnline
