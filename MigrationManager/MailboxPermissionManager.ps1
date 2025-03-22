# Define the output path
#$outputPath = "C:\ExchangeMigration"
#Start-Transcript -OutputDirectory $outputPath
#Install-Module -Name SharePointPnPPowerShellOnline -Scope AllUsers
#Import-Module ExchangeOnlineManagement
Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue
#Import-Module CredentialManager
$certThumbprint = "5E79521C7F31B61A5870F20BE06CF78E3947655C"
$appid = '0609911f-6d2a-47a7-a49c-d41535b2d4a3'
#$Tenantdomain = 'polsinelli.onmicrosoft.com'
$tenantid = 'c824b048-0922-4f96-87fc-bf920e1a756d'

Connect-PnPOnline -Url "https://polsinelli.sharepoint.com/sites/ExchangeOnlineProject" -ClientId $appid -Thumbprint $certThumbprint -Tenant $tenantId -WarningAction SilentlyContinue

# Load the Exchange Management PowerShell Snap-in if it's not already loaded
if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}

# Get all on-premises mailboxes
$OnpremMailboxes = Get-Mailbox -Resultsize Unlimited

# Make this a function for parallel parocessing
#Process each mailbox to examine mailbox permissions
foreach ($moveRequest in $OnpremMailboxes) {
    $mailboxIdentity = $moveRequest.Alias
    #$DisplayName = $moveRequest.DisplayName
    $ExchangeGuid = $moveRequest.ExchangeGuid
    # Retrieve FullAccess permissions, while excluding specific system accounts and Exchange groups
    $fullAccessPermissions = Get-MailboxPermission $mailboxIdentity | Where-Object {
        $_.AccessRights -like "*FullAccess*" `
        -and $_.User -notmatch "NT AUTHORITY\\SELF" `
        -and $_.User -notmatch "^S-1-5-21" `
        -and $_.User -notmatch "POLSINELLI\\Exchange Services" `
        -and $_.User -notmatch "POLSINELLI\\Exchange Domain Servers" `
        -and $_.User -notmatch "POLSINELLI\\Exchange Organization Administrators" `
        -and $_.User -notmatch "POLSINELLI\\Domain Admins" `
        -and $_.User -notmatch "POLSINELLI\\Enterprise Admins" `
        -and $_.User -notmatch "POLSINELLI\\Organization Management" `
        -and $_.User -notmatch "POLSINELLI\\Exchange Servers1" `
        -and $_.User -notmatch "POLSINELLI\\Exchange Trusted Subsystem" `
        -and $_.User -notmatch "POLSINELLI\\iaexch" `
        -and $_.User -notmatch "POLSINELLI\\PolAdmin" `
        -and $_.User -notmatch "POLSINELLI\\SVC_EFSAdmin" `
        -and $_.User -notmatch "POLSINELLI\\pwebbadmin" `
        -and $_.User -notmatch "POLSINELLI\\bartistadmin" `
        -and $_.User -notmatch "POLSINELLI\\calleadmin" `
        -and $_.User -notmatch "POLSINELLI\\EntVaultAdmin" `
        -and $_.User -notmatch "POLSINELLI\\$mailboxIdentity" `
        -and $_.User -notmatch "NT AUTHORITY\\SYSTEM"
    } | Select-Object User, @{Name='PermissionType'; Expression={"FullAccess"}}

    # Extract SendAs permissions
    $sendAsPermissions = Get-MailboxPermission $mailboxIdentity | Where-Object {
        $_.AccessRights -like "*SendAs*" -and -not $_.IsInherited
    } | Select-Object User, @{Name='PermissionType'; Expression={"SendAs"}}

    # Collect SendOnBehalf permissions
    $mailbox = Get-Mailbox $mailboxIdentity
    $sendOnBehalfPermissions = $mailbox.GrantSendOnBehalfTo | ForEach-Object {
        [PSCustomObject]@{
            User           = $_.Name
            PermissionType = "SendOnBehalf"
        }
    }
    # Initialize permissions as an empty array
    $permissions = @()

    # Check if each permission variable is not null and add it to the permissions array
    if ($null -ne $fullAccessPermissions) {
        $permissions += $fullAccessPermissions
    }
    if ($null -ne $sendAsPermissions) {
        $permissions += $sendAsPermissions
    }
    if ($null -ne $sendOnBehalfPermissions) {
        $permissions += $sendOnBehalfPermissions
    }

    # Check if permissions are null or empty
    if (-not $permissions -or $permissions.Count -eq 0) {
        $csvContent = "No Dependents"
    } else {
        $csvContent = $permissions | ConvertTo-Csv -NoTypeInformation
    }

    $item = Get-PnPListItem -List "OnPremMailboxes" | Where-Object {$_.FieldValues.Title -eq $ExchangeGuid}
    if ($item) {
        Set-PnPListItem -List "OnPremMailboxes" -Identity $item.Id -Values @{"Dependents" = "$csvContent"}
    }

}

Disconnect-PnPOnline



