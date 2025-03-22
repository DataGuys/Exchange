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
Connect-PnPOnline -Url "https://polsinelli.sharepoint.com/sites/ExchangeOnlineProject" -ClientId $appid -Thumbprint $certThumbprint -Tenant $tenantId -WarningAction SilentlyContinue
$items = Get-PnPListItem -List "OnPremMailboxes" 
$TableView = $items | Select-Object @{Name='DisplayName';Expression={$_.FieldValues.field_0}},
                        @{Name='Alias';Expression={$_.FieldValues.field_1}},
                        @{Name='UserPrincipalName';Expression={$_.FieldValues.field_2}},
                        @{Name='PrimarySmtpAddress';Expression={$_.FieldValues.field_3}},
                        @{Name='OUPath';Expression={$_.FieldValues.field_5}},
                        @{Name='TotalItemSizeMB';Expression={$_.FieldValues.field_6}},
                        @{Name='TotalItemCount';Expression={$_.FieldValues.field_7}},
                        @{Name='ReceipientType';Expression={$_.FieldValues.field_8}},
                        @{Name='Tminus14';Expression={$_.FieldValues.T_x002d_14}},
                        @{Name='Tminus7';Expression={$_.FieldValues.T_x002d_7}},
                        @{Name='Tminus3';Expression={$_.FieldValues.T_x002d_3}},
                        @{Name='Tminus0';Expression={if ($_.FieldValues.T_x002d_0) {Get-Date $_.FieldValues.T_x002d_0 -Format "MM/dd/yyyy"}}}, # Format Tminus0 here
                        @{Name='Title';Expression={$_.FieldValues.Title}},
                        @{Name='AcceptDataLoss';Expression={$_.FieldValues.AcceptDataLoss}},
                        @{Name='CompleteMigration';Expression={$_.FieldValues.CompleteMigration}},
                        @{Name='BadItems';Expression={$_.FieldValues.BadItems}},
                        @{Name='LargeItems';Expression={$_.FieldValues.LargeItems}},
                        @{Name='BatchName';Expression={$_.FieldValues.BatchName}},
                        @{Name='InteractionUser';Expression={$_.FieldValues.InteractionUser}}

$TableView | Format-Table -AutoSize

$InteractionUserList = Import-Csv C:\ExchangeMigration\Interaction.csv
$InteractionUserList.Count
# Ensure the Exchange Management Snap-in is loaded
if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}
$InteractionUserList | ForEach-Object {
    $User = $_.DisplayName
    $ExchangeGuid = Get-Mailbox $User -ReadFromDomainController | Select-Object ExchangeGuid
    # Retrieve the list item matching the ExchangeGuid
    $item = Get-PnPListItem -List "OnPremMailboxes" | Where-Object {$_.FieldValues.Title -eq $ExchangeGuid.ExchangeGuid}
    if ($item) {
        # Tag user in SPO list as Interaction user
        Set-PnPListItem -List "OnPremMailboxes" -Identity $item.Id -Values @{"InteractionUser" = $true}
    }
    }