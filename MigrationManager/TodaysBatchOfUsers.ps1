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
$TableView = $items | Select-Object @{Name='Display Name';Expression={$_.FieldValues.field_0}},
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
                        @{Name='Title';Expression={$_.FieldValues.Title}}

$TodaysBatchedUsers = $TableView | Where-Object { $_.Tminus0 -eq (Get-Date -Format "MM/dd/yyyy") }
$TodaysBatchedUsersHTML = $TodaysBatchedUsers | ConvertTo-Html -As Table | Out-String
Send-MailMessage -From "gregory.halladmin@polsinelli.com" -To "ghall@helient.com" -Subject "Today's Batch of Users" -Body $TodaysBatchedUsersHTML -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) -SmtpServer "dc1-p-hex-01" -Credential $creds