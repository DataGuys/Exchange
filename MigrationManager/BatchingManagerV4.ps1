# Define the array of recipient email addresses`
$emailRecipients = @("ghall@helient.com")
# Calculate the date for the sync interval, change the -1 to how many days between syncs. current setting is 1 day / 24 hours
$Syncinterval = (Get-Date).AddDays(-1)
#Modules necessary for use
#Install-Module -Name SharePointPnPPowerShellOnline -Scope AllUsers
Import-Module ExchangeOnlineManagement
Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue
Import-Module CredentialManager
# Connection settings for the API in Entra ID - This is setup during onboarding
$certThumbprint = "5E79521C7F31B61A5870F20BE06CF78E3947655C"
$appid = '0609911f-6d2a-47a7-a49c-d41535b2d4a3'
$Tenantdomain = 'polsinelli.onmicrosoft.com'
$tenantid = 'c824b048-0922-4f96-87fc-bf920e1a756d'
# Store creds in the local windows credential manager if needed. In this case its used during new-moverequest and sending status emails.
$creds = Get-StoredCredential -Target "Batcher"

#Scan SharePoint List for T-14 Updates
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
                        @{Name='PercentComplete';Expression={$_.FieldValues.PercentComplete}},
                        @{Name='CompleteMigration';Expression={$_.FieldValues.CompleteMigration}},
                        @{Name='BadItems';Expression={$_.FieldValues.BadItems}},
                        @{Name='LargeItems';Expression={$_.FieldValues.LargeItems}},
                        @{Name='RepairMailbox';Expression={$_.FieldValues.RepairMailbox}},
                        @{Name='RepairMailboxResults';Expression={$_.FieldValues.RepairMailboxResults}},
                        @{Name='LastSync';Expression={$_.FieldValues.LastSync}},
                        @{Name='BatchName';Expression={$_.FieldValues.BatchName}}

$BatchedUsers = $TableView | Where-Object {($_.Tminus14 -ne $null) -and ($_.PercentComplete -ne "100")}

#Connect to Exchange Online and run the batching logic then collect results
Connect-ExchangeOnline -AppId $appid -Organization $Tenantdomain -CertificateThumbprint $certThumbprint
$endpoint = Get-MigrationEndpoint -Identity "Hybrid Migration Endpoint - EWS (Default Web Site)"

$CompleteUsers = $BatchedUsers | Where-Object {$_.CompleteMigration -eq $true}
 foreach ($user in $CompleteUsers) {
    try {
            Set-MoveRequest -Identity $user.PrimarySmtpAddress -SkippedItemApprovalTime $([DateTime]::UtcNow) -SuspendWhenReadyToComplete:$false -PreventCompletion:$false -CompleteAfter:$null -BadItemLimit 10000 -LargeItemLimit 10000 -AcceptLargeDataLoss -WarningAction SilentlyContinue
            Get-MoveRequest $user.PrimarySmtpAddress | Resume-MoveRequest
            $result = "Move Request Completing For $($user.PrimarySmtpAddress)"
    } catch {
        $result = "Failed to set complete on $($user.PrimarySmtpAddress)."
    }
    $results += [PSCustomObject]@{ User = $user.PrimarySmtpAddress; Action = 'Complete'; BatchName = $newBatchName; Result = $result }
    continue # Skip to the next iteration
}
# Initialize results collection
$results = @()
# Retrieve all move requests
$AllMoveRequests = Get-MoveRequest -ErrorAction SilentlyContinue

# Clear Move Request. Runs first in order of operations to kill any moves before processing new ones. 
#Iterate through each move request Look for ones to clear
foreach ($moveRequest in $AllMoveRequests) {
    $moveRequestUser = Get-MailUser $moveRequest.Identity
    $userInBatch = $BatchedUsers | Where-Object { $_.PrimarySmtpAddress -eq $moveRequestUser.PrimarySmtpAddress}

    if (-not $userInBatch) {
        # User not found in BatchedUsers, remove their move request
        try {
            Remove-MoveRequest -Identity $moveRequest.Identity -Confirm:$false
            $result = "Move request for $($moveRequestUser.DisplayName) has been removed due to cleared Tminus14 date."
        } catch {
            $result = "Failed to remove MoveRequest for $($moveRequestUser.DisplayName): $($_.Exception.Message)"
        }
        $results += [PSCustomObject]@{ User = $moveRequestUser.DisplayName; Action = 'Remove'; Result = $result }
    }

    if (-not $userInBatch) {
# Clear SharePoint list item operations
    $item = Get-PnPListItem -List "OnPremMailboxes" | Where-Object {$_.FieldValues.Title -eq $moveRequestUser.ExchangeGuid}
    if ($item) {
        # Clear the SharePoint list item fields
        $fieldsToUpdate = @{"DataConsistencyScore" = $null; "PercentComplete" = $null; "BatchName" = $null; "LargeItems" = $null; "BadItems" = $null; "LastSync" = $null}
        foreach ($field in $fieldsToUpdate.Keys) {
            Set-PnPListItem -List "OnPremMailboxes" -Identity $item.Id -Values @{$field = $fieldsToUpdate[$field]}
        }
    }
    }
}

# Batching Logic #
foreach ($user in $BatchedUsers) {
    $formattedTminus0 = $user."Tminus0".Replace("/", "-")
    $newBatchName = "MB_$formattedTminus0"
    $userPrimarySmtpAddress = $user.PrimarySmtpAddress

    # Attempt to get an existing move request
    $MoveRequest = Get-MoveRequest -Identity $userPrimarySmtpAddress -erroraction silentlycontinue
    
    # Check if the MoveRequest already exists with the same BatchName
    if ($null -ne $MoveRequest -and $MoveRequest.BatchName -eq $newBatchName) {
        $result = "No change for $userPrimarySmtpAddress, already in batch $newBatchName"
        $results += [PSCustomObject]@{ User = $userPrimarySmtpAddress; Action = 'None'; BatchName = $newBatchName; Result = $result }
        Continue # Skip to the next iteration
    }
    if ($null -ne $MoveRequest -and $MoveRequest.Status -eq "Completed") {
        $result = "No change for $userPrimarySmtpAddress, already completed migration"
        $results += [PSCustomObject]@{ User = $userPrimarySmtpAddress; Action = 'Completed'; BatchName = "Completed"; Result = $result }
        continue # Skip to the next iteration
    }
    if ($null -eq $MoveRequest ) {
        try {
            New-MoveRequest -Identity $userPrimarySmtpAddress -BatchName $newBatchName -Remote -RemoteHostName $endpoint.RemoteServer -TargetDeliveryDomain "polsinelli.mail.onmicrosoft.com" -CompleteAfter "09/01/2025 5:00 PM" -PreventCompletion:$true -RemoteCredential $creds -ErrorAction Stop
            $result = "New MoveRequest created for $userPrimarySmtpAddress with batch name $newBatchName"
        } catch {
            $result = "Failed to create New MoveRequest for $userPrimarySmtpAddress $($_.Exception.Message)"
        }
        $results += [PSCustomObject]@{ User = $userPrimarySmtpAddress; Action = 'Create'; BatchName = $newBatchName; Result = $result }
        continue # Skip to the next iteration
    }
    if ($MoveRequest.BatchName -ne $newBatchName) {
        try {
            Set-MoveRequest -Identity $userPrimarySmtpAddress -BatchName $newBatchName -CompleteAfter "09/01/2025 5:00 PM" -PreventCompletion:$true -erroraction Stop
            $result = "MoveRequest for $userPrimarySmtpAddress updated to batch name $newBatchName"
        } catch {
            $result = "Failed to update MoveRequest for $userPrimarySmtpAddress $($_.Exception.Message)"
        }
        $results += [PSCustomObject]@{ User = $userPrimarySmtpAddress; Action = 'Update'; BatchName = $newBatchName; Result = $result }
        Continue
    }
   
}

#convert Batching results to HTML
$resultsHtml = $results | ConvertTo-Html -As Table | Out-String

$UpdatedMoveRequests = Get-MoveRequest

# Single call to Get-MoveRequestStatistics
$MoveStats = $UpdatedMoveRequests | Get-MoveRequestStatistics | Select-Object DisplayName, BatchName, Status, PercentComplete, DataConsistencyScore, BadItemsEncountered, LargeItemsEncountered, ExchangeGuid, LastSuccessfulSyncTimestamp, StatusDetail

# Convert Move Statistics to HTML for email
$MoveStatsHTML = $MoveStats | Select-Object @{Name='DisplayName';Expression={$_.DisplayName}}, 
    @{Name='Batch';Expression={$_.BatchName}}, 
    @{Name='Status';Expression={$_.Status}}, 
    @{Name='%Complete';Expression={$_.PercentComplete}}, 
    @{Name='DCS';Expression={$_.DataConsistencyScore}}, 
    @{Name='BadItems';Expression={$_.BadItemsEncountered}}, 
    @{Name='LargeItems';Expression={$_.LargeItemsEncountered}}, 
    @{Name='LastSync';Expression={$_.LastSuccessfulSyncTimestamp}}, 
    @{Name='MailboxGUID';Expression={$_.ExchangeGuid}} | Sort-Object DisplayName |
    ConvertTo-Html -As Table -Property 'DisplayName', 'Batch', 'Status', '%Complete', 'DCS', 'LastSync', 'BadItems', 'LargeItems', 'MailboxGUID' | Out-String

# Send emails
foreach ($recipient in $emailRecipients) {
    Send-MailMessage -From "migrationteam@polsinelli.com" -To $recipient -Subject "Migration Batch Update Results" -Body $resultsHtml -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) -SmtpServer "dc1-p-hex-01"
    Send-MailMessage -From "migrationteam@polsinelli.com" -To $recipient -Subject "Move Stats" -Body $MoveStatsHTML -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) -SmtpServer "dc1-p-hex-01"
}

# Update SharePoint List
$MoveStats | ForEach-Object {
    $currentStat = $_
    $item = Get-PnPListItem -List "OnPremMailboxes" | Where-Object {$_.FieldValues.Title -eq $currentStat.ExchangeGuid}
    if ($item) {
        Set-PnPListItem -List "OnPremMailboxes" -Identity $item.Id -Values @{
            "DataConsistencyScore" = $currentStat.DataConsistencyScore;
            "PercentComplete" = $currentStat.PercentComplete;
            "BatchName" = $currentStat.BatchName;
            "LargeItems" = $currentStat.LargeItemsEncountered;
            "BadItems" = $currentStat.BadItemsEncountered;
            "LastSync" = $currentStat.LastSuccessfulSyncTimestamp
        }
    }
}

# Emulate incremental syncing as Microsoft does in the Migration Batch settings
# Get move requests, retrieve their statistics, and filter based on the LastSuccessfulSyncTimestamp
$LastSyncMoreThan24hoursAgo = $MoveStats | Where-Object {($_.LastSuccessfulSyncTimestamp -le $Syncinterval) -and ($_.StatusDetail -like "AutoSuspended")}
# Resume each move request that hasn't synchronized in the last 24 hours and is not completed
ForEach ($IncMoveRequest in $LastSyncMoreThan24HoursAgo) {
    Try {
    Get-MoveRequest -Identity $IncMoveRequest.DisplayName | Resume-MoveRequest -ErrorAction Stop
            Write-Host "Successfully resumed move request for $($IncMoveRequest.DisplayName)"
    } Catch {
        Write-Error "Failed to resume move request for $($IncMoveRequest.DisplayName). Error: $_"
    }
}

Disconnect-ExchangeOnline -Confirm:$false

# Ensure the Exchange Management Snap-in is loaded for on-premises commands
if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}

#Find all batched users that are missing the polsinelli.mail.onmicrosoft.com email alias.
$Results | Where-Object {$_.Result -like "*TargetDeliveryDomainMismatchPermanentException*"} | ForEach-Object {
    $userIdentity = $_.user  # User identity with the error
    
    # Retrieve the current mailbox
    $mailbox = Get-Mailbox -Identity $userIdentity 
    $mailbox.EmailAddresses
    # Construct the new email alias correctly
    # Assuming you want to add this as an additional alias, not replace the existing one
    $newAlias = "smtp:" + $mailbox.Alias + "@polsinelli.mail.onmicrosoft.com"

    # Attempt to add the new email alias if it does not already exist
    if ($mailbox.EmailAddresses -notcontains $newAlias) {
        try {
            # Use EmailAddresses+=$newAlias to add the new alias
            Set-Mailbox -Identity $userIdentity -EmailAddresses @{Add=$newAlias}
            Write-Host "Successfully added $newAlias to $userIdentity."
        } catch {
            Write-Host "Error adding $newAlias to $userIdentity $_"
        }
    } else {
        Write-Host "$userIdentity already has the alias $newAlias."
    }
}


#Set SPO list to 100 percent complete for this error on the move request.
#Find all batched users that are already migrated.
$Results | Where-Object {$_.Result -like "*TargetUserAlreadyHasPrimaryMailboxException*"} | ForEach-Object {
    $userIdentity = $_.user  # User identity - already completed migration
    # Set 100 percent loop
    $mailbox = Get-Recipient -Identity $userIdentity
        $item = Get-PnPListItem -List "OnPremMailboxes" | Where-Object {$_.FieldValues.Title -eq $mailbox.ExchangeGuid}
    if ($item) {
        # Set 100 percent complete and clear the noise from the SPO list for the completed user
        Set-PnPListItem -List "OnPremMailboxes" -Identity $item.Id -Values @{"DataConsistencyScore" = $null; "BatchName" = "Completed"; "PercentComplete" = "100"; "LargeItems" = $null; "BadItems" = $null;}
        }
      
    }
    # Add Users to Intune_EXO Security Group
    $CompletedUsers = $TableView | Where-Object {$_.PercentComplete -eq "100"}
    foreach ($user in $CompletedUsers) {
    $userIdentity = $user.PrimarySmtpAddress # User identity - already completed migration
    # Set 100 percent loop
    $mailbox = Get-Recipient -Identity $userIdentity -ReadFromDomainController
    $ADUserDN = $mailbox.DistinguishedName
        # Check if user is already a member of the group
        $groupMembers = Get-ADGroupMember "Intune-EXO" | Select-Object -ExpandProperty DistinguishedName
        if ($ADUserDN -notin $groupMembers) {
            # User is not a member, add them to the group
          Add-ADGroupMember -Identity "Intune-EXO" -Members $ADUserDN 
        }
        }
Get-PSSnapin | Remove-PSSnapin -Confirm:$false -ErrorAction SilentlyContinue
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-PnPOnline
Exit