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
                        @{Name='BatchName';Expression={$_.FieldValues.BatchName}}

$BatchedUsers = $TableView | Where-Object {$_.Tminus14 -ne $null}
Disconnect-PnPOnline -ErrorAction SilentlyContinue

Connect-ExchangeOnline -AppId $appid -Organization $Tenantdomain -CertificateThumbprint $certThumbprint
$endpoint = Get-MigrationEndpoint -Identity "Hybrid Migration Endpoint - EWS (Default Web Site)"
$results = @() # Ensure $results is initialized to store the outcome

# Cleanup logic for removing move requests whose Tminus14 dates have been cleared
$AllMoveRequests = Get-MoveRequest -ErrorAction SilentlyContinue
#Clear Move request Logic
foreach ($moveRequest in $AllMoveRequests) {
    $moveRequestUser = Get-MailUser $moveRequest.Identity
    $userInBatch = $BatchedUsers | Where-Object { $_.PrimarySmtpAddress -eq $moveRequestUser.PrimarySmtpAddress}

    if (-not $userInBatch) {
        # User not found in BatchedUsers, remove their move request
        try {
            Remove-MoveRequest -Identity $moveRequestUser -Confirm:$false
            $result = "Move request for $moveRequestUser has been removed due to cleared Tminus14 date."
        } catch {
            $result = "Failed to remove MoveRequest for $moveRequestUser $($_.Exception.Message)"
        }
        $results += [PSCustomObject]@{ User = $moveRequestUser; Action = 'Remove'; Result = $result }
    }
}

# Batching Logic
foreach ($user in $BatchedUsers) {
    $formattedTminus0 = $user."Tminus0".Replace("/", "-")
    $newBatchName = "MB_$formattedTminus0"
    $userPrimarySmtpAddress = $user.PrimarySmtpAddress

    # Attempt to get an existing move request
    $MoveRequest = Get-MoveRequest -Identity $userPrimarySmtpAddress -erroraction silentlycontinue

    # Check if the MoveRequest already exists with the same BatchName
    if ($MoveRequest -ne $null -and $MoveRequest.BatchName -eq $newBatchName) {
        $result = "No change for $userPrimarySmtpAddress, already in batch $newBatchName"
        $results += [PSCustomObject]@{ User = $userPrimarySmtpAddress; Action = 'None'; BatchName = $newBatchName; Result = $result }
        continue # Skip to the next iteration
    }
    New-MoveRequest -in
    if ($MoveRequest -eq $null) {
        try {
            $output = New-MoveRequest -Identity $userPrimarySmtpAddress -BatchName $newBatchName -Remote -RemoteHostName $endpoint.RemoteServer -TargetDeliveryDomain "polsinelli.mail.onmicrosoft.com" -CompleteAfter "09/01/2025 5:00 PM" -IncrementalSyncInterval 24:00:00 -PreventCompletion:$true -RemoteCredential $creds -ErrorAction Stop
            $result = "New MoveRequest created for $userPrimarySmtpAddress with batch name $newBatchName"
        } catch {
            $result = "Failed to create New MoveRequest for $userPrimarySmtpAddress $($_.Exception.Message)"
        }
        $results += [PSCustomObject]@{ User = $userPrimarySmtpAddress; Action = 'Create'; BatchName = $newBatchName; Result = $result }
    } elseif ($MoveRequest.BatchName -ne $newBatchName) {
        try {
            $output = Set-MoveRequest -Identity $userPrimarySmtpAddress -BatchName $newBatchName -CompleteAfter "09/01/2025 5:00 PM" -IncrementalSyncInterval 24:00:00 -PreventCompletion:$true -erroraction Stop
            $result = "MoveRequest for $userPrimarySmtpAddress updated to batch name $newBatchName"
        } catch {
            $result = "Failed to update MoveRequest for $userPrimarySmtpAddress $($_.Exception.Message)"
        }
        $results += [PSCustomObject]@{ User = $userPrimarySmtpAddress; Action = 'Update'; BatchName = $newBatchName; Result = $result }
    } elseif ($BatchedUsers.AcceptDataLoss -eq $true) {
        try {
            $output = Set-MoveRequest -Identity $userPrimarySmtpAddress -AcceptLargeDataLoss -ErrorAction Stop
            $result = "Updated Move Request for $userPrimarySmtpAddress To Accept Large Data Loss. Be sure this is what you wanted before completion"
        } catch {
            $result = "Failed to update move request for $userPrimarySmtpAddress $($_.Exception.Message)"
        }
        $results += [PSCustomObject]@{ User = $userPrimarySmtpAddress; Action = 'Update'; BatchName = $newBatchName; Result = $result }
    } elseif ($BatchedUsers.ForceComplete -eq $true) {
        try {
            $output = Set-MoveRequest -Identity $userPrimarySmtpAddress -SkippedItemApprovalTime $([DateTime]::UtcNow) -SuspendWhenReadyToComplete:$false -PreventCompletion:$false -CompleteAfter $null -ApproveSkippedItems -SyncNow
            Get-MoveRequest $userPrimarySmtpAddress | Resume-MoveRequest
            $result = "Updated Move Requiest for $userPrimarySmtpAddress to force it to complete, if this does not work then you need to dig into the problem."
        } catch {
            $result = "Failed to set force complete on $userPrimarySmtpAddress. Something is wrong with this user check it"
        }
        $results += [PSCustomObject]@{ User = $userPrimarySmtpAddress; Action = 'Force Complete'; BatchName = $newBatchName; Result = $result }
    } elseif ($BatchedUsers.CompleteMigration -eq $true) {
        try {
            $output = Set-MoveRequest -Identity $userPrimarySmtpAddress -SuspendWhenReadyToComplete:$false -PreventCompletion:$false -CompleteAfter $null -SyncNow -CompletedRequestAgeLimit 7
            Get-MoveRequest $userPrimarySmtpAddress | Resume-MoveRequest
            $result = "Updated Move Requiest for $userPrimarySmtpAddress to set graceful complete, if the DataConsistancyScore is Good or Perfect this will run without worry"
        } catch {
            $result = "Failed to set graceful complete on $userPrimarySmtpAddress. Something is wrong with this users migration check it"
        }
        $results += [PSCustomObject]@{ User = $userPrimarySmtpAddress; Action = 'Graceful Complete'; BatchName = $newBatchName; Result = $result }
}
}

#convert Batching results to HTML
$resultsHtml = $results | ConvertTo-Html -As Table | Out-String

#Get Move Statistics and email updates to the migration team
$MoveStats = Get-MoveRequest | Get-MoveRequestStatistics | Select-Object DisplayName, BatchName, Status, PercentComplete, DataConsistencyScore, BadItemsEncountered, LargeItemsEncountered, ExchangeGuid
$MoveStatsHTML = $MoveStats | ConvertTo-Html -As Table | Out-String

# Define the array of recipient email addresses
$emailRecipients = @("ghall@helient.com", "sstockton@polsinelli.com", "callen@polsinelli.com")

# Loop through each email address in the array and send the emails
foreach ($recipient in $emailRecipients) {
    Send-MailMessage -From "migrationteam@polsinelli.com" -To $recipient -Subject "Migration Batch Update Results" -Body $resultsHTML -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) -SmtpServer "dc1-p-hex-01" -Credential $creds
    Send-MailMessage -From "migrationteam@polsinelli.com" -To $recipient -Subject "Move Stats" -Body $MoveStatsHTML -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) -SmtpServer "dc1-p-hex-01"
}

Disconnect-ExchangeOnline -Confirm:$false

#Reconnect to Sharepoint 
Connect-PnPOnline -Url "https://polsinelli.sharepoint.com/sites/ExchangeOnlineProject" -ClientId $appid -Thumbprint $certThumbprint -Tenant $tenantId -WarningAction SilentlyContinue
#Post all move request Data Consistancy Scores, Percent Complete, Batch Name, BadItems, and LargeItems back to the SharePoint List
$MoveStats | ForEach-Object {
    $currentStat = $_
    # Retrieve the list item matching the ExchangeGuid
    $item = Get-PnPListItem -List "OnPremMailboxes" | Where-Object {$_.FieldValues.Title -eq $currentStat.ExchangeGuid}
    if ($item) {
        # Update the DataConsistencyScore for the found item
        Set-PnPListItem -List "OnPremMailboxes" -Identity $item.Id -Values @{"DataConsistencyScore" = $currentStat.DataConsistencyScore}
        Set-PnPListItem -List "OnPremMailboxes" -Identity $item.Id -Values @{"PercentComplete" = $currentStat.PercentComplete}
        Set-PnPListItem -List "OnPremMailboxes" -Identity $item.Id -Values @{"BatchName" = $currentStat.BatchName}
        Set-PnPListItem -List "OnPremMailboxes" -Identity $item.Id -Values @{"LargeItems" = $currentStat.LargeItemsEncountered}
        Set-PnPListItem -List "OnPremMailboxes" -Identity $item.Id -Values @{"BadItems" = $currentStat.BadItemsEncountered}
    }
    }
    
# Ensure the Exchange Management Snap-in is loaded
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

#Add completed users to the "Intune-EXO" security group
foreach ($moveRequest in $AllMoveRequests) {
if ($moveRequest.Status -eq "Completed") {
$ADUserDN = (Get-User $moveRequest.DisplayName).DistinguishedName
Get-Group "Intune-EXO" | Add-ADGroupMember $ADUserDN
}
}

Get-PSSnapin | Remove-PSSnapin -Confirm:$false -ErrorAction SilentlyContinue
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-PnPOnline
Exit


