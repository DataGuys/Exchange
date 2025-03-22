# Define the array of recipient email addresses`
$emailRecipients = @("ghall@helient.com")
$every6thEmailRecipients = @("sstockton@polsinelli.com", "ghall@helient.com")
# Set up send email parameters
$smtpServer = "dc1-p-hex-01"
$fromEmail = "migrationteam@polsinelli.com"
# Calculate the date for the sync interval, change the -1 to how many days between syncs. current setting is 1 day / 24 hours
$Syncinterval = (Get-Date).AddDays(-1)
$AADConnectServer = "DC2-P-ASYNC-01"
# Limit the number of concurrent jobs for the get-mailboxstatistics part
# Configure runspace pool
$minRunspacesset = 3
$maxRunspacesset = 5
#Import Modules for run
Import-Module ExchangeOnlineManagement
Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue
Import-Module CredentialManager
# Connection settings for the API in Entra ID - This is setup during onboarding
$certThumbprint = "5E79521C7F31B61A5870F20BE06CF78E3947655C"
$appid = '0609911f-6d2a-47a7-a49c-d41535b2d4a3'
$Tenantdomain = 'polsinelli.onmicrosoft.com'
$tenantid = 'c824b048-0922-4f96-87fc-bf920e1a756d'
$creds = Get-StoredCredential -Target "Batcher"

#Scan SharePoint List for T-14 Updates, minus 100 percent mailboxes.
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

#Connect to Exchange Online and run the batching logic and collect results
Connect-ExchangeOnline -AppId $appid -Organization $Tenantdomain -CertificateThumbprint $certThumbprint
# Get the hybrid endpoint info into a variable
$endpoint = Get-MigrationEndpoint -Identity "Hybrid Migration Endpoint - EWS (Default Web Site)"

# Initialize results collection
$results = @()
# Retrieve all current move requests in EXO
$AllMoveRequests = Get-MoveRequest -ErrorAction SilentlyContinue

# Clear Move Request. Runs first in order of operations to kill any moves before processing new ones. 
# Iterate through each move request Look for ones to clear if T-14 is null and a move request exists
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
    # Clear SharePoint list item operations if T-14 is null and a move request exists
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

############## Mailbox Migration Batching Logic #########
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
        continue # Skip to the next iteration
    }

    # Check if MoveRequest is Completed and report results
    if ($null -ne $MoveRequest -and $MoveRequest.Status -eq "Completed") {
        $result = "No change for $userPrimarySmtpAddress, already completed migration"
        $results += [PSCustomObject]@{ User = $userPrimarySmtpAddress; Action = 'Completed'; BatchName = "Completed"; Result = $result }
        continue # Skip to the next iteration
    }

    # Create new-moverequest if the above if statements are not true
    if ($null -eq $MoveRequest ) {
        try {
            New-MoveRequest -Identity $userPrimarySmtpAddress -BatchName $newBatchName -Remote -RemoteHostName $endpoint.RemoteServer -TargetDeliveryDomain "polsinelli.mail.onmicrosoft.com" -CompleteAfter "09/01/2025 5:00 PM" -PreventCompletion:$true -RemoteCredential $creds -ErrorAction Stop
            $result = "New MoveRequest created for $userPrimarySmtpAddress with batch name $newBatchName"
        } catch {
            $result = "Failed to create New MoveRequest for $userPrimarySmtpAddress $($_.Exception.Message)"
        }
        $results += [PSCustomObject]@{ User = $userPrimarySmtpAddress; Action = 'Create'; BatchName = $newBatchName; Result = $result }
        continue # Skip to the next iteration
    
    # Update the batch name if T-14 date has been changed
    if ($MoveRequest.BatchName -ne $newBatchName) {
        try {
            Set-MoveRequest -Identity $userPrimarySmtpAddress -BatchName $newBatchName -CompleteAfter "09/01/2025 5:00 PM" -PreventCompletion:$true -erroraction Stop
            $result = "MoveRequest for $userPrimarySmtpAddress updated to batch name $newBatchName"
        } catch {
            $result = "Failed to update MoveRequest for $userPrimarySmtpAddress $($_.Exception.Message)"
        }
        $results += [PSCustomObject]@{ User = $userPrimarySmtpAddress; Action = 'Update'; BatchName = $newBatchName; Result = $result }
        continue # Skip to the next iteration
    
    # Set the move request to complete, force that mailbox across the line.
    if ($BatchedUsers.CompleteMigration -eq $true) {
        try {
            Set-MoveRequest -Identity $userPrimarySmtpAddress -SkippedItemApprovalTime $([DateTime]::UtcNow) -SuspendWhenReadyToComplete:$false -PreventCompletion:$false -CompleteAfter:$null -BadItemLimit 10000 -LargeItemLimit 10000 -AcceptLargeDataLoss -WarningAction SilentlyContinue
            Get-MoveRequest $userPrimarySmtpAddress | Resume-MoveRequest
            $result = "Completing User $userPrimarySmtpAddress"
        } catch {
        $result = "Failed to set complete on $userPrimarySmtpAddress."
        }
        $results += [PSCustomObject]@{ User = $userPrimarySmtpAddress; Action = 'Complete'; BatchName = $newBatchName; Result = $result }
        continue # Skip to the next iteration
        }
    }
}
}
#convert Batching results to HTML for email
$resultsHtml = $results | ConvertTo-Html -As Table | Out-String

# Load necessary assemblies for runspace use
Add-Type -AssemblyName System.Threading
Add-Type -AssemblyName System.Collections.Concurrent

# Creating a concurrent queue to handle move requests
$UpdatedMoveRequests = Get-MoveRequest
$requestQueue = [System.Collections.Concurrent.ConcurrentQueue[Object]]::new()
foreach ($request in $UpdatedMoveRequests) {
    $requestQueue.Enqueue($request)
}

# Configure runspace pool
$minRunspaces = $minRunspacesset
$maxRunspaces = $maxRunspacesset
$pool = [runspacefactory]::CreateRunspacePool($minRunspaces, $maxRunspaces)
$pool.Open()

# Script block for processing each move request
$scriptBlock = {
    param ($request, $appid, $Tenantdomain, $certThumbprint)
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AppId $appid -Organization $Tenantdomain -CertificateThumbprint $certThumbprint
    $stats = $request | Get-MoveRequestStatistics
    $stats | Select-Object DisplayName, BatchName, Status, PercentComplete, DataConsistencyScore, BadItemsEncountered, LargeItemsEncountered, ExchangeGuid, LastSuccessfulSyncTimestamp, StatusDetail
}

# Runspace management
$runspaces = @()
while ($requestQueue.Count -gt 0) {
    if ($requestQueue.TryDequeue([ref]$request)) {
        $runspace = [powershell]::Create().AddScript($scriptBlock).AddArgument($request).AddArgument($appid).AddArgument($Tenantdomain).AddArgument($certThumbprint)
        $runspace.RunspacePool = $pool
        $runspaces += [PSCustomObject]@{ Pipe = $runspace; Status = $runspace.BeginInvoke() }
    }
}

# Collecting results
$MoveStats = @()
foreach ($runspace in $runspaces) {
    $MoveStats += $runspace.Pipe.EndInvoke($runspace.Status)
    $runspace.Pipe.Dispose()
}

# Close runspace pool
$pool.Close()
$pool.Dispose()

# Output the collected move statistics
#$MoveStats

# Script run counter so you can send email summaries on a different cadence. 
# Current cadence is every 6th script run send the same emails to a different set of email addresses.
# Path to the counter file
$counterPath = "C:\Scripts\runCounter.txt"
# Check if the counter file exists and read/increment the counter
if (Test-Path $counterPath) {
    $runCounter = Get-Content $counterPath | ForEach-Object { [int]$_ }
    $runCounter++
}
else {
    $runCounter = 1
}
Set-Content -Path $counterPath -Value $runCounter

# Generate HTML for email
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

# Decide which email addresses to send to based on counter
if ($runCounter % 6 -eq 0) {
    # Every 6th run, send to a different email address
    $toEmail = $every6thEmailRecipients
}
else {
    # Otherwise, send to the regular recipient list
    $toEmail = $emailRecipients
}

# Send the emails
Send-MailMessage -From $fromEmail -To $toEmail -Subject "Migration Batch Update Results" -Body $resultsHtml -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) -SmtpServer $smtpServer
Send-MailMessage -From $fromEmail -To $toEmail -Subject "Move Stats" -Body $MoveStatsHTML -BodyAsHtml -Encoding ([System.Text.Encoding]::UTF8) -SmtpServer $smtpServer

# Update SharePoint List with the new stats
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


# PowerShell script block to set SPO list to 100 percent complete for migrated users
$Results | Where-Object {$_.Result -like "*TargetUserAlreadyHasPrimaryMailboxException*"} | ForEach-Object {
    $userIdentity = $_.user  # User identity - already completed migration
    $mailbox = Get-Recipient -Identity $userIdentity
    $item = Get-PnPListItem -List "OnPremMailboxes" | Where-Object {$_.FieldValues.Title -eq $mailbox.ExchangeGuid}
    if ($item) {
        $fieldsToUpdate = @{
            "DataConsistencyScore" = $null; 
            "BatchName" = "Completed"; 
            "PercentComplete" = "100"; 
            "LargeItems" = $null; 
            "BadItems" = $null;
        }
        foreach ($field in $fieldsToUpdate.Keys) {
            Set-PnPListItem -List "OnPremMailboxes" -Identity $item.Id -Values @{$field = $fieldsToUpdate[$field]}
        }
        # Add completed user to the Intune_EXO group. Check before adding to be sure they are not already present.
        $ADUserDN = $mailbox.DistinguishedName
        $groupMembers = Get-ADGroupMember "Intune-EXO" | Select-Object -ExpandProperty DistinguishedName
        if ($ADUserDN -notin $groupMembers) {
            Get-Group "Intune-EXO" | Add-ADGroupMember -Members $ADUserDN

        # Remote call to trigger a delta sync in AAD Connect for the updated user
        # Replace with your AAD Connect server name
        $ScriptBlock = {
            Import-Module ADSync
            Start-ADSyncSyncCycle -PolicyType Delta
        }
        Invoke-Command -ComputerName $AADConnectServer -ScriptBlock $ScriptBlock -Credential $creds
    }
}
}

Get-PSSnapin | Remove-PSSnapin -Confirm:$false -ErrorAction SilentlyContinue
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-PnPOnline
Exit