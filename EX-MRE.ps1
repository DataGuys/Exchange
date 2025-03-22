# Load the EWS Managed API DLL
$dllpath = $env:ExchangeInstallPath + "Bin\Microsoft.Exchange.WebServices.dll"
Add-Type -Path $dllpath
$sizeThresholdGB = 75
$processAllLargeMailboxes = $false
Measure-Command{
Write-Host "Getting all mailboxes over the threshold, about 2 min tops..." -ForegroundColor Cyan
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue
$mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object { (Get-MailboxStatistics -Identity $_).TotalItemSize.Value.ToBytes() / $bytesToGB -gt $sizeThresholdGB }
$mailboxessorted = $mailboxes | Sort-Object
$largeMailboxList = $mailboxessorted
Write-Host "Finished geting all mailboxes over the threshold of $($sizeThresholdGB) GB" -ForegroundColor Cyan
Write-Host "Found $($largeMailboxList.count) mailboxes over the threshold" -ForegroundColor Green
}
# Prompt for credentials
$credential = Get-Credential -Credential gregory.halladmin@polsinelli.law

# Stack the Exchange EWS service connection variable
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016)
$service.Credentials = New-Object System.Net.NetworkCredential($credential.UserName, $credential.GetNetworkCredential().Password)
# Specify the EWS URL (replace with your Exchange server's EWS URL)
$service.Url = New-Object Uri("https://DC2-P-MAIL-01.polsinelli.law/ews/exchange.asmx")
$totalResults = New-Object System.Collections.ArrayList
# Process large mailboxes function
function Process-Mailbox {
    param (
        [string]$mailboxIdentity
    )

$results = New-Object System.Collections.ArrayList
    
    try {
        #Bind to EWS
        $rootFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $mailboxIdentity)
        $rootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $rootFolderId)
        #Get all folders under the root
        $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(5000)
        $folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
        $allFolders = $rootFolder.FindFolders($folderView)
        #Filter the folder list to those with emails
        $allfoldersfiltered = $allFolders | Where-Object {$_.FolderClass -eq "IPF.Note"}
        #Display the list of folders found
        #Look at each folder and get the items in them and the item size
        foreach ($folder in $allfoldersfiltered) {
            $sizeBefore = 0
            $sizesAfterRetention = @{}

            foreach ($years in 1, 2, 3, 5) {
                $sizesAfterRetention[$years] = 0
            }
            # Page view setting adjust the 3000 to adjust performance of Item view
            $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(10000)
            #Set the properties to only get the item properties we are after, in this instance is is the id, subject and size
            $propertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject, [Microsoft.Exchange.WebServices.Data.ItemSchema]::Size, [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeReceived)
            $view.PropertySet = $propertySet

            do {
                $findResults = $folder.FindItems($view)
                Write-Host "Found more items to review" -ForegroundColor Cyan
                foreach ($item in $findResults.Items) {
                    if ($item -is [Microsoft.Exchange.WebServices.Data.EmailMessage]) {
                        $sizeBefore += $item.Size
                    if ($null -ne $item.DateTimeReceived) {
                    $itemAgeDays = ((Get-Date) - $item.DateTimeReceived).TotalDays
                        } else {
                        }                   
                        foreach ($years in 1, 2, 3, 5) {
                            if ($itemAgeDays -le ($years * 365)) {
                                $sizesAfterRetention[$years] += $item.Size
                            }
                        }
                    } else {
                        write-host "Skipping non-email item..." $item -foregroundcolor yellow
                    }
                }
                $view.Offset += $findResults.Items.Count
                Write-Host "Working on more results in mailbox $($mailboxIdentity) $($view.Offset)" -ForegroundColor Yellow
            } while ($findResults.MoreAvailable)

            $result = New-Object PSObject -Property @{
                Mailbox = $mailboxIdentity
                Folder = $folder.DisplayName
                SizeBeforeRetentionGB = [Math]::Round($sizeBefore / $bytesToGB, 2)
                SizeAfter1YearGB = [Math]::Round($sizesAfterRetention[1] / $bytesToGB, 2)
                SizeAfter2YearsGB = [Math]::Round($sizesAfterRetention[2] / $bytesToGB, 2)
                SizeAfter3YearsGB = [Math]::Round($sizesAfterRetention[3] / $bytesToGB, 2)
                SizeAfter5YearsGB = [Math]::Round($sizesAfterRetention[5] / $bytesToGB, 2)
            }

            [void]$results.Add($result)
        }
    } catch {
        Write-Host "An error occurred for $mailboxIdentity $_" -ForegroundColor Red
    }
    return $results
}

####### Main Script ##############
Set-Location $env:USERPROFILE\Documents 
Measure-Command {
foreach ($mailbox in $largeMailboxList) {
    if ($processAllLargeMailboxes -eq $false) {
        Write-Host "Processing Only The First Large Mailbox: $($largeMailboxList[0].PrimarySmtpAddress)" -ForegroundColor Cyan
        $FirstLargeMailboxResults = Process-Mailbox -mailboxIdentity $($largeMailboxList[0].PrimarySmtpAddress)
        Write-Host "Completed First Large Mailbox analysis" -ForegroundColor Yellow
        $FirstLargeMailboxResultsformated = $FirstLargeMailboxResults | Select-Object Folder, SizeBeforeRetentionGB, SizeAfter1YearGB, SizeAfter2YearsGB, SizeAfter3YearsGB, SizeAfter5YearsGB, Mailbox | Sort SizeBeforeRetentionGB -Descending
        $totalResults += $FirstLargeMailboxResultsformated
        #$FirstLargeMailboxResultsformated | ft -AutoSize
        $FirstLargeCSVPath = "$env:USERPROFILE\Documents\RetentionSizeEstimate_$($largeMailboxList[0].PrimarySmtpAddress.Address.Replace('@', '_')).csv"
        $FirstLargeMailboxResultsformated | Export-csv -Path $FirstLargeCSVPath -NoTypeInformation
        Write-Host "First Large Mailbox Results for $($largeMailboxList[0].Mailbox.Address) exported to $FirstLargeCSVPath" -ForegroundColor Green
        break
    } else {
        Write-Host "Processing all large mailboxes, working on... $($mailbox.PrimarySmtpAddress)" -ForegroundColor Cyan
        $LargemailboxResults = Process-Mailbox -mailboxIdentity $mailbox.PrimarySmtpAddress
        $LargeMailboxResultsSorted = $LargeMailboxResults | Select-Object Folder, SizeBeforeRetentionGB, SizeAfter1YearGB, SizeAfter2YearsGB, SizeAfter3YearsGB, SizeAfter5YearsGB, Mailbox | Sort SizeBeforeRetentionGB -Descending
        $totalResults += $LargemailboxResultsSorted

        # Export individual mailbox results to CSV
        $LargeindividualCsvPath = "$env:USERPROFILE\Documents\RetentionSizeEstimate_$($mailbox.Mailbox.Replace('@', '_')).csv"
        $LargemailboxResultsSorted | Export-Csv -Path $LargeindividualCsvPath -NoTypeInformation
        Write-Host "Individual results for $($mailbox.PrimarySmtpAddress) exported to $LargeindividualCsvPath" -ForegroundColor Green
    }
}
}
# Export total results to CSV
$csvPath = "$env:USERPROFILE\Documents\TotalRetentionSizeEstimateOutput.csv"
$totalResults | Export-Csv -Path $csvPath -NoTypeInformation
Write-Host "Total results exported to CSV at $csvPath" -ForegroundColor Green