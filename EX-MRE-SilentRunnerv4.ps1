Clear-Host
$credential = Get-StoredCredential -Target "silentrunner"
# Define the output folder path
$outputFolderPath = "$env:UserProfile\Documents\Results\MRE"
# Check if the path exists
if (-not (Test-Path -Path $outputFolderPath)) {
    # Path does not exist, so create it
    New-Item -Path $outputFolderPath -ItemType Directory
    Write-Host "Created directory: $outputFolderPath" -ForegroundColor Cyan
} else {
    # Path already exists
    Write-Host "Directory already exists: $outputFolderPath" -ForegroundColor Green
}
# Load the EWS Managed API DLL
$dllpath = "$env:ExchangeInstallPath\Bin\Microsoft.Exchange.WebServices.dll"
Add-Type -Path $dllpath
Set-Location $outputFolderPath
Remove-Item MRE-*.csv -Recurse

############################ Functions and encapsulated script block for parallel processing ####################
$sizeThresholdGB = 75
Write-Host "Getting all mailboxes over the threshold $($sizeThresholdGB) GB, about 2 min..." -ForegroundColor Cyan
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue
$mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {(Get-MailboxStatistics -Identity $_).TotalItemSize.Value.ToBytes() / 1GB -gt $sizeThresholdGB}
$mailboxessorted = $mailboxes | Sort-Object
$largeMailboxList = $mailboxessorted
$largeMailboxList | Select-Object DisplayName, PrimarySmtpAddress | Export-Csv  $outputFolderPath\largemailboxlist.csv -NoTypeInformation
Write-Host "Finished geting all mailboxes over the threshold of $($sizeThresholdGB) GB" -ForegroundColor Cyan
Write-Host "Found $($largeMailboxList.count) mailboxes over the threshold" -ForegroundColor Green
# Retrieve Exchange Server FQDNs and construct EWS Endpoints
try {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue
    $exchangeServers = Get-ExchangeServer | Where-Object { $_.IsClientAccessServer -eq $true }
    $ewsEndpoints = $exchangeServers | ForEach-Object { "https://$($_.Fqdn)/ews/exchange.asmx" }
} catch {
    Write-Host "Error retrieving Exchange servers: $_" -ForegroundColor Red
    break
}

# Validate that EWS endpoints were retrieved
if (-not $ewsEndpoints -or $ewsEndpoints.Count -eq 0) {
    Write-Host "No EWS endpoints found. Exiting script." -ForegroundColor Red
    break
}

$email = $null
$scriptBlock = {
param($email, $EWSURL)

# Define necessary variables
#$sizeThresholdGB = 75
# Recreate the EWS service connection inside the runspace
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016)
$service.Credentials = New-Object System.Net.NetworkCredential($credential.UserName, $credential.GetNetworkCredential().Password)
$service.Url = New-Object Uri($EWSURL)

# Process large mailboxes function
function Process-Mailbox {
    param (
        [string]$email
    )

$results = New-Object System.Collections.ArrayList
    
    try {
        #Bind to EWS
        $rootFolderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $email)
        $rootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $rootFolderId)
        #Get all folders under the root
        $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(2000)
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
            $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(2000)
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
                        $itemAgeDays = ((Get-Date) - $item.DateTimeReceived).TotalDays } else {
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
                Write-Host "Working on more results in mailbox $($email) $($view.Offset)" -ForegroundColor Magenta
            } while ($findResults.MoreAvailable)
                $result = New-Object PSObject -Property @{
                Mailbox = $email
                Folder = $folder.DisplayName
                SizeBeforeRetentionGB = [Math]::Round($sizeBefore / 1GB, 2)
                SizeAfter1YearGB = [Math]::Round($sizesAfterRetention[1] / 1GB, 2)
                SizeAfter2YearsGB = [Math]::Round($sizesAfterRetention[2] / 1GB, 2)
                SizeAfter3YearsGB = [Math]::Round($sizesAfterRetention[3] / 1GB, 2)
                SizeAfter5YearsGB = [Math]::Round($sizesAfterRetention[5] / 1GB, 2)
            }

            [void]$results.Add($result)
        }
    } catch {
        Write-Host "An error occurred for $email $_" -ForegroundColor Red
    }
    return $results
}
# Run the Process-Mailbox Function inside the script block
# Process the mailbox and return results
    $results = Process-Mailbox -email $email
    # Construct the CSV file path
    $IndividualCsvPath = "$env:USERPROFILE\Documents\Results\MRE\MRE-$($email.Replace('@', '_')).csv"
    # Export the results to a CSV file
    $resultssorted = $results | Select-Object Folder,SizeBeforeRetentionGB,SizeAfter1YearGB,SizeAfter2YearsGB,SizeAfter3YearsGB,SizeAfter5YearsGB,Mailbox | Sort-Object SizeBeforeRetentionGB -Descending
    $resultssorted | Export-Csv -Path $IndividualCsvPath -NoTypeInformation
    # Optionally, you can return the path of the CSV file as part of the result
    return @{ 'Results' = $results; 'CsvPath' = $IndividualCsvPath }
}
#Test Script block with a single email if needed.
#& $scriptBlock -email "ajackson@polsinelli.com"

############### Jobs Area #################################
# Cleanup any old runspaces using the runspacepool variable
$runspacePool = $null
# Create runspace pool max spaces based on number of processors available.
$minRunspaces = 1
$maxRunspaces = 4
$runspacePool = [runspacefactory]::CreateRunspacePool($minRunspaces, $maxRunspaces)
$runspacePool.Open()

# Create and invoke runspace jobs
$runspaceJobs =  New-Object System.Collections.ArrayList
$endpointIndex = 0

foreach ($mailbox in $largeMailboxList) {
    $email = $mailbox.PrimarySmtpAddress.Address
    # Round-robin distribution of EWS endpoints
    $EWSURL = $ewsEndpoints[$endpointIndex % $ewsEndpoints.Count]
    $endpointIndex++
    Write-Host "Creating Job for $($mailbox.PrimarySmtpAddress.Address)" -ForegroundColor Cyan
    $powershell = [powershell]::Create().AddScript($scriptBlock).AddArgument($email).AddArgument($EWSURL)
    $powershell.RunspacePool = $runspacePool
    $runspaceJobs += [PSCustomObject]@{
        Pipe = $powershell
        Mailbox = $email
        Job = $powershell.BeginInvoke()
    }
}

# Visual Job Monitor
Write-Host "Monitoring jobs..." -ForegroundColor Yellow

# Initialize counters and job duration tracker
$jobDurations = @{}

do {
    Clear-Host
    Write-Host "Job Status Monitor" -ForegroundColor Cyan
    Write-Host "==================" -ForegroundColor Cyan

    # Reset counters
    $runningJobs = 0
    $completedJobs = 0
    $erroredJobs = 0

    foreach ($job in Get-Job) {
        $status = $job.State

        switch ($status) {
            "Completed" {
                $completedJobs++
                if ($null -ne $job.PSEndTime -and $null -ne $job.PSBeginTime) {
                    $jobDurations[$job.Id] = $job.PSEndTime - $job.PSBeginTime
                }
            }
            "Failed" { $erroredJobs++ }
            default { $runningJobs++ }
        }
    }

    # Calculate average duration
    if ($jobDurations.Count -gt 0) {
        $totalDuration = [TimeSpan]::Zero
        foreach ($duration in $jobDurations.Values) {
            $totalDuration += $duration
        }
        $averageDuration = $totalDuration / $jobDurations.Count
    } else {
        $averageDuration = [TimeSpan]::Zero
    }

    Write-Host "`nRunning Jobs: $runningJobs" -foregroundcolor Cyan
    Write-Host "Completed Jobs: $completedJobs" -ForegroundColor Green
    Write-Host "Errored Jobs: $erroredJobs" -ForegroundColor Yellow
    Write-Host "Average Job Duration: $($averageDuration.ToString())"

    Start-Sleep -Seconds 10  # Refresh every 10 seconds
} while ($runningJobs -gt 0)

# Export job durations to CSV
$exportData = $jobDurations.GetEnumerator() | ForEach-Object {
    [PSCustomObject]@{
        JobId = $_.Key
        Duration = $_.Value.ToString()
    }
}
$exportData | Export-Csv -Path "JobDurationsReport.csv" -NoTypeInformation
Write-Host "Job durations exported to 'JobDurationsReport.csv'" -ForegroundColor Green


# Initialize an empty hash table to store the aggregated results
$mailboxSizes = @{}
# Iterate over each CSV file in the folder
$allCSVFiles = Get-ChildItem -Path $outputFolderPath -Filter "MRE-*.csv"
$csvFiles = $allCSVFiles | Where-Object {$_.Name -ne "MRE-Summary.csv"} | Sort-Object Name
$csvFiles | ForEach-Object {
    # Read the CSV file
    $csvData = Import-Csv -Path $_.Name

    # Aggregate the sizes for each mailbox
    foreach ($row in $csvData) {
        # Check if the mailbox already exists in the hash table
        if (-not $mailboxSizes.ContainsKey($row.Mailbox)) {
            $mailboxSizes[$row.Mailbox] = [PSCustomObject]@{
                Mailbox = $row.Mailbox
                SizeBeforeRetentionGB = 0
                SizeAfter1YearGB = 0
                SizeAfter2YearsGB = 0
                SizeAfter3YearsGB = 0
                SizeAfter5YearsGB = 0
            }
        }

        # Sum the sizes for the mailbox
        $mailboxSizes[$row.Mailbox].SizeBeforeRetentionGB += [double]$row.SizeBeforeRetentionGB
        $mailboxSizes[$row.Mailbox].SizeAfter1YearGB += [double]$row.SizeAfter1YearGB
        $mailboxSizes[$row.Mailbox].SizeAfter2YearsGB += [double]$row.SizeAfter2YearsGB
        $mailboxSizes[$row.Mailbox].SizeAfter3YearsGB += [double]$row.SizeAfter3YearsGB
        $mailboxSizes[$row.Mailbox].SizeAfter5YearsGB += [double]$row.SizeAfter5YearsGB
    }
}
$outputFolderPath = "$env:UserProfile\Documents\Results\MRE"
# Output the aggregated results
$mailboxSizes.Values | Export-Csv ($outputFolderPath + "\MRE-Summary.csv") -NoTypeInformation
############## Email Results
# Define file paths and email parameters
$zipFilePath = "$outputFolderPath\MRE-Files$(Get-Date -Format 'yyyyMMdd').zip"
$csvFiles = Get-ChildItem -Path $outputFolderPath -Filter "MRE-*.csv"
# Compress the CSV files into a zip file
Compress-Archive -Path $csvFiles.FullName -DestinationPath $zipFilePath -Force

# Email parameters
$smtpServer = "localhost" # Replace with your SMTP server
$from = "MRE@polsinelli.com" # Replace with your sender email address
$to = "7431cf4d.helient.com@amer.teams.ms" # Replace with the recipient's email address
$subject = "POL MRE CSV Files"
$body = "Attached are the latest MRE CSV files."
# Send the email with the zip file attached
Send-MailMessage -From $from -To $to -Subject $subject -Body $body -Attachments $zipFilePath -SmtpServer $smtpServer -UseSsl -Port 25 -Credential $credential
############################ End Functions and encapsulated script block for parallel processing ####################

# Cleanup
$runspacePool.Close()
$runspacePool.Dispose()