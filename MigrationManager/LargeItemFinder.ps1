# Ensure the Exchange Management PowerShell snap-in is loaded for on-premises Exchange
if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}

# Script to find large items in on-premises Exchange mailboxes
$LargeItemReportPath = "C:\Scripts\LargeItemreport.csv"
$LargeItems = @()

Get-Mailbox -ResultSize Unlimited | ForEach-Object {
    $mailbox = $_
    Write-Host "Scanning $($mailbox.PrimarySmtpAddress)" -ForegroundColor Cyan
    $largeItemsInMailbox = Get-MailboxFolderStatistics -Identity ashipman@polsinelli.com -IncludeAnalysis -FolderScope All -IncludeOldestAndNewestItems | Where-Object {
        # Extract the numeric part of the TopSubjectSize and convert it to a number
        $sizeValue = $_.TopSubjectSize -replace "[^\d.]", ""
        # Check if size is in MB and greater than or equal to 150 MB
        if ($_.TopSubjectSize -match "MB" -and [double]$sizeValue -ge 150.0) {
            $true
        }
        else {
            $false
        }
    } | Select-Object @{Name="UserPrincipalName"; Expression={$mailbox.UserPrincipalName}}, Identity, TopSubject, TopSubjectSize, OldestItemReceivedDate, NewestItemReceivedDate

    $LargeItems += $largeItemsInMailbox
}

# Export the collected large item details to a CSV file
$LargeItems | Export-CSV -Path $LargeItemReportPath -NoTypeInformation

Write-Host "Large item report generated at $LargeItemReportPath" -ForegroundColor Green

