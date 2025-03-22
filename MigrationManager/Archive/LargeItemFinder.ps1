if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}
Get-Mailbox -ResultSize Unlimited | Get-MailboxFolderStatistics -IncludeAnalysis -FolderScope All | Where-Object {(($_.TopSubjectSize -Match "MB") -and ($_.TopSubjectSize -GE 150.0))} | Select-Object UserPrincipalName, Identity, TopSubject, TopSubjectSize | Export-CSV -path "C:\Scripts\LargeItemreport.csv" -notype