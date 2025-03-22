Install-Module ExchangeOnlineManagement
Import-Module ExchangeOnlineManagement
Connect-IPPSSession
$Search=New-ComplianceSearch -Name "Remove Phishing Message" -ExchangeLocation All -ContentMatchQuery '(Received:9/27/2023..9/28/2023) AND (From:brian.morrow@ihydrant.com) AND (Subject:"New Transmission Main Work*")'
Start-ComplianceSearch -Identity $Search.Identity

New-ComplianceSearchAction -SearchName "Remove Phishing Message" -Purge -PurgeType HardDelete

