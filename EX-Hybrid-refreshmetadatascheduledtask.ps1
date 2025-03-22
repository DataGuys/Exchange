Schtasks /create `
  /sc Daily `
  /tn "FedRefresh" `
  /tr "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -command `
    'Add-PSSnapIn Microsoft.Exchange.Management.PowerShell.E2010; `
    $fedTrust = Get-FederationTrust; `
    Set-FederationTrust -Identity `$fedTrust.Name -RefreshMetadata'"
