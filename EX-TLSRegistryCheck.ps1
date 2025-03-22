# Import Exchange Management Shell if not already imported
if (-not (Get-PSSnapin -Name "Microsoft.Exchange.Management.PowerShell.SnapIn" -ErrorAction SilentlyContinue)) {
    Add-PSSnapin "Microsoft.Exchange.Management.PowerShell.SnapIn"
}

# Initialize an empty array to store the results
$results = @()

# Automatically get the list of Exchange Servers
$exchangeServers = Get-ExchangeServer | Where-Object {$_.AdminDisplayVersion -like "Version 15.1*"} | ForEach-Object {$_.Name}

# Loop through each Exchange Server
foreach ($server in $exchangeServers) {
    Write-Host "Checking $server ..."

    # Define the registry keys and values to check
    $keysToCheck = @(
        # Existing keys from your list
        @{
            "Path" = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client"
            "Values" = @("DisabledByDefault", "Enabled")
        },
        # ... [Your other existing keys here]

        # Additional keys for Windows Server 2016 and Exchange 2016
        @{
            "Path" = "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319"
            "Values" = @("SchUseStrongCrypto")
        },
        @{
            "Path" = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319"
            "Values" = @("SchUseStrongCrypto")
        }
    )

    # Loop through each registry key and value
    foreach ($key in $keysToCheck) {
        $keyPath = $key.Path
        $keyValues = $key.Values

        foreach ($valueName in $keyValues) {
            $fullPath = "\\$server\$($keyPath.Replace(':', '$'))"
            $value = ""

            try {
                $value = Invoke-Command -ComputerName $server -ScriptBlock {
                    Get-ItemPropertyValue -Path $using:keyPath -Name $using:valueName
                }
            }
            catch {
                $value = "Not Found"
            }

            # Create an object to store the result
            $result = New-Object PSObject -Property @{
                "Server" = $server
                "RegistryKey" = $keyPath
                "RegistryValue" = $valueName
                "Value" = $value
            }

            # Add the result to the results array
            $results += $result
        }
    }
}

# Export results to a CSV file
$results | Select-Object Server, RegistryKey, RegistryValue, Value | Export-Csv -Path "ExchangeServers_RegistryCheck.csv" -NoTypeInformation
Write-Host "Exported results to ExchangeServers_RegistryCheck.csv"
