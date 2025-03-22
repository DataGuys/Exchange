# Function Definitions

# Get-ServerResourceUtilization
function Get-ServerResourceUtilization {
    $serverResources = Get-WmiObject -Class Win32_OperatingSystem | Select-Object CSName, @{Name="CPUUsage";Expression={(Get-Counter "\Processor(_Total)\% Processor Time").CounterSamples.CookedValue}}, @{Name="FreeMemory";Expression={($_.FreePhysicalMemory/1024)}}, @{Name="FreeSpaceInC";Expression={(Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'").FreeSpace/1GB}}
    return $serverResources
}

# Get-DatabaseStatus
function Get-DatabaseStatus {
    $databases = Get-MailboxDatabase -Status | Select-Object Name, Server, Mounted, DatabaseSize
    return $databases
}

# Get-DatabaseWhiteSpace
function Get-DatabaseWhiteSpace {
    $whiteSpace = Get-MailboxDatabase -Status | ForEach-Object { [PSCustomObject] @{ Database = $_.Name; AvailableNewMailboxSpace = $_.AvailableNewMailboxSpace } }
    return $whiteSpace
}

# Test-ClientConnectivity
function Test-ClientConnectivity {
    $connectivity = Test-MAPIConnectivity | Select-Object MailboxServer, Mailbox, Result, Error
    return $connectivity
}

# Test-MailFlow
function Test-MailFlow {
    $mailFlow = Test-Mailflow | Select-Object TestMailflowResult, MessageLatencyTime
    return $mailFlow
}

# Get-TransportQueue
function Get-TransportQueue {
    $transportQueues = Get-Queue | Select-Object Identity, Status, MessageCount, NextHopDomain
    return $transportQueues
}

# Get-MailboxStatisticsDetail
function Get-MailboxStatisticsDetail {
    $mailboxStats = Get-Mailbox -ResultSize Unlimited | Get-MailboxStatistics | Select-Object DisplayName, TotalItemSize, ItemCount, LastLogonTime
    return $mailboxStats
}

# Get-CASHealth
function Get-CASHealth {
    $owaConnectivity = Test-OwaConnectivity -MonitoringContext:$true | Select-Object MailboxServer, Url, Result
    $ecpConnectivity = Test-EcpConnectivity -MonitoringContext:$true | Select-Object MailboxServer, Url, Result
    $activesyncConnectivity = Test-ActiveSyncConnectivity -MonitoringContext:$true | Select-Object MailboxServer, Url, Result

    [PSCustomObject]@{
        OWAConnectivity = $owaConnectivity
        ECPConnectivity = $ecpConnectivity
        ActiveSyncConnectivity = $activesyncConnectivity
    }
}

# Get-BackupRestoreStatus
function Get-BackupRestoreStatus {
    $databases = Get-MailboxDatabase -Status
    $backupStatus = $databases | Select-Object Name, LastFullBackup, LastIncrementalBackup, LastDifferentialBackup, LastCopyBackup

    return $backupStatus
}

# Get-ExchangeSecurityStatus
function Get-ExchangeSecurityStatus {
    $certificates = Get-ExchangeCertificate | Select-Object Thumbprint, Services, NotAfter
    $virtualDirectories = Get-OwaVirtualDirectory | Select-Object Name, InternalUrl, ExternalUrl

    [PSCustomObject]@{
        Certificates = $certificates
        VirtualDirectories = $virtualDirectories
    }
}

# Get-ExchangePerformanceCounters
function Get-ExchangePerformanceCounters {
    $performanceCounters = @("\MSExchange Database(\*)\I/O Database Reads (Attached) per sec",
                            "\MSExchange Database(\*)\I/O Database Writes (Attached) per sec",
                            "\MSExchange Transport Queues(_total)\Active Mailbox Delivery Queue Length",
                            "\Processor(_Total)\% Processor Time")

    $counterData = $performanceCounters | ForEach-Object {
        Get-Counter -Counter $_ -SampleInterval 1 -MaxSamples 1
    }

    return $counterData
}

# Analyze-ExchangeLogs
function Get-ExchangeLogs {
    $logFilePath = "C:\Path\To\Your\Logs"  # Specify the log file path
    $logEntries = Get-Content $logFilePath | Select-String -Pattern "Error", "Warning" -Context 0,2

    return $logEntries
}

# Get-ExchangeServiceStatus
function Get-ExchangeServiceStatus {
    $exchangeServices = Get-Service *exchange* | Select-Object Name, Status

    return $exchangeServices
}

# Main Script Block for Remote Execution
$scriptBlock = {
    param($ServerName)

    # Ensure Exchange snap-in is loaded
    if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
    }

    # Perform the checks
    $transportQueue = Get-TransportQueue
    $mailboxStats = Get-MailboxStatisticsDetail
    $casHealth = Get-CASHealth
    $backupStatus = Get-BackupRestoreStatus
    $securityCheck = Get-ExchangeSecurityStatus
    $performanceCounters = Get-ExchangePerformanceCounters
    $logAnalysis = Get-ExchangeLogs
    $serviceStatus = Get-ExchangeServiceStatus
    $serverUtilization = Get-ServerResourceUtilization
    $dbStatus = Get-DatabaseStatus
    $dbWhiteSpace = Get-DatabaseWhiteSpace
    $clientConnectivity = Test-ClientConnectivity
    $mailFlowTest = Test-MailFlow
    
    # Compile results into a single object with Server Name
    [PSCustomObject]@{
        ServerName = $ServerName
        TransportQueue = $transportQueue
        MailboxStatistics = $mailboxStats
        CASHealth = $casHealth
        BackupRestoreStatus = $backupStatus
        Security = $securityCheck
        PerformanceCounters = $performanceCounters
        LogAnalysis = $logAnalysis
        ExchangeServiceStatus = $serviceStatus
        ServerResourceUtilization = $serverUtilization
        DatabaseStatus = $dbStatus
        DatabaseWhiteSpace = $dbWhiteSpace
        ClientConnectivity = $clientConnectivity
        MailFlowTest = $mailFlowTest
    }
}

# Find Exchange Servers in Active Directory and Execute Script Block Remotely
$ExchangeServers = Get-ADComputer -Filter {OperatingSystem -Like "*Exchange*"} | Select-Object -ExpandProperty Name

foreach ($server in $ExchangeServers) {
    Invoke-Command -ComputerName $server -ScriptBlock $scriptBlock -ArgumentList $server
}
# Define Health Checker script path
$HealthCheckerScriptPath = $env:TEMP + "\HealthChecker.ps1"  # Update with actual path
Set-Location $env:TEMP
# Download HealthChecker script if not present
if (-not (Test-Path $HealthCheckerScriptPath)) {
    $HealthCheckerScriptUrl = "https://github.com/microsoft/CSS-Exchange/releases/latest/download/HealthChecker.ps1"
    Invoke-WebRequest -Uri $HealthCheckerScriptUrl -OutFile $HealthCheckerScriptPath
}

# Add Exchange snap-in if not already loaded
if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}
# Fetch Exchange Servers
$ExchangeServers = Get-ExchangeServer | Select-Object -ExpandProperty Name

# Run Health Checker on all Exchange servers
.\HealthChecker.ps1 -Server $ExchangeServers

# Optional: Build HTML report for all servers
.\HealthChecker.ps1 -BuildHtmlServersReport
