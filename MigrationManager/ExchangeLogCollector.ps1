<#
    MIT License

    Copyright (c) Microsoft Corporation.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
#>

# Version 24.03.28.2048

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Value is used')]
[CmdletBinding(DefaultParameterSetName = "LogAge", SupportsShouldProcess, ConfirmImpact = "High")]
param (
    [string]$FilePath = "C:\MS_Logs_Collection",
    [Parameter(Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [Alias('Fqdn')]
    [string[]]$Servers = @($env:COMPUTERNAME),
    [switch]$AcceptedRemoteDomain,
    [switch]$ADDriverLogs,
    [bool]$AppSysLogs = $true,
    [bool]$AppSysLogsToXml = $true,
    [switch]$AutoDLogs,
    [switch]$CollectFailoverMetrics,
    [switch]$DAGInformation,
    [switch]$DailyPerformanceLogs,
    [switch]$DefaultTransportLogging,
    [switch]$EASLogs,
    [switch]$ECPLogs,
    [switch]$EWSLogs,
    [Alias("ExchangeServerInfo")]
    [switch]$ExchangeServerInformation,
    [switch]$ExMon,
    [switch]$ExPerfWiz,
    [switch]$FrontEndConnectivityLogs,
    [switch]$FrontEndProtocolLogs,
    [switch]$GetVDirs,
    [switch]$HighAvailabilityLogs,
    [switch]$HubConnectivityLogs,
    [switch]$HubProtocolLogs,
    [switch]$IISLogs,
    [switch]$ImapLogs,
    [switch]$MailboxAssistantsLogs,
    [switch]$MailboxConnectivityLogs,
    [switch]$MailboxDeliveryThrottlingLogs,
    [switch]$MailboxProtocolLogs,
    [Alias("ManagedAvailability")]
    [switch]$ManagedAvailabilityLogs,
    [switch]$MapiLogs,
    [switch]$MessageTrackingLogs,
    [switch]$MitigationService,
    [switch]$OABLogs,
    [switch]$OrganizationConfig,
    [switch]$OWALogs,
    [switch]$PipelineTracingLogs,
    [switch]$PopLogs,
    [switch]$PowerShellLogs,
    [switch]$QueueInformation,
    [switch]$ReceiveConnectors,
    [switch]$RPCLogs,
    [switch]$SearchLogs,
    [switch]$SendConnectors,
    [Alias("ServerInfo")]
    [switch]$ServerInformation,
    [switch]$TransportAgentLogs,
    [switch]$TransportConfig,
    [switch]$TransportRoutingTableLogs,
    [switch]$TransportRules,
    [switch]$WindowsSecurityLogs,
    [switch]$AllPossibleLogs,
    [Alias("CollectAllLogsBasedOnDaysWorth")]
    [bool]$CollectAllLogsBasedOnLogAge = $true,
    [switch]$ConnectivityLogs,
    [switch]$DatabaseFailoverIssue,
    [Parameter(ParameterSetName = "Worth")]
    [int]$DaysWorth = 3,
    [Parameter(ParameterSetName = "Worth")]
    [int]$HoursWorth = 0,
    [switch]$DisableConfigImport,
    [string]$ExMonLogmanName = "ExMon_Trace",
    [array]$ExPerfWizLogmanName = @("Exchange_PerfWiz", "ExPerfWiz", "SimplePerf"),
    [Parameter(ParameterSetName = "LogAge")]
    [TimeSpan]$LogAge = "3.00:00:00",
    [Parameter(ParameterSetName = "LogPeriod")]
    [DateTime]$LogStartDate = (Get-Date).AddDays(-3),
    [Parameter(ParameterSetName = "LogPeriod")]
    [DateTime]$LogEndDate = (Get-Date),
    [switch]$OutlookConnectivityIssues,
    [switch]$PerformanceIssues,
    [switch]$PerformanceMailFlowIssues,
    [switch]$ProtocolLogs,
    [switch]$ScriptDebug,
    [bool]$SkipEndCopyOver
)

begin {


function Enter-YesNoLoopAction {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = "High")]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Question,

        [Parameter(Mandatory = $false)]
        [string]$Target = $env:COMPUTERNAME,

        [Parameter(Mandatory = $true)]
        [ScriptBlock]$YesAction,

        [Parameter(Mandatory = $true)]
        [ScriptBlock]$NoAction
    )

    Write-Verbose "Calling: Enter-YesNoLoopAction"
    Write-Verbose "Passed: [string]Question: $Question"

    if ($PSCmdlet.ShouldProcess($Target, $Question)) {
        & $YesAction
    } else {
        & $NoAction
    }
}

function Import-ScriptConfigFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path $_ })]
        [string]$ScriptConfigFileLocation
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    Write-Verbose "Passed: [string]ScriptConfigFileLocation: '$ScriptConfigFileLocation'"

    try {
        $content = Get-Content $ScriptConfigFileLocation -ErrorAction Stop
        $jsonContent = $content | ConvertFrom-Json
    } catch {
        throw "Failed to convert ScriptConfigFileLocation from a json type object."
    }

    $jsonContent |
        Get-Member |
        Where-Object { $_.Name -ne "Method" } |
        ForEach-Object {
            Write-Verbose "Adding variable $($_.Name)"
            Set-Variable -Name $_.Name -Value ($jsonContent.$($_.Name)) -Scope Script
        }
}

    #Need to do this here otherwise can't find the script path
    $configPath = "{0}\{1}.json" -f (Split-Path -Parent $MyInvocation.MyCommand.Path), (Split-Path -Leaf $MyInvocation.MyCommand.Path)

    if ((Test-Path $configPath) -and
        !$DisableConfigImport) {
        try {
            Import-ScriptConfigFile -ScriptConfigFileLocation $configPath
        } catch {
            # can't monitor this because monitor needs to start in the end function.
            Write-Host "Failed to load the config file at $configPath. `r`nPlease update the config file to be able to run 'ConvertFrom-Json' against it" -ForegroundColor "Red"
            Enter-YesNoLoopAction -Question "Do you wish to continue?" -YesAction {} -NoAction { exit }
        }
    }

    $BuildVersion = "24.03.28.2048"
    $serversToProcess = New-Object System.Collections.ArrayList
}

process {
    foreach ($server in $Servers) {
        [void]$serversToProcess.Add($server)
    }
}

end {

    Write-Host "Exchange Log Collector v$($BuildVersion)"
    # Used throughout the script for checking for free space available.
    $Script:StandardFreeSpaceInGBCheckSize = 10

    if ($PSBoundParameters["Verbose"]) { $Script:ScriptDebug = $true }

    if ($PSCmdlet.ParameterSetName -eq "Worth") {
        $Script:LogAge = New-TimeSpan -Days $DaysWorth -Hours $HoursWorth
        $Script:LogEndAge = New-TimeSpan -Days 0 -Hours 0
    }

    if ($PSCmdlet.ParameterSetName -eq "LogPeriod") {
        $Script:LogAge = ((Get-Date) - $LogStartDate)
        $Script:LogEndAge = ((Get-Date) - $LogEndDate)
    } else {
        $Script:LogEndAge = New-TimeSpan -Days 0 -Hours 0
    }

    function Invoke-RemoteFunctions {
        param(
            [Parameter(Mandatory = $true)][object]$PassedInfo
        )


function Write-Host {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'Proper handling of write host with colors')]
    [CmdletBinding()]
    param(
        [Parameter(Position = 1, ValueFromPipeline)]
        [object]$Object,
        [switch]$NoNewLine,
        [string]$ForegroundColor
    )
    process {
        $consoleHost = $host.Name -eq "ConsoleHost"

        if ($null -ne $Script:WriteHostManipulateObjectAction) {
            $Object = & $Script:WriteHostManipulateObjectAction $Object
        }

        $params = @{
            Object    = $Object
            NoNewLine = $NoNewLine
        }

        if ([string]::IsNullOrEmpty($ForegroundColor)) {
            if ($null -ne $host.UI.RawUI.ForegroundColor -and
                $consoleHost) {
                $params.Add("ForegroundColor", $host.UI.RawUI.ForegroundColor)
            }
        } elseif ($ForegroundColor -eq "Yellow" -and
            $consoleHost -and
            $null -ne $host.PrivateData.WarningForegroundColor) {
            $params.Add("ForegroundColor", $host.PrivateData.WarningForegroundColor)
        } elseif ($ForegroundColor -eq "Red" -and
            $consoleHost -and
            $null -ne $host.PrivateData.ErrorForegroundColor) {
            $params.Add("ForegroundColor", $host.PrivateData.ErrorForegroundColor)
        } else {
            $params.Add("ForegroundColor", $ForegroundColor)
        }

        Microsoft.PowerShell.Utility\Write-Host @params

        if ($null -ne $Script:WriteHostDebugAction -and
            $null -ne $Object) {
            &$Script:WriteHostDebugAction $Object
        }
    }
}

function SetProperForegroundColor {
    $Script:OriginalConsoleForegroundColor = $host.UI.RawUI.ForegroundColor

    if ($Host.UI.RawUI.ForegroundColor -eq $Host.PrivateData.WarningForegroundColor) {
        Write-Verbose "Foreground Color matches warning's color"

        if ($Host.UI.RawUI.ForegroundColor -ne "Gray") {
            $Host.UI.RawUI.ForegroundColor = "Gray"
        }
    }

    if ($Host.UI.RawUI.ForegroundColor -eq $Host.PrivateData.ErrorForegroundColor) {
        Write-Verbose "Foreground Color matches error's color"

        if ($Host.UI.RawUI.ForegroundColor -ne "Gray") {
            $Host.UI.RawUI.ForegroundColor = "Gray"
        }
    }
}

function RevertProperForegroundColor {
    $Host.UI.RawUI.ForegroundColor = $Script:OriginalConsoleForegroundColor
}

function SetWriteHostAction ($DebugAction) {
    $Script:WriteHostDebugAction = $DebugAction
}

function SetWriteHostManipulateObjectAction ($ManipulateObject) {
    $Script:WriteHostManipulateObjectAction = $ManipulateObject
}

function Write-Verbose {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'In order to log Write-Verbose from Shared functions')]
    [CmdletBinding()]
    param(
        [Parameter(Position = 1, ValueFromPipeline)]
        [string]$Message
    )

    process {

        if ($null -ne $Script:WriteVerboseManipulateMessageAction) {
            $Message = & $Script:WriteVerboseManipulateMessageAction $Message
        }

        Microsoft.PowerShell.Utility\Write-Verbose $Message

        if ($null -ne $Script:WriteVerboseDebugAction) {
            & $Script:WriteVerboseDebugAction $Message
        }

        # $PSSenderInfo is set when in a remote context
        if ($PSSenderInfo -and
            $null -ne $Script:WriteRemoteVerboseDebugAction) {
            & $Script:WriteRemoteVerboseDebugAction $Message
        }
    }
}

function SetWriteVerboseAction ($DebugAction) {
    $Script:WriteVerboseDebugAction = $DebugAction
}

function SetWriteRemoteVerboseAction ($DebugAction) {
    $Script:WriteRemoteVerboseDebugAction = $DebugAction
}

function SetWriteVerboseManipulateMessageAction ($DebugAction) {
    $Script:WriteVerboseManipulateMessageAction = $DebugAction
}

function Get-NewLoggerInstance {
    [CmdletBinding()]
    param(
        [string]$LogDirectory = (Get-Location).Path,

        [ValidateNotNullOrEmpty()]
        [string]$LogName = "Script_Logging",

        [bool]$AppendDateTime = $true,

        [bool]$AppendDateTimeToFileName = $true,

        [int]$MaxFileSizeMB = 10,

        [int]$CheckSizeIntervalMinutes = 10,

        [int]$NumberOfLogsToKeep = 10
    )

    $fileName = if ($AppendDateTimeToFileName) { "{0}_{1}.txt" -f $LogName, ((Get-Date).ToString('yyyyMMddHHmmss')) } else { "$LogName.txt" }
    $fullFilePath = [System.IO.Path]::Combine($LogDirectory, $fileName)

    if (-not (Test-Path $LogDirectory)) {
        try {
            New-Item -ItemType Directory -Path $LogDirectory -ErrorAction Stop | Out-Null
        } catch {
            throw "Failed to create Log Directory: $LogDirectory. Inner Exception: $_"
        }
    }

    return [PSCustomObject]@{
        FullPath                 = $fullFilePath
        AppendDateTime           = $AppendDateTime
        MaxFileSizeMB            = $MaxFileSizeMB
        CheckSizeIntervalMinutes = $CheckSizeIntervalMinutes
        NumberOfLogsToKeep       = $NumberOfLogsToKeep
        BaseInstanceFileName     = $fileName.Replace(".txt", "")
        Instance                 = 1
        NextFileCheckTime        = ((Get-Date).AddMinutes($CheckSizeIntervalMinutes))
        PreventLogCleanup        = $false
        LoggerDisabled           = $false
    } | Write-LoggerInstance -Object "Starting Logger Instance $(Get-Date)"
}

function Write-LoggerInstance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object]$LoggerInstance,

        [Parameter(Mandatory = $true, Position = 1)]
        [object]$Object
    )
    process {
        if ($LoggerInstance.LoggerDisabled) { return }

        if ($LoggerInstance.AppendDateTime -and
            $Object.GetType().Name -eq "string") {
            $Object = "[$([System.DateTime]::Now)] : $Object"
        }

        # Doing WhatIf:$false to support -WhatIf in main scripts but still log the information
        $Object | Out-File $LoggerInstance.FullPath -Append -WhatIf:$false

        #Upkeep of the logger information
        if ($LoggerInstance.NextFileCheckTime -gt [System.DateTime]::Now) {
            return
        }

        #Set next update time to avoid issues so we can log things
        $LoggerInstance.NextFileCheckTime = ([System.DateTime]::Now).AddMinutes($LoggerInstance.CheckSizeIntervalMinutes)
        $item = Get-ChildItem $LoggerInstance.FullPath

        if (($item.Length / 1MB) -gt $LoggerInstance.MaxFileSizeMB) {
            $LoggerInstance | Write-LoggerInstance -Object "Max file size reached rolling over" | Out-Null
            $directory = [System.IO.Path]::GetDirectoryName($LoggerInstance.FullPath)
            $fileName = "$($LoggerInstance.BaseInstanceFileName)-$($LoggerInstance.Instance).txt"
            $LoggerInstance.Instance++
            $LoggerInstance.FullPath = [System.IO.Path]::Combine($directory, $fileName)

            $items = Get-ChildItem -Path ([System.IO.Path]::GetDirectoryName($LoggerInstance.FullPath)) -Filter "*$($LoggerInstance.BaseInstanceFileName)*"

            if ($items.Count -gt $LoggerInstance.NumberOfLogsToKeep) {
                $item = $items | Sort-Object LastWriteTime | Select-Object -First 1
                $LoggerInstance | Write-LoggerInstance "Removing Log File $($item.FullName)" | Out-Null
                $item | Remove-Item -Force
            }
        }
    }
    end {
        return $LoggerInstance
    }
}

function Invoke-LoggerInstanceCleanup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object]$LoggerInstance
    )
    process {
        if ($LoggerInstance.LoggerDisabled -or
            $LoggerInstance.PreventLogCleanup) {
            return
        }

        Get-ChildItem -Path ([System.IO.Path]::GetDirectoryName($LoggerInstance.FullPath)) -Filter "*$($LoggerInstance.BaseInstanceFileName)*" |
            Remove-Item -Force
    }
}


function WriteErrorInformationBase {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0],
        [ValidateSet("Write-Host", "Write-Verbose")]
        [string]$Cmdlet
    )

    if ($null -ne $CurrentError.OriginInfo) {
        & $Cmdlet "Error Origin Info: $($CurrentError.OriginInfo.ToString())"
    }

    & $Cmdlet "$($CurrentError.CategoryInfo.Activity) : $($CurrentError.ToString())"

    if ($null -ne $CurrentError.Exception -and
        $null -ne $CurrentError.Exception.StackTrace) {
        & $Cmdlet "Inner Exception: $($CurrentError.Exception.StackTrace)"
    } elseif ($null -ne $CurrentError.Exception) {
        & $Cmdlet "Inner Exception: $($CurrentError.Exception)"
    }

    if ($null -ne $CurrentError.InvocationInfo.PositionMessage) {
        & $Cmdlet "Position Message: $($CurrentError.InvocationInfo.PositionMessage)"
    }

    if ($null -ne $CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage) {
        & $Cmdlet "Remote Position Message: $($CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage)"
    }

    if ($null -ne $CurrentError.ScriptStackTrace) {
        & $Cmdlet "Script Stack: $($CurrentError.ScriptStackTrace)"
    }
}

function Write-VerboseErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Verbose"
}

function Write-HostErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Host"
}

function Invoke-CatchActions {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    $script:ErrorsExcluded += $CurrentError
    Write-Verbose "Error Excluded Count: $($Script:ErrorsExcluded.Count)"
    Write-Verbose "Error Count: $($Error.Count)"
    Write-VerboseErrorInformation $CurrentError
}

function Get-UnhandledErrors {
    [CmdletBinding()]
    param ()
    $index = 0
    return $Error |
        ForEach-Object {
            $currentError = $_
            $handledError = $Script:ErrorsExcluded |
                Where-Object { $_.Equals($currentError) }

                if ($null -eq $handledError) {
                    [PSCustomObject]@{
                        ErrorInformation = $currentError
                        Index            = $index
                    }
                }
                $index++
            }
}

function Get-HandledErrors {
    [CmdletBinding()]
    param ()
    $index = 0
    return $Error |
        ForEach-Object {
            $currentError = $_
            $handledError = $Script:ErrorsExcluded |
                Where-Object { $_.Equals($currentError) }

                if ($null -ne $handledError) {
                    [PSCustomObject]@{
                        ErrorInformation = $currentError
                        Index            = $index
                    }
                }
                $index++
            }
}

function Test-UnhandledErrorsOccurred {
    return $Error.Count -ne $Script:ErrorsExcluded.Count
}

function Invoke-ErrorCatchActionLoopFromIndex {
    [CmdletBinding()]
    param(
        [int]$StartIndex
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    Write-Verbose "Start Index: $StartIndex Error Count: $($Error.Count)"

    if ($StartIndex -ne $Error.Count) {
        # Write the errors out in reverse in the order that they came in.
        $index = $Error.Count - $StartIndex - 1
        do {
            Invoke-CatchActions $Error[$index]
            $index--
        } while ($index -ge 0)
    }
}

function Invoke-ErrorMonitoring {
    # Always clear out the errors
    # setup variable to monitor errors that occurred
    $Error.Clear()
    $Script:ErrorsExcluded = @()
}

function Invoke-WriteDebugErrorsThatOccurred {

    function WriteErrorInformation {
        [CmdletBinding()]
        param(
            [object]$CurrentError
        )
        Write-VerboseErrorInformation $CurrentError
        Write-Verbose "-----------------------------------`r`n`r`n"
    }

    if ($Error.Count -gt 0) {
        Write-Verbose "`r`n`r`nErrors that occurred that wasn't handled"

        Get-UnhandledErrors | ForEach-Object {
            Write-Verbose "Error Index: $($_.Index)"
            WriteErrorInformation $_.ErrorInformation
        }

        Write-Verbose "`r`n`r`nErrors that were handled"
        Get-HandledErrors | ForEach-Object {
            Write-Verbose "Error Index: $($_.Index)"
            WriteErrorInformation $_.ErrorInformation
        }
    } else {
        Write-Verbose "No errors occurred in the script."
    }
}

function Get-ExchangeInstallDirectory {
    [CmdletBinding()]
    param()

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    $installDirectory = [string]::Empty
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup') {
        Write-Verbose "Detected v14"
        $installDirectory = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath
    } elseif (Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup') {
        Write-Verbose "Detected v15"
        $installDirectory = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath
    } else {
        Write-Host "Something went wrong trying to find Exchange Install path on this server: $env:COMPUTERNAME"
    }

    Write-Verbose "Returning: $installDirectory"

    return $installDirectory
}


function Get-ItemsSize {
    param(
        [Parameter(Mandatory = $true)][array]$FilePaths
    )
    Write-Verbose("Calling: Get-ItemsSize")
    $totalSize = 0
    $hashSizes = @{}
    foreach ($file in $FilePaths) {
        if (Test-Path $file) {
            $totalSize += ($fileSize = (Get-Item $file).Length)
            Write-Verbose("File: {0} | Size: {1} MB" -f $file, ($fileSizeMB = $fileSize / 1MB))
            $hashSizes.Add($file, ("{0}" -f $fileSizeMB))
        } else {
            Write-Verbose("File no longer exists: {0}" -f $file)
        }
    }
    Set-Variable -Name ItemSizesHashed -Value $hashSizes -Scope Script
    Write-Verbose("Returning: {0}" -f $totalSize)
    return $totalSize
}

function Compress-Folder {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Position = 1)][string]$Folder,
        [Parameter(Position = 2)][bool]$IncludeMonthDay = $false,
        [Parameter(Position = 3)][bool]$IncludeDisplayZipping = $true,
        [Parameter(Position = 4)][bool]$ReturnCompressedLocation = $false
    )

    $Folder = $Folder.TrimEnd("\")
    $compressedLocation = [string]::Empty
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    Write-Verbose "Passed - [string]Folder: $Folder | [bool]IncludeDisplayZipping: $IncludeDisplayZipping | [bool]ReturnCompressedLocation: $ReturnCompressedLocation"

    if (-not (Test-Path $Folder)) {
        Write-Host "Failed to find the folder $Folder"
        return $null
    }

    $successful = ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.Location -like "*System.IO.Compression.Filesystem*" }).Count -ge 1
    Write-Verbose "Found IO Compression loaded: $successful"

    if ($successful -eq $false) {
        # Try to load the IO Compression
        try {
            Add-Type -AssemblyName System.IO.Compression.Filesystem -ErrorAction Stop
            Write-Verbose "Loaded .NET Compression Assembly."
        } catch {
            Write-Host "Failed to load .NET Compression assembly. Unable to compress up the data."
            return $null
        }
    }

    if ($IncludeMonthDay) {
        $zipFolderNoEXT = "{0}-{1}" -f $Folder, (Get-Date -Format Md)
    } else {
        $zipFolderNoEXT = $Folder
    }
    Write-Verbose "[string]zipFolderNoEXT: $zipFolderNoEXT"
    $zipFolder = "{0}.zip" -f $zipFolderNoEXT
    [int]$i = 1
    while (Test-Path $zipFolder) {
        $zipFolder = "{0}-{1}.zip" -f $zipFolderNoEXT, $i
        $i++
    }
    Write-Verbose "Using Zip Folder Path: $zipFolder"

    if ($IncludeDisplayZipping) {
        Write-Host "Compressing Folder $Folder"
    }
    $sizeBytesBefore = 0
    Get-ChildItem $Folder -Recurse |
        Where-Object { -not ($_.Mode.StartsWith("d-")) } |
        ForEach-Object { $sizeBytesBefore += $_.Length }

    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    [System.IO.Compression.ZipFile]::CreateFromDirectory($Folder, $zipFolder)
    $timer.Stop()
    $sizeBytesAfter = (Get-Item $zipFolder).Length
    Write-Verbose ("Compressing directory size of {0} MB down to the size of {1} MB took {2} seconds." -f ($sizeBytesBefore / 1MB), ($sizeBytesAfter / 1MB), $timer.Elapsed.TotalSeconds)

    if ((Test-Path -Path $zipFolder)) {
        Write-Verbose "Compress successful, removing folder."
        Remove-Item $Folder -Force -Recurse
    }

    if ($ReturnCompressedLocation) {
        $compressedLocation = $zipFolder
    }

    Write-Verbose "Returning: $compressedLocation"
    return $compressedLocation
}
function Invoke-ZipFolder {
    param(
        [string]$Folder,
        [bool]$ZipItAll,
        [bool]$AddCompressedSize = $true
    )

    if ($ZipItAll) {
        Write-Verbose("Disabling Logger before zipping up the directory")
        $Script:Logger.LoggerDisabled = $true
        Compress-Folder -Folder $Folder -IncludeMonthDay $true
    } else {
        $compressedLocation = Compress-Folder -Folder $Folder -ReturnCompressedLocation $AddCompressedSize
        if ($AddCompressedSize -and ($compressedLocation -ne [string]::Empty)) {
            $Script:TotalBytesSizeCompressed += ($size = Get-ItemsSize -FilePaths $compressedLocation)
            $Script:FreeSpaceMinusCopiedAndCompressedGB -= ($size / 1GB)
            Write-Verbose("Current Sizes after compression: [double]TotalBytesSizeCompressed: {0} | [double]FreeSpaceMinusCopiedAndCompressedGB: {1}" -f $Script:TotalBytesSizeCompressed,
                $Script:FreeSpaceMinusCopiedAndCompressedGB)
        }
    }
}

# Small set of functions that are used to help override the Write-Host and Write-Verbose functions
function Get-ManipulateWriteHostValue {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Position = 1, ValueFromPipeline)]
        [string]$Message
    )

    process {
        return "[$env:COMPUTERNAME] : $Message"
    }
}

function Get-ManipulateWriteVerboseValue {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Position = 1, ValueFromPipeline)]
        [string]$Message
    )

    process {
        return "[$env:COMPUTERNAME - Script Debug] : $Message"
    }
}

#Calls the $Script:Logger object to write the data to file only.
function Write-DebugLog($message) {
    if ($null -ne $message -and
        ![string]::IsNullOrEmpty($message) -and
        $null -ne $Script:Logger) {
        $Script:Logger = $Script:Logger | Write-LoggerInstance $message
    }
}


function Get-FreeSpace {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWMICmdlet', '', Justification = 'Different types returned')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][ValidateScript( { $_.ToString().EndsWith("\") })][string]$FilePath
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    Write-Verbose "Passed: [string]FilePath: $FilePath"

    $drivesList = Get-CimInstance Win32_Volume -Filter "DriveType = 3"
    $testPath = $FilePath
    $freeSpaceSize = -1
    while ($true) {
        if ($testPath -eq [string]::Empty) {
            Write-Host "Unable to fine a drive that matches the file path: $FilePath"
            return $freeSpaceSize
        }
        Write-Verbose "Trying to find path that matches path: $testPath"
        foreach ($drive in $drivesList) {
            if ($drive.Name -eq $testPath) {
                Write-Verbose "Found a match"
                $freeSpaceSize = $drive.FreeSpace / 1GB
                Write-Verbose "Have $freeSpaceSize`GB of Free Space"
                return $freeSpaceSize
            }
            Write-Verbose "Drive name: '$($drive.Name)' didn't match"
        }

        $itemTarget = [string]::Empty
        if ((Test-Path $testPath)) {
            $item = Get-Item $testPath
            if ($item.Target -like "Volume{*}\") {
                Write-Verbose "File Path appears to be a mount point target: $($item.Target)"
                $itemTarget = $item.Target
            } else {
                Write-Verbose "Path didn't appear to be a mount point target"
            }
        } else {
            Write-Verbose "Path isn't a true path yet."
        }

        if ($itemTarget -ne [string]::Empty) {
            foreach ($drive in $drivesList) {
                if ($drive.DeviceID.Contains($itemTarget)) {
                    $freeSpaceSize = $drive.FreeSpace / 1GB
                    Write-Verbose "Have $freeSpaceSize`GB of Free Space"
                    return $freeSpaceSize
                }
                Write-Verbose "DeviceID didn't appear to match: $($drive.DeviceID)"
            }
            if ($freeSpaceSize -eq -1) {
                Write-Host "Unable to fine a drive that matches the file path: $FilePath"
                Write-Host "This shouldn't have happened."
                return $freeSpaceSize
            }
        }
        $testPath = $testPath.Substring(0, $testPath.LastIndexOf("\", $testPath.Length - 2) + 1)
    }
}


function Test-CommandExists {
    param(
        [string]$command
    )

    try {
        if (Get-Command $command -ErrorAction Stop) {
            return $true
        }
    } catch {
        Invoke-CatchActions
        return $false
    }
}
function Get-IISLogDirectory {
    Write-Verbose("Function Enter: Get-IISLogDirectory")

    function Get-IISDirectoryFromGetWebSite {
        Write-Verbose("Get-WebSite command exists")
        return Get-Website |
            ForEach-Object {
                $logFile = "$($_.LogFile.Directory)\W3SVC$($_.id)".Replace("%SystemDrive%", $env:SystemDrive)
                Write-Verbose("Found Directory: $logFile")
                return $logFile
            }
    }

    if ((Test-CommandExists -command "Get-WebSite")) {
        [array]$iisLogDirectory = Get-IISDirectoryFromGetWebSite
    } else {
        #May need to load the module
        try {
            Write-Verbose("Going to attempt to load the WebAdministration Module")
            Import-Module WebAdministration -ErrorAction Stop
            Write-Verbose("Successful loading the module")

            if ((Test-CommandExists -command "Get-WebSite")) {
                [array]$iisLogDirectory = Get-IISDirectoryFromGetWebSite
            }
        } catch {
            Invoke-CatchActions
            [array]$iisLogDirectory = "C:\inetPub\logs\LogFiles\" #Default location for IIS Logs
            Write-Verbose("Get-WebSite command doesn't exists. Set IISLogDirectory to: {0}" -f $iisLogDirectory)
        }
    }

    Write-Verbose("Function Exit: Get-IISLogDirectory")
    return $iisLogDirectory
}

function NewTaskAction {
    [CmdletBinding()]
    param(
        [string]$FunctionName,
        [object]$Parameters
    )
    return [PSCustomObject]@{
        FunctionName = $FunctionName
        Parameters   = $Parameters
    }
}

function NewLogCopyParameters {
    param(
        [string]$LogPath,
        [string]$CopyToThisLocation
    )
    return @{
        LogPath            = $LogPath
        CopyToThisLocation = [System.IO.Path]::Combine($Script:RootCopyToDirectory, $CopyToThisLocation)
    }
}

function NewLogCopyBasedOffTimeParameters {
    param(
        [string]$LogPath,
        [string]$CopyToThisLocation,
        [bool]$IncludeSubDirectory
    )
    return (NewLogCopyParameters $LogPath $CopyToThisLocation) + @{
        IncludeSubDirectory = $IncludeSubDirectory
    }
}

function GetTaskActionToString {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [object]$TaskAction
    )
    $params = $TaskAction.Parameters
    $line = "$($TaskAction.FunctionName)"

    if ($null -ne $params) {
        $line += " LogPath: '$($params.LogPath)' CopyToThisLocation: '$($params.CopyToThisLocation)'"
    }
    return $line
}

function Add-TaskAction {
    param(
        [string]$FunctionName
    )
    $Script:taskActionList.Add((NewTaskAction $FunctionName))
}

function Add-LogCopyBasedOffTimeTaskAction {
    param(
        [string]$LogPath,
        [string]$CopyToThisLocation,
        [bool]$IncludeSubDirectory = $true
    )
    $timeCopyParams = @{
        LogPath             = $LogPath
        CopyToThisLocation  = $CopyToThisLocation
        IncludeSubDirectory = $IncludeSubDirectory
    }
    $params = @{
        FunctionName = "Copy-LogsBasedOnTime"
        Parameters   = (NewLogCopyBasedOffTimeParameters @timeCopyParams)
    }
    $Script:taskActionList.Add((NewTaskAction @params))
}

function Add-LogCopyFullTaskAction {
    param (
        [string]$LogPath,
        [string]$CopyToThisLocation
    )
    $params = @{
        FunctionName = "Copy-FullLogFullPathRecurse"
        Parameters   = (NewLogCopyParameters $LogPath $CopyToThisLocation)
    }
    $Script:taskActionList.Add((NewTaskAction @params))
}

function Add-DefaultLogCopyTaskAction {
    param(
        [string]$LogPath,
        [string]$CopyToThisLocation
    )
    if ($PassedInfo.CollectAllLogsBasedOnLogAge) {
        Add-LogCopyBasedOffTimeTaskAction $LogPath $CopyToThisLocation
    } else {
        Add-LogCopyFullTaskAction $LogPath $CopyToThisLocation
    }
}


function Get-StringDataForNotEnoughFreeSpaceFile {
    param(
        [Parameter(Mandatory = $true)][Hashtable]$FileSizes
    )
    Write-Verbose("Calling: Get-StringDataForNotEnoughFreeSpaceFile")
    $reader = [string]::Empty
    $totalSizeMB = 0
    foreach ($key in $FileSizes.Keys) {
        $reader += ("File: {0} | Size: {1} MB`r`n" -f $key, ($keyValue = $FileSizes[$key]).ToString())
        $totalSizeMB += $keyValue
    }
    $reader += ("`r`nTotal Size Attempted To Copy Over: {0} MB`r`nCurrent Available Free Space: {1} GB" -f $totalSizeMB, $Script:CurrentFreeSpaceGB)
    return $reader
}

function Test-FreeSpace {
    param(
        [Parameter(Mandatory = $false)][array]$FilePaths
    )
    Write-Verbose("Calling: Test-FreeSpace")

    if ($null -eq $FilePaths -or
        $FilePaths.Count -eq 0) {
        Write-Verbose("Null FilePaths provided returning true.")
        return $true
    }

    $passed = $true
    $currentSizeCopy = Get-ItemsSize -FilePaths $FilePaths
    #It is better to be safe than sorry, checking against probably a value way higher than needed.
    if (($Script:FreeSpaceMinusCopiedAndCompressedGB - ($currentSizeCopy / 1GB)) -lt $Script:AdditionalFreeSpaceCushionGB) {
        Write-Verbose("Estimated free space is getting low, going to recalculate.")
        Write-Verbose("Current values: [double]FreeSpaceMinusCopiedAndCompressedGB: {0} | [double]currentSizeCopy: {1} | [double]AdditionalFreeSpaceCushionGB: {2} | [double]CurrentFreeSpaceGB: {3}" -f $Script:FreeSpaceMinusCopiedAndCompressedGB,
            ($currentSizeCopy / 1GB),
            $Script:AdditionalFreeSpaceCushionGB,
            $Script:CurrentFreeSpaceGB)
        $freeSpace = Get-FreeSpace -FilePath ("{0}\" -f $Script:RootCopyToDirectory)
        Write-Verbose("True current free space: {0}" -f $freeSpace)

        if ($freeSpace -lt ($Script:CurrentFreeSpaceGB - .5)) {
            #If we off by .5GB, we need to know about this and look at the data to determine if we might have some logical errors. It is possible that the disk is that active, but that wouldn't be good either for this script.
            Write-Verbose("CRIT: Disk Space logic is off. CurrentFreeSpaceGB: {0} | ActualFreeSpace: {1}" -f $Script:CurrentFreeSpaceGB, $freeSpace)
        }

        $Script:CurrentFreeSpaceGB = $freeSpace
        $Script:FreeSpaceMinusCopiedAndCompressedGB = $freeSpace
        $passed = $freeSpace -gt ($addSize = $Script:AdditionalFreeSpaceCushionGB + ($currentSizeCopy / 1GB))

        if (!($passed)) {
            Write-Host "Free space on the drive has appear to be used up past recommended thresholds. Going to stop this execution of the script. If you feel this is an Error, please notify ExToolsFeedback@microsoft.com" -ForegroundColor "Red"
            Write-Host "FilePath: $($Script:RootCopyToDirectory) | FreeSpace: $freeSpace | Looking for: $(($freeSpace + $addSize))" -ForegroundColor "Red"
            return $passed
        }
    }

    $Script:TotalBytesSizeCopied += $currentSizeCopy
    $Script:FreeSpaceMinusCopiedAndCompressedGB = $Script:FreeSpaceMinusCopiedAndCompressedGB - ($currentSizeCopy / 1GB)

    Write-Verbose("Current values [double]FreeSpaceMinusCopiedAndCompressedGB: {0} | [double]TotalBytesSizeCopied: {1}" -f $Script:FreeSpaceMinusCopiedAndCompressedGB, $Script:TotalBytesSizeCopied)
    Write-Verbose("Returning: {0}" -f $passed)
    return $passed
}
function Copy-BulkItems {
    param(
        [string]$CopyToLocation,
        [Array]$ItemsToCopyLocation
    )

    New-Item -ItemType Directory -Path $CopyToLocation -Force | Out-Null

    if (Test-FreeSpace -FilePaths $ItemsToCopyLocation) {
        foreach ($item in $ItemsToCopyLocation) {
            Copy-Item -Path $item -Destination $CopyToLocation -ErrorAction SilentlyContinue
        }
    } else {
        Write-Host "Not enough free space to copy over this data set."
        New-Item -Path ("{0}\NotEnoughFreeSpace.txt" -f $CopyToLocation) -ItemType File -Value (Get-StringDataForNotEnoughFreeSpaceFile -FileSizes $Script:ItemSizesHashed) | Out-Null
    }
}

function Copy-FullLogFullPathRecurse {
    param(
        [Parameter(Mandatory = $true)][string]$LogPath,
        [Parameter(Mandatory = $true)][string]$CopyToThisLocation
    )
    Write-Verbose("Function Enter: Copy-FullLogFullPathRecurse")
    Write-Verbose("Passed: [string]LogPath: {0} | [string]CopyToThisLocation: {1}" -f $LogPath, $CopyToThisLocation)
    New-Item -ItemType Directory -Path $CopyToThisLocation -Force | Out-Null
    if (Test-Path $LogPath) {
        $childItems = Get-ChildItem $LogPath -Recurse
        $items = @()
        foreach ($childItem in $childItems) {
            if (!($childItem.Mode.StartsWith("d-"))) {
                $items += $childItem.VersionInfo.FileName
            }
        }

        if ($null -ne $items -and
            $items.Count -gt 0) {
            if (Test-FreeSpace -FilePaths $items) {
                Copy-Item $LogPath\* $CopyToThisLocation -Recurse -ErrorAction SilentlyContinue
                Invoke-ZipFolder $CopyToThisLocation
            } else {
                Write-Verbose("Not going to copy over this set of data due to size restrictions.")
                New-Item -Path ("{0}\NotEnoughFreeSpace.txt" -f $CopyToThisLocation) -ItemType File -Value (Get-StringDataForNotEnoughFreeSpaceFile -FileSizes $Script:ItemSizesHashed) | Out-Null
            }
        } else {
            Write-Host "No data at path '$LogPath'. Unable to copy this data."
            New-Item -Path ("{0}\NoDataDetected.txt" -f $CopyToThisLocation) -ItemType File -Value $LogPath | Out-Null
        }
    } else {
        Write-Host "No Folder at $LogPath. Unable to copy this data."
        New-Item -Path ("{0}\NoFolderDetected.txt" -f $CopyToThisLocation) -ItemType File -Value $LogPath | Out-Null
    }
    Write-Verbose("Function Exit: Copy-FullLogFullPathRecurse")
}

<#
    Copy Log Directory Based Off Time.
    The IncludeSubDirectory bool set to false should only be use if we don't want to include sub directories
    Otherwise, in each sub directory try to collect logs based off the TimeSpan.
    If there is a directory that doesn't contain logs within the TimeSpan,
    Collect the latest log or provide there is no logs in the directory
#>
function Copy-LogsBasedOnTime {
    param(
        [Parameter(Mandatory = $true)][string]$LogPath,
        [Parameter(Mandatory = $true)][string]$CopyToThisLocation,
        [Parameter(Mandatory = $true)][bool]$IncludeSubDirectory
    )
    begin {
        function NoFilesInLocation {
            param(
                [string]$Value = "No data in the location"
            )
            $line = "It doesn't look like you have any data in this location $LogPath."

            if (-not ($IncludeSubDirectory)) {
                Write-Host $line -ForegroundColor "Yellow"
            } else {
                Write-Verbose $line
            }

            $params = @{
                Path     = "$CopyToThisLocation\NoFilesDetected.txt"
                ItemType = "File"
                Value    = $( "Location: $LogPath`r`n$Value" )
            }
            New-Item @params | Out-Null
        }

        function CopyItemsFromDirectory {
            param(
                [object]$AllItems,
                [string]$CopyToLocation
            )

            if ($null -eq $AllItems) {
                Write-Verbose "No items were found in the directory."
                NoFilesInLocation
            } else {
                $timeRangeFiles = $AllItems | Where-Object { $_.LastWriteTime -ge $copyFromDate -and $_.LastWriteTime -le $copyToDate }

                if ($null -eq $timeRangeFiles) {
                    Write-Verbose "no files found in the range. Getting the last file."
                    Copy-BulkItems -CopyToLocation $CopyToLocation -ItemsToCopyLocation $AllItems[0].FullName
                } else {
                    Write-Verbose "Found files within the time range."
                    $timeRangeFiles | ForEach-Object { Write-Verbose "$($_.FullName)" }
                    $copyItemPaths = $timeRangeFiles | ForEach-Object { $_.FullName }
                    Copy-BulkItems -CopyToLocation $CopyToLocation -ItemsToCopyLocation $copyItemPaths
                }
                Invoke-ZipFolder -Folder $CopyToLocation
            }
        }

        Write-Verbose "Function Enter: $($MyInvocation.MyCommand)"
        Write-Verbose "LogPath: '$LogPath' | CopyToThisLocation: '$CopyToThisLocation'"
        New-Item -ItemType Directory -Path $CopyToThisLocation -Force | Out-Null
        $copyFromDate = [DateTime]::Now - $PassedInfo.TimeSpan
        $copyToDate = [DateTime]::Now - $PassedInfo.EndTimeSpan
        Write-Verbose "Copy From Date: $copyFromDate"
        Write-Verbose "Copy To Date: $copyToDate"
    }
    process {

        # need to have the return in process
        if (-not (Test-Path $LogPath)) {
            # If the directory isn't there, provide that
            Write-Verbose "$LogPath doesn't exist"
            NoFilesInLocation "Path doesn't exist"
            return
        }

        if ($IncludeSubDirectory) {
            $getChildItem = Get-ChildItem -Path $LogPath -Recurse
            [array]$directories = Get-Item -Path $LogPath
            $directories += @($getChildItem |
                    Where-Object {
                        $_.Mode -like "d*"
                    })

            # Map and find all the items per directory
            foreach ($directory in $directories) {

                Write-Verbose "Working on finding items for directory $($directory.FullName)"
                if ($directory.FullName -eq $LogPath) {
                    $newCopyToThisLocation = $CopyToThisLocation
                } else {
                    $newCopyToThisLocation = "$CopyToThisLocation\$($directory.Name)"
                    New-Item -ItemType Directory -Path $newCopyToThisLocation -Force | Out-Null
                }
                # all the items that match this directory. Don't need to worry about directories because DirectoryName doesn't exist there.
                $items = $getChildItem | Where-Object { $_.DirectoryName -eq $directory.FullName } | Sort-Object LastWriteTime -Descending
                CopyItemsFromDirectory -AllItems $items -CopyToLocation $newCopyToThisLocation
            }
        } else {
            $getChildItem = Get-ChildItem -Path $LogPath |
                Sort-Object LastWriteTime -Descending |
                Where-Object { $_.Mode -notlike "d*" }

            CopyItemsFromDirectory -AllItems $getChildItem -CopyToLocation $CopyToThisLocation
        }
    }
    end {
        Write-Verbose("Function Exit: $($MyInvocation.MyCommand)")
    }
}

function CopyLogmanData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$LogmanObject
    )

    $copyTo = "$Script:RootCopyToDirectory\$($LogmanObject.LogmanName)_Data"
    New-Item -ItemType Directory -Path $copyTo -Force | Out-Null
    $directory = $LogmanObject.RootPath
    $filterDate = $LogmanObject.StartDate
    $copyFromDate = [DateTime]::Now - $PassedInfo.TimeSpan
    $copyToDate = [DateTime]::Now - $PassedInfo.EndTimeSpan
    Write-Verbose "Copy From Date: $filterDate"
    Write-Verbose "Copy To Date: $filterToDate"

    if ([DateTime]$filterDate -lt [DateTime]$copyFromDate) {
        $filterDate = $copyFromDate
        Write-Verbose "Updating Copy From Date: $filterDate"
    }

    if ([DateTime]$filterToDate -lt [DateTime]$copyToDate) {
        $filterToDate = $copyToDate
        Write-Verbose "Updating Copy to Date: $filterToDate"
    }

    if ((Test-Path $directory)) {

        $childItems = Get-ChildItem $directory -Recurse |
            Where-Object { $_.Name -like "*$($LogmanObject.Extension)" }

        if ($null -ne $childItems) {
            $items = $childItems |
                Where-Object { $_.CreationTime -ge $filterDate -and $_.CreationTime -le $filterToDate } |
                ForEach-Object { $_.VersionInfo.FileName }

            if ($null -ne $items) {
                Copy-BulkItems -CopyToLocation $copyTo -ItemsToCopyLocation $items
                Invoke-ZipFolder -Folder $copyTo
                return
            } else {
                Write-Host "Failed to find any files in the directory: $directory that was greater than or equal to this time: $filterDate and lower than $filterToDate" -ForegroundColor "Yellow"
                $filterDate = ($childItems |
                        Sort-Object CreationTime -Descending |
                        Select-Object -First 1).CreationTime.AddDays(-1)
                Write-Verbose "Updated filter time to $filterDate"
                $items = $childItems |
                    Where-Object { $_.CreationTime -ge $filterDate -and $_.CreationTime -le $filterToDate } |
                    ForEach-Object { $_.VersionInfo.FileName }

                if ($null -ne $items) {
                    Copy-BulkItems -CopyToLocation $copyTo -ItemsToCopyLocation $items
                    Invoke-ZipFolder -Folder $copyTo
                    return
                }
                Write-Verbose "Something went really wrong..."
            }
        }
        Write-Host "Failed to find any files in the directory $directory" -ForegroundColor "Yellow"
        New-Item -Path "$copyTo\NoFiles.txt" -Value $directory | Out-Null
    } else {
        Write-Host "Doesn't look like this Directory is valid: $directory" -ForegroundColor "Yellow"
        New-Item -Path "$copyTo\NotValidDirectory.txt" -Value $directory | Out-Null
    }
}

function GetLogmanObject {
    [CmdletBinding()]
    param(
        [string]$LogmanName
    )

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $status = "Stopped"
        $rootPath = [string]::Empty
        $extension = ".blg"
        $startDate = [DateTime]::MinValue
        $foundLogman = $false
    }
    process {
        try {
            $dataCollectorSetList = New-Object -ComObject Pla.DataCollectorSetCollection
            $dataCollectorSetList.GetDataCollectorSets($null, $null)
            $existingLogmanDataCollectorSetList = $dataCollectorSetList | Where-Object { $_.Name -eq $LogmanName }

            if ($null -eq $existingLogmanDataCollectorSetList) { return }

            if ($existingLogmanDataCollectorSetList.Status -eq 1) {
                $status = "Running"
            }

            $rootPath = $existingLogmanDataCollectorSetList.RootPath
            $outputLocation = $existingLogmanDataCollectorSetList.DataCollectors._NewEnum.OutputLocation
            Write-Verbose "Output Location: $outputLocation"
            $extension = $outputLocation.Substring($outputLocation.LastIndexOf("."))
            $startDate = $existingLogmanDataCollectorSetList.Schedules._NewEnum.StartDate
            Write-Verbose "Status: $status RootPath: $rootPath Extension: $extension StartDate: $startDate"
            $foundLogman = $true
        } catch {
            Write-Verbose "Failed to get the Logman information. Exception $_"
        }

        finally {
            if ($null -ne $dataCollectorSetList) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($dataCollectorSetList) | Out-Null
                $dataCollectorSetList = $null
                $existingLogmanDataCollectorSetList = $null
            }
        }
    }
    end {
        return [PSCustomObject]@{
            LogmanName  = $LogmanName
            Status      = $status
            RootPath    = $rootPath
            Extension   = $extension
            StartDate   = $startDate
            FoundLogman = $foundLogman
        }
    }
}

function GetLogmanData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogmanName
    )
    $logmanObject = GetLogmanObject -LogmanName $LogmanName

    if ($logmanObject.FoundLogman) {
        if ($logmanObject.Status -eq "Running") {
            Write-Host "$LogmanName is running. Going to stop to prevent corruption for collection...."
            logman stop $LogmanName | Write-Verbose

            if ($LASTEXITCODE) {
                Write-Host "Failed to stop $LogmanName. $LastExitCode" -ForegroundColor "Red"
            }

            CopyLogmanData -LogmanObject $logmanObject
            Write-Host "Going to start $LogmanName again for you...."
            logman start $LogmanName | Write-Verbose

            if ($LASTEXITCODE) {
                Write-Host "Failed to start $LogmanName. $LastExitCode" -ForegroundColor "Red"
            }
        } else {
            Write-Host "$LogmanName isn't running, therefore not going to stop it prior to collection."
            CopyLogmanData -LogmanObject $logmanObject
        }
        Write-Host "Done copying $LogmanName"
    } else {
        Write-Host "Can't find Logman '$LogmanName'. Moving on..."
    }
}

function Save-LogmanExMonData {
    GetLogmanData -LogmanName $PassedInfo.ExMonLogmanName
}

function Save-LogmanExPerfWizData {
    $PassedInfo.ExPerfWizLogmanName |
        ForEach-Object {
            GetLogmanData -LogmanName $_
        }
}


function Save-DataToFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][object]$DataIn,
        [Parameter(Mandatory = $true)][string]$SaveToLocation,
        [Parameter(Mandatory = $false)][bool]$FormatList = $true,
        [Parameter(Mandatory = $false)][bool]$SaveTextFile = $true,
        [Parameter(Mandatory = $false)][bool]$SaveXMLFile = $true
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    Write-Verbose "Passed: [string]SaveToLocation: $SaveToLocation | [bool]FormatList: $FormatList | [bool]SaveTextFile: $SaveTextFile | [bool]SaveXMLFile: $SaveXMLFile"
    $xmlSaveLocation = "{0}.xml" -f $SaveToLocation
    $txtSaveLocation = "{0}.txt" -f $SaveToLocation

    if ($DataIn -ne [string]::Empty -and
        $null -ne $DataIn) {
        if ($SaveXMLFile) {
            $DataIn | Export-Clixml $xmlSaveLocation -Encoding UTF8
        }
        if ($SaveTextFile) {
            if ($FormatList) {
                $DataIn | Format-List * | Out-File $txtSaveLocation
            } else {
                $DataIn | Format-Table -AutoSize | Out-File $txtSaveLocation
            }
        }
    } else {
        Write-Verbose("DataIn was an empty string. Not going to save anything.")
    }
    Write-Verbose ("Returning from Save-DataToFile")
}

function Add-ServerNameToFileName {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath
    )
    Write-Verbose("Calling: Add-ServerNameToFileName")
    Write-Verbose("Passed: [string]FilePath: {0}" -f $FilePath)
    $fileName = "{0}_{1}" -f $env:COMPUTERNAME, ($name = $FilePath.Substring($FilePath.LastIndexOf("\") + 1))
    $filePathWithServerName = $FilePath.Replace($name, $fileName)
    Write-Verbose("Returned: {0}" -f $filePathWithServerName)
    return $filePathWithServerName
}
function Save-DataInfoToFile {
    param(
        [Parameter(Mandatory = $false)][object]$DataIn,
        [Parameter(Mandatory = $true)][string]$SaveToLocation,
        [Parameter(Mandatory = $false)][bool]$FormatList = $true,
        [Parameter(Mandatory = $false)][bool]$SaveTextFile = $true,
        [Parameter(Mandatory = $false)][bool]$SaveXMLFile = $true,
        [Parameter(Mandatory = $false)][bool]$AddServerName = $true
    )
    [System.Diagnostics.Stopwatch]$timer = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Verbose "Function Enter: Save-DataInfoToFile"

    if ($AddServerName) {
        $SaveToLocation = Add-ServerNameToFileName $SaveToLocation
    }

    Save-DataToFile -DataIn $DataIn -SaveToLocation $SaveToLocation -FormatList $FormatList -SaveTextFile $SaveTextFile -SaveXMLFile $SaveXMLFile
    $timer.Stop()
    Write-Verbose("Took {0} seconds to save out the data." -f $timer.Elapsed.TotalSeconds)
}


function Save-RegistryHive {
    [CmdletBinding()]
    param(
        [string]$RegistryPath,
        [string]$SaveName,
        [string]$SaveToPath,
        [switch]$UseGetChildItem
    )
    Write-Verbose "Function Enter: $($MyInvocation.MyCommand)"

    try {
        if ($UseGetChildItem) {
            $results = Get-ChildItem -Path $RegistryPath -Recurse -ErrorAction Stop
        } else {
            $results = Get-Item -Path $RegistryPath -ErrorAction Stop
        }
        Write-Verbose "Successfully got registry hive information for: $RegistryPath"
        Save-DataInfoToFile -DataIn $results -SaveToLocation "$SaveToPath\$SaveName" -FormatList $false
    } catch {
        Write-Verbose "Failed to get registry hive for: $RegistryPath"
        Invoke-CatchActions
    }

    $updatedRegistryPath = $RegistryPath.Replace("HKLM:", "HKEY_LOCAL_MACHINE\")
    $baseSaveName = Add-ServerNameToFileName "$SaveToPath\$SaveName"
    try {
        reg export $updatedRegistryPath "$baseSaveName.reg" | Out-Null

        if ($LASTEXITCODE) {
            throw "Failed to export the registry hive for: $updatedRegistryPath"
        }
        reg save $updatedRegistryPath "$baseSaveName.hiv" | Out-Null

        if ($LASTEXITCODE) {
            throw "Failed to save the registry hive for: $updatedRegistryPath"
        }
        "To read the registry hive. Run 'reg load HKLM\TempHive $SaveName.hiv'. Then Open your regedit then go to HKLM:\TempHive to view the data." |
            Out-File -FilePath "$baseSaveName`_HowToRead.txt"
    } catch {
        Write-Verbose "failed to export/save the registry hive for: $updatedRegistryPath"
        Invoke-CatchActions
    }
}

function Get-ClusterNodeFileVersions {
    [CmdletBinding()]
    param(
        [string]$ClusterDirectory = "C:\Windows\Cluster"
    )

    $fileHashes = @{}

    Get-ChildItem $ClusterDirectory |
        Where-Object {
            $_.Name.EndsWith(".dll") -or
            $_.Name.EndsWith(".exe")
        } |
        ForEach-Object {
            $item = [PSCustomObject]@{
                FileName        = $_.Name
                FileMajorPart   = $_.VersionInfo.FileMajorPart
                FileMinorPart   = $_.VersionInfo.FileMinorPart
                FileBuildPart   = $_.VersionInfo.FileBuildPart
                FilePrivatePart = $_.VersionInfo.FilePrivatePart
                ProductVersion  = $_.VersionInfo.ProductVersion
                LastWriteTime   = $_.LastWriteTimeUtc
            }
            $fileHashes.Add($_.Name, $item)
        }

    return [PSCustomObject]@{
        ComputerName = $env:COMPUTERNAME
        Files        = $fileHashes
    }
}
#Save out the failover cluster information for the local node, besides the event logs.
function Save-FailoverClusterInformation {
    Write-Verbose("Function Enter: Save-FailoverClusterInformation")
    $copyTo = "$Script:RootCopyToDirectory\Cluster_Information"
    New-Item -ItemType Directory -Path $copyTo -Force | Out-Null

    try {
        Save-DataInfoToFile -DataIn (Get-Cluster -ErrorAction Stop) -SaveToLocation "$copyTo\GetCluster"
    } catch {
        Write-Verbose "Failed to run Get-Cluster"
        Invoke-CatchActions
    }

    try {
        Save-DataInfoToFile -DataIn (Get-ClusterGroup -ErrorAction Stop) -SaveToLocation "$copyTo\GetClusterGroup"
    } catch {
        Write-Verbose "Failed to run Get-ClusterGroup"
        Invoke-CatchActions
    }

    try {
        Save-DataInfoToFile -DataIn (Get-ClusterNode -ErrorAction Stop) -SaveToLocation "$copyTo\GetClusterNode"
    } catch {
        Write-Verbose "Failed to run Get-ClusterNode"
        Invoke-CatchActions
    }

    try {
        Save-DataInfoToFile -DataIn (Get-ClusterNetwork -ErrorAction Stop) -SaveToLocation "$copyTo\GetClusterNetwork"
    } catch {
        Write-Verbose "Failed to run Get-ClusterNetwork"
        Invoke-CatchActions
    }

    try {
        Save-DataInfoToFile -DataIn (Get-ClusterNetworkInterface -ErrorAction Stop) -SaveToLocation "$copyTo\GetClusterNetworkInterface"
    } catch {
        Write-Verbose "Failed to run Get-ClusterNetworkInterface"
        Invoke-CatchActions
    }

    try {
        Get-ClusterLog -Node $env:ComputerName -Destination $copyTo -ErrorAction Stop | Out-Null
    } catch {
        Write-Verbose "Failed to run Get-ClusterLog"
        Invoke-CatchActions
    }

    try {
        $clusterNodeFileVersions = Get-ClusterNodeFileVersions
        Save-DataInfoToFile -DataIn $clusterNodeFileVersions -SaveToLocation "$copyTo\ClusterNodeFileVersions" -SaveTextFile $false
        Save-DataInfoToFile -DataIn ($clusterNodeFileVersions.Files.Values) -SaveToLocation "$copyTo\ClusterNodeFileVersions" -SaveXMLFile $false -FormatList $false
    } catch {
        Write-Verbose "Failed to run Get-ClusterNodeFileVersions"
        Invoke-CatchActions
    }

    $params = @{
        RegistryPath    = "HKLM:Cluster"
        SaveName        = "Cluster_Hive"
        SaveToPath      = $copyTo
        UseGetChildItem = $true
    }
    Save-RegistryHive @params
    Invoke-ZipFolder -Folder $copyTo
    Write-Verbose "Function Exit: Save-FailoverClusterInformation"
}

function Save-ServerInfoData {
    Write-Verbose("Function Enter: Save-ServerInfoData")
    $copyTo = $Script:RootCopyToDirectory + "\General_Server_Info"
    New-Item -ItemType Directory -Path $copyTo -Force | Out-Null

    #Get MSInfo from server
    msInfo32.exe /nfo (Add-ServerNameToFileName -FilePath ("{0}\msInfo.nfo" -f $copyTo))
    Write-Host "Waiting for msInfo32.exe process to end before moving on..." -ForegroundColor "Yellow"
    while ((Get-Process | Where-Object { $_.ProcessName -eq "msInfo32" }).ProcessName -eq "msInfo32") {
        Start-Sleep 5
    }

    $tlsRegistrySettingsName = "TLS_RegistrySettings"
    $tlsProtocol = @{
        RegistryPath    = "HKLM:SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols"
        SaveName        = $tlsRegistrySettingsName
        SaveToPath      = $copyTo
        UseGetChildItem = $true
    }
    Save-RegistryHive @tlsProtocol

    $net4Protocol = @{
        RegistryPath = "HKLM:SOFTWARE\Microsoft\.NETFramework\v4.0.30319"
        SaveName     = "NET4_$tlsRegistrySettingsName"
        SaveToPath   = $copyTo
    }
    Save-RegistryHive @net4Protocol

    $net4WowProtocol = @{
        RegistryPath = "HKLM:SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319"
        SaveName     = "NET4_Wow_$tlsRegistrySettingsName"
        SaveToPath   = $copyTo
    }
    Save-RegistryHive @net4WowProtocol

    $net2Protocol = @{
        RegistryPath = "HKLM:SOFTWARE\Microsoft\.NETFramework\v2.0.50727"
        SaveName     = "NET2_$tlsRegistrySettingsName"
        SaveToPath   = $copyTo
    }
    Save-RegistryHive @net2Protocol

    $net2WowProtocol = @{
        RegistryPath = "HKLM:SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v2.0.50727"
        SaveName     = "NET2_Wow_$tlsRegistrySettingsName"
        SaveToPath   = $copyTo
    }
    Save-RegistryHive @net2WowProtocol

    #Running Processes #35
    Save-DataInfoToFile -dataIn (Get-Process) -SaveToLocation ("{0}\Running_Processes" -f $copyTo) -FormatList $false

    #Services Information #36
    Save-DataInfoToFile -dataIn (Get-Service) -SaveToLocation ("{0}\Services_Information" -f $copyTo) -FormatList $false

    #VSSAdmin Information #39
    Save-DataInfoToFile -DataIn (vssadmin list Writers) -SaveToLocation ("{0}\VSS_Writers" -f $copyTo) -SaveXMLFile $false

    #Driver Information #34
    Save-DataInfoToFile -dataIn (Get-ChildItem ("{0}\System32\drivers" -f $env:SystemRoot) | Where-Object { $_.Name -like "*.sys" }) -SaveToLocation ("{0}\System32_Drivers" -f $copyTo)

    Save-DataInfoToFile -DataIn (Get-HotFix | Select-Object Source, Description, HotFixID, InstalledBy, InstalledOn) -SaveToLocation ("{0}\HotFixInfo" -f $copyTo)

    #TCP IP Networking Information #38
    Save-DataInfoToFile -DataIn (ipconfig /all) -SaveToLocation ("{0}\IPConfiguration" -f $copyTo) -SaveXMLFile $false
    Save-DataInfoToFile -DataIn (netstat -anob) -SaveToLocation ("{0}\NetStat_ANOB" -f $copyTo) -SaveXMLFile $false
    Save-DataInfoToFile -DataIn (route print) -SaveToLocation ("{0}\Network_Routes" -f $copyTo) -SaveXMLFile $false
    Save-DataInfoToFile -DataIn (arp -a) -SaveToLocation ("{0}\Network_ARP" -f $copyTo) -SaveXMLFile $false
    Save-DataInfoToFile -DataIn (netstat -naTo) -SaveToLocation ("{0}\Netstat_NATO" -f $copyTo) -SaveXMLFile $false
    Save-DataInfoToFile -DataIn (netstat -es) -SaveToLocation ("{0}\Netstat_ES" -f $copyTo) -SaveXMLFile $false

    #IPsec
    Save-DataInfoToFile -DataIn (netsh ipsec dynamic show all) -SaveToLocation ("{0}\IPsec_netsh_dynamic" -f $copyTo) -SaveXMLFile $false
    Save-DataInfoToFile -DataIn (netsh ipsec static show all) -SaveToLocation ("{0}\IPsec_netsh_static" -f $copyTo) -SaveXMLFile $false

    #FLTMC
    Save-DataInfoToFile -DataIn (fltmc) -SaveToLocation ("{0}\FLTMC_FilterDrivers" -f $copyTo) -SaveXMLFile $false
    Save-DataInfoToFile -DataIn (fltmc volumes) -SaveToLocation ("{0}\FLTMC_Volumes" -f $copyTo) -SaveXMLFile $false
    Save-DataInfoToFile -DataIn (fltmc instances) -SaveToLocation ("{0}\FLTMC_Instances" -f $copyTo) -SaveXMLFile $false

    Save-DataInfoToFile -DataIn (TaskList /M) -SaveToLocation ("{0}\TaskList_Modules" -f $copyTo) -SaveXMLFile $false

    if (!$Script:localServerObject.Edge) {

        $params = @{
            RegistryPath    = "HKLM:SOFTWARE\Microsoft\Exchange"
            SaveName        = "Exchange_Registry_Hive"
            SaveToPath      = $copyTo
            UseGetChildItem = $true
        }
        Save-RegistryHive @params

        $params = @{
            RegistryPath    = "HKLM:SOFTWARE\Microsoft\ExchangeServer"
            SaveName        = "ExchangeServer_Registry_Hive"
            SaveToPath      = $copyTo
            UseGetChildItem = $true
        }
        Save-RegistryHive @params
    }

    Save-DataInfoToFile -DataIn (gpResult /R /Z) -SaveToLocation ("{0}\GPResult" -f $copyTo) -SaveXMLFile $false
    gpResult /H (Add-ServerNameToFileName -FilePath ("{0}\GPResult.html" -f $copyTo))

    #Storage Information
    if (Test-CommandExists -command "Get-Volume") {
        Save-DataInfoToFile -DataIn (Get-Volume) -SaveToLocation ("{0}\Volume" -f $copyTo)
    } else {
        Write-Verbose("Get-Volume isn't a valid command")
    }

    if (Test-CommandExists -command "Get-Disk") {
        Save-DataInfoToFile -DataIn (Get-Disk) -SaveToLocation ("{0}\Disk" -f $copyTo)
    } else {
        Write-Verbose("Get-Disk isn't a valid command")
    }

    if (Test-CommandExists -command "Get-Partition") {
        Save-DataInfoToFile -DataIn (Get-Partition) -SaveToLocation ("{0}\Partition" -f $copyTo)
    } else {
        Write-Verbose("Get-Partition isn't a valid command")
    }

    Invoke-ZipFolder -Folder $copyTo
    Write-Verbose("Function Exit: Save-ServerInfoData")
}



function Get-RemoteRegistrySubKey {
    [CmdletBinding()]
    param(
        [string]$RegistryHive = "LocalMachine",
        [string]$MachineName,
        [string]$SubKey,
        [ScriptBlock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Attempting to open the Base Key $RegistryHive on Machine $MachineName"
        $regKey = $null
    }
    process {

        try {
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegistryHive, $MachineName)
            Write-Verbose "Attempting to open the Sub Key '$SubKey'"
            $regKey = $reg.OpenSubKey($SubKey)
            Write-Verbose "Opened Sub Key"
        } catch {
            Write-Verbose "Failed to open the registry"

            if ($null -ne $CatchActionFunction) {
                & $CatchActionFunction
            }
        }
    }
    end {
        return $regKey
    }
}

function Get-RemoteRegistryValue {
    [CmdletBinding()]
    param(
        [string]$RegistryHive = "LocalMachine",
        [string]$MachineName,
        [string]$SubKey,
        [string]$GetValue,
        [string]$ValueType,
        [ScriptBlock]$CatchActionFunction
    )

    <#
    Valid ValueType return values (case-sensitive)
    (https://docs.microsoft.com/en-us/dotnet/api/microsoft.win32.registryvaluekind?view=net-5.0)
    Binary = REG_BINARY
    DWord = REG_DWORD
    ExpandString = REG_EXPAND_SZ
    MultiString = REG_MULTI_SZ
    None = No data type
    QWord = REG_QWORD
    String = REG_SZ
    Unknown = An unsupported registry data type
    #>

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $registryGetValue = $null
    }
    process {

        try {

            $regSubKey = Get-RemoteRegistrySubKey -RegistryHive $RegistryHive `
                -MachineName $MachineName `
                -SubKey $SubKey

            if (-not ([System.String]::IsNullOrWhiteSpace($regSubKey))) {
                Write-Verbose "Attempting to get the value $GetValue"
                $registryGetValue = $regSubKey.GetValue($GetValue)
                Write-Verbose "Finished running GetValue()"

                if ($null -ne $registryGetValue -and
                    (-not ([System.String]::IsNullOrWhiteSpace($ValueType)))) {
                    Write-Verbose "Validating ValueType $ValueType"
                    $registryValueType = $regSubKey.GetValueKind($GetValue)
                    Write-Verbose "Finished running GetValueKind()"

                    if ($ValueType -ne $registryValueType) {
                        Write-Verbose "ValueType: $ValueType is different to the returned ValueType: $registryValueType"
                        $registryGetValue = $null
                    } else {
                        Write-Verbose "ValueType matches: $ValueType"
                    }
                }
            }
        } catch {
            Write-Verbose "Failed to get the value on the registry"

            if ($null -ne $CatchActionFunction) {
                & $CatchActionFunction
            }
        }
    }
    end {
        if ($registryGetValue.Length -le 100) {
            Write-Verbose "$($MyInvocation.MyCommand) Return Value: '$registryGetValue'"
        } else {
            Write-Verbose "$($MyInvocation.MyCommand) Return Value is too long to log"
        }
        return $registryGetValue
    }
}
function Save-WindowsEventLogs {

    Write-Verbose("Function Enter: Save-WindowsEventLogs")
    $baseSaveLocation = $Script:RootCopyToDirectory + "\Windows_Event_Logs"
    $SaveLogs = @{}
    $rootLogPath = "$env:SystemRoot\System32\WinEvt\Logs"
    $allLogPaths = Get-ChildItem $rootLogPath |
        ForEach-Object {
            $_.VersionInfo.FileName
        }

    if ($PassedInfo.AppSysLogs -or
        $PassedInfo.WindowsSecurityLogs) {

        $baseRegistryLocation = "SYSTEM\CurrentControlSet\Services\EventLog\"
        $logs = @()
        $baseParams = @{
            MachineName = $env:COMPUTERNAME
            GetValue    = "File"
        }

        Write-Verbose("Adding Windows Default Event Logging: AppSysLogs: $($PassedInfo.AppSysLogs) WindowsSecurityLogs: $($PassedInfo.WindowsSecurityLogs)")

        foreach ($logName in @("Application", "System", "MSExchange Management", "Security")) {

            if ((-not ($PassedInfo.WindowsSecurityLogs)) -and
                $logName -eq "Security") { continue }
            elseif ((-not ($PassedInfo.AppSysLogs)) -and
                $logName -ne "Security") { continue }

            Write-Verbose "Adding LogName: $logName"
            $params = $baseParams + @{
                SubKey = "$baseRegistryLocation$logName"
            }
            $logLocation = Get-RemoteRegistryValue @params

            if ($null -eq $logLocation) { $logLocation = "$rootLogPath\$logName.evtx" }
            $logs += $logLocation
        }

        $SaveLogs.Add("Windows-Logs", $logs)
    }

    if ($PassedInfo.ManagedAvailabilityLogs) {
        Write-Verbose("Adding Managed Availability Logs")

        $logs = $allLogPaths | Where-Object { $_.Contains("Microsoft-Exchange-ActiveMonitoring") }
        $SaveLogs.Add("Microsoft-Exchange-ActiveMonitoring", $Logs)

        $logs = $allLogPaths | Where-Object { $_.Contains("Microsoft-Exchange-ManagedAvailability") }
        $SaveLogs.Add("Microsoft-Exchange-ManagedAvailability", $Logs)
    }

    if ($PassedInfo.HighAvailabilityLogs) {
        Write-Verbose("Adding High Availability Logs")

        $logs = $allLogPaths | Where-Object { $_.Contains("Microsoft-Exchange-HighAvailability") }
        $SaveLogs.Add("Microsoft-Exchange-HighAvailability", $Logs)

        $logs = $allLogPaths | Where-Object { $_.Contains("Microsoft-Exchange-MailboxDatabaseFailureItems") }
        $SaveLogs.Add("Microsoft-Exchange-MailboxDatabaseFailureItems", $Logs)

        $logs = $allLogPaths | Where-Object { $_.Contains("Microsoft-Windows-FailoverClustering") }
        $SaveLogs.Add("Microsoft-Windows-FailoverClustering", $Logs)
    }

    foreach ($directory in $SaveLogs.Keys) {
        Write-Verbose("Working on directory: {0}" -f $directory)

        $logs = $SaveLogs[$directory]
        $saveLocation = "$baseSaveLocation\$directory"

        Copy-BulkItems -CopyToLocation $saveLocation -ItemsToCopyLocation $logs
        Get-ChildItem $saveLocation | Rename-Item -NewName { $_.Name -replace "%4", "-" }

        if ($directory -eq "Windows-Logs" -and
            $PassedInfo.AppSysLogsToXml) {
            try {
                Write-Verbose("starting to collect event logs and saving out to xml files.")
                Save-DataInfoToFile -DataIn (Get-EventLog Application -After ([DateTime]::Now - $PassedInfo.TimeSpan) -Before ([DateTime]::Now - $PassedInfo.EndTimeSpan)) -SaveToLocation ("{0}\Application" -f $saveLocation) -SaveTextFile $false
                Save-DataInfoToFile -DataIn (Get-EventLog System -After ([DateTime]::Now - $PassedInfo.TimeSpan) -Before ([DateTime]::Now - $PassedInfo.EndTimeSpan)) -SaveToLocation ("{0}\System" -f $saveLocation) -SaveTextFile $false
                Write-Verbose("end of collecting event logs and saving out to xml files.")
            } catch {
                Write-Verbose("Error occurred while trying to export out the Application and System logs to xml")
                Invoke-CatchActions
            }
        }

        Invoke-ZipFolder -Folder $saveLocation
    }
}
function Invoke-RemoteMain {
    [CmdletBinding()]
    param()
    Write-Verbose("Function Enter: Remote-Main")
    Invoke-ErrorMonitoring

    $Script:localServerObject = $PassedInfo.ServerObjects |
        Where-Object { $_.ServerName -eq $env:COMPUTERNAME }

    if ($null -eq $Script:localServerObject -or
        $Script:localServerObject.ServerName -ne $env:COMPUTERNAME) {
        Write-Host "Something went wrong trying to find the correct Server Object for this server. Stopping this instance of execution."
        exit
    }

    $Script:TotalBytesSizeCopied = 0
    $Script:TotalBytesSizeCompressed = 0
    $Script:AdditionalFreeSpaceCushionGB = $PassedInfo.StandardFreeSpaceInGBCheckSize
    $Script:CurrentFreeSpaceGB = Get-FreeSpace -FilePath ("{0}\" -f $Script:RootCopyToDirectory)
    $Script:FreeSpaceMinusCopiedAndCompressedGB = $Script:CurrentFreeSpaceGB
    $Script:localExInstall = Get-ExchangeInstallDirectory
    $Script:localExBin = $Script:localExInstall + "Bin\"
    $Script:taskActionList = New-Object "System.Collections.Generic.List[object]"
    #############################################
    #                                           #
    #              Exchange 2013 +              #
    #                                           #
    #############################################

    if ($Script:localServerObject.Version -ge 15) {
        Write-Verbose("Server Version greater than 15")

        if ($PassedInfo.EWSLogs) {

            if ($Script:localServerObject.Mailbox) {
                Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\EWS" "EWS_BE_Logs"
            }

            if ($Script:localServerObject.CAS) {
                Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\HttpProxy\Ews" "EWS_Proxy_Logs"
            }
        }

        if ($PassedInfo.RPCLogs) {

            if ($Script:localServerObject.Mailbox) {
                Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\RPC Client Access" "RCA_Logs"
            }

            if ($Script:localServerObject.CAS) {
                Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\HttpProxy\RpcHttp" "RCA_Proxy_Logs"
            }

            if (-not($Script:localServerObject.Edge)) {
                Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\RpcHttp" "RPC_Http_Logs"
            }
        }

        if ($Script:localServerObject.CAS -and $PassedInfo.EASLogs) {
            Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\HttpProxy\Eas" "EAS_Proxy_Logs"
        }

        if ($PassedInfo.AutoDLogs) {

            if ($Script:localServerObject.Mailbox) {
                Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\Autodiscover" "AutoD_Logs"
            }

            if ($Script:localServerObject.CAS) {
                Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\HttpProxy\Autodiscover" "AutoD_Proxy_Logs"
            }
        }

        if ($PassedInfo.OWALogs) {

            if ($Script:localServerObject.Mailbox) {
                Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\OWA" "OWA_Logs"
            }

            if ($Script:localServerObject.CAS) {
                Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\HttpProxy\OwaCalendar" "OWA_Proxy_Calendar_Logs"
                Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\HttpProxy\Owa" "OWA_Proxy_Logs"
            }
        }

        if ($PassedInfo.ADDriverLogs) {
            Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\ADDriver" "AD_Driver_Logs"
        }

        if ($PassedInfo.MapiLogs) {

            if ($Script:localServerObject.Mailbox -and $Script:localServerObject.Version -eq 15) {
                Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\MAPI Client Access" "MAPI_Logs"
            } elseif ($Script:localServerObject.Mailbox) {
                Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\MapiHttp\Mailbox" "MAPI_Logs"
            }

            if ($Script:localServerObject.CAS) {
                Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\HttpProxy\Mapi" "MAPI_Proxy_Logs"
            }
        }

        if ($PassedInfo.ECPLogs) {

            if ($Script:localServerObject.Mailbox) {
                Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\ECP" "ECP_Logs"
            }

            if ($Script:localServerObject.CAS) {
                Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\HttpProxy\Ecp" "ECP_Proxy_Logs"
            }
        }

        if ($Script:localServerObject.Mailbox -and $PassedInfo.SearchLogs) {
            Add-LogCopyBasedOffTimeTaskAction "$Script:localExBin`Search\Ceres\Diagnostics\Logs" "Search_Diag_Logs"
            Add-LogCopyBasedOffTimeTaskAction "$Script:localExBin`Search\Ceres\Diagnostics\ETLTraces" "Search_Diag_ETLs"
            Add-LogCopyFullTaskAction "$Script:localExInstall`Logging\Search" "Search"
            Add-LogCopyFullTaskAction "$Script:localExInstall`Logging\Monitoring\Search" "Search_Monitoring"

            if ($Script:localServerObject.Version -ge 19) {
                Add-LogCopyBasedOffTimeTaskAction "$Script:localExInstall`Logging\BigFunnelMetricsCollectionAssistant" "BigFunnelMetricsCollectionAssistant"
                Add-LogCopyBasedOffTimeTaskAction  "$Script:localExInstall`Logging\BigFunnelQueryParityAssistant" "BigFunnelQueryParityAssistant" #This might not provide anything
                Add-LogCopyBasedOffTimeTaskAction "$Script:localExInstall`Logging\BigFunnelRetryFeederTimeBasedAssistant" "BigFunnelRetryFeederTimeBasedAssistant"
            }
        }

        if ($PassedInfo.DailyPerformanceLogs) {
            #Daily Performance Logs are always by days worth
            $copyFrom = "$Script:localExInstall`Logging\Diagnostics\DailyPerformanceLogs"

            try {
                $logmanOutput = logman ExchangeDiagnosticsDailyPerformanceLog
                $logmanRootPath = $logmanOutput | Select-String "Root Path:"

                if (!$logmanRootPath.ToString().Contains($copyFrom)) {
                    $copyFrom = $logmanRootPath.ToString().Replace("Root Path:", "").Trim()
                    Write-Verbose "Changing the location to get the daily performance logs to '$copyFrom'"
                }
            } catch {
                Write-Verbose "Couldn't get logman info to verify Daily Performance Logs location"
                Invoke-CatchActions
            }
            Add-LogCopyBasedOffTimeTaskAction $copyFrom "Daily_Performance_Logs"
        }

        if ($PassedInfo.ManagedAvailabilityLogs) {
            Add-LogCopyFullTaskAction "$Script:localExInstall`Logging\Monitoring" "ManagedAvailabilityMonitoringLogs"
        }

        if ($PassedInfo.OABLogs) {
            Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\HttpProxy\OAB" "OAB_Proxy_Logs"
            Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\OABGeneratorLog" "OAB_Generation_Logs"
            Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\OABGeneratorSimpleLog" "OAB_Generation_Simple_Logs"
            Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\MAPI AddressBook Service" "MAPI_AddressBook_Service_Logs"
        }

        if ($PassedInfo.PowerShellLogs) {
            Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\HttpProxy\PowerShell" "PowerShell_Proxy_Logs"
            Add-LogCopyFullTaskAction "$Script:localExInstall`Logging\CmdletInfra" "CmdletInfra_Logs"
        }

        if ($Script:localServerObject.DAGMember -and
            $PassedInfo.DAGInformation) {
            Add-TaskAction "Save-FailoverClusterInformation"
        }

        if ($PassedInfo.MitigationService) {
            Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\MitigationService" "Mitigation_Service_Logs"
        }

        if ($PassedInfo.MailboxAssistantsLogs) {
            Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\MailboxAssistantsLog" "Mailbox_Assistants_Logs"
            Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\MailboxAssistantsSlaReportLog" "Mailbox_Assistants_Sla_Report_Logs"

            if ($Script:localServerObject.Version -eq 15) {
                Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\MailboxAssistantsDatabaseSlaLog" "Mailbox_Assistants_Database_Sla_Logs"
            }
        }

        if ($PassedInfo.PipelineTracingLogs) {

            if ($Script:localServerObject.Hub -or
                $Script:localServerObject.Edge) {
                Add-LogCopyFullTaskAction $Script:localServerObject.TransportInfo.HubLoggingInfo.PipelineTracingPath "Hub_Pipeline_Tracing_Logs"
            }

            if ($Script:localServerObject.Mailbox) {
                Add-LogCopyFullTaskAction $Script:localServerObject.TransportInfo.MBXLoggingInfo.PipelineTracingPath "Mailbox_Pipeline_Tracing_Logs"
            }
        }
    }

    ############################################
    #                                          #
    #              Exchange 2010               #
    #                                          #
    ############################################
    if ($Script:localServerObject.Version -eq 14) {

        if ($Script:localServerObject.CAS) {

            if ($PassedInfo.RPCLogs) {
                Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\RPC Client Access" "RCA_Logs"
            }

            if ($PassedInfo.EWSLogs) {
                Add-DefaultLogCopyTaskAction "$Script:localExInstall`Logging\EWS" "EWS_BE_Logs"
            }
        }
    }

    ############################################
    #                                          #
    #          All Exchange Versions           #
    #                                          #
    ############################################
    if ($PassedInfo.AnyTransportSwitchesEnabled -and
        $Script:localServerObject.TransportInfoCollect) {

        if ($PassedInfo.MessageTrackingLogs -and
            (-not ($Script:localServerObject.Version -eq 15 -and
                $Script:localServerObject.CASOnly))) {
            Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.TransportInfo.HubLoggingInfo.MessageTrackingLogPath "Message_Tracking_Logs" $false
        }

        if ($PassedInfo.HubProtocolLogs -and
            (-not ($Script:localServerObject.Version -eq 15 -and
                $Script:localServerObject.CASOnly))) {
            Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.TransportInfo.HubLoggingInfo.ReceiveProtocolLogPath "Hub_Receive_Protocol_Logs"
            Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.TransportInfo.HubLoggingInfo.SendProtocolLogPath "Hub_Send_Protocol_Logs"
        }

        if ($PassedInfo.HubConnectivityLogs -and
            (-not ($Script:localServerObject.Version -eq 15 -and
                $Script:localServerObject.CASOnly))) {
            Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.TransportInfo.HubLoggingInfo.ConnectivityLogPath "Hub_Connectivity_Logs"
        }

        if ($PassedInfo.QueueInformation -and
            (-not ($Script:localServerObject.Version -eq 15 -and
                $Script:localServerObject.CASOnly))) {

            if ($Script:localServerObject.Version -ge 15 -and
                $null -ne $Script:localServerObject.TransportInfo.HubLoggingInfo.QueueLogPath) {
                Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.TransportInfo.HubLoggingInfo.QueueLogPath "Queue_V15_Data"
            }
        }

        if ($PassedInfo.TransportConfig) {

            $items = @()
            if ($Script:localServerObject.Version -ge 15 -and (-not($Script:localServerObject.Edge))) {
                $items += $Script:localExBin + "\EdgeTransport.exe.config"
                $items += $Script:localExBin + "\MSExchangeFrontEndTransport.exe.config"
                $items += $Script:localExBin + "\MSExchangeDelivery.exe.config"
                $items += $Script:localExBin + "\MSExchangeSubmission.exe.config"
            } else {
                $items += $Script:localExBin + "\EdgeTransport.exe.config"
            }

            # TODO: Make into a task vs in the main loop
            Copy-BulkItems -CopyToLocation ($Script:RootCopyToDirectory + "\Transport_Configuration") -ItemsToCopyLocation $items
        }

        if ($PassedInfo.TransportAgentLogs) {

            if ($Script:localServerObject.CAS) {
                Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.TransportInfo.FELoggingInfo.AgentLogPath "FE_Transport_Agent_Logs"
            }

            if ($Script:localServerObject.Hub -or
                $Script:localServerObject.Edge) {
                Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.TransportInfo.HubLoggingInfo.AgentLogPath "Hub_Transport_Agent_Logs"
            }

            if ($Script:localServerObject.Mailbox) {
                Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.TransportInfo.MBXLoggingInfo.MailboxSubmissionAgentLogPath "Mbx_Submission_Transport_Agent_Logs"
                Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.TransportInfo.MBXLoggingInfo.MailboxDeliveryAgentLogPath "Mbx_Delivery_Transport_Agent_Logs"
            }
        }

        if ($PassedInfo.TransportRoutingTableLogs) {

            if ($Script:localServerObject.Version -ne 15 -and
                (-not ($Script:localServerObject.Edge))) {
                Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.TransportInfo.FELoggingInfo.RoutingTableLogPath "FE_Transport_Routing_Table_Logs"
            }

            if ($Script:localServerObject.Hub -or
                $Script:localServerObject.Edge) {
                Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.TransportInfo.HubLoggingInfo.RoutingTableLogPath "Hub_Transport_Routing_Table_Logs"
            }

            if ($Script:localServerObject.Version -ne 15 -and
                (-not ($Script:localServerObject.Edge))) {
                Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.TransportInfo.MBXLoggingInfo.RoutingTableLogPath "Mbx_Transport_Routing_Table_Logs"
            }
        }

        #Exchange 2013+ only
        if ($Script:localServerObject.Version -ge 15 -and
            (-not($Script:localServerObject.Edge))) {

            if ($PassedInfo.FrontEndConnectivityLogs -and
                (-not ($Script:localServerObject.Version -eq 15 -and
                    $Script:localServerObject.MailboxOnly))) {
                Write-Verbose("Collecting FrontEndConnectivityLogs")
                Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.TransportInfo.FELoggingInfo.ConnectivityLogPath "FE_Connectivity_Logs"
            }

            if ($PassedInfo.FrontEndProtocolLogs -and
                (-not ($Script:localServerObject.Version -eq 15 -and
                    $Script:localServerObject.MailboxOnly))) {
                Write-Verbose("Collecting FrontEndProtocolLogs")
                Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.TransportInfo.FELoggingInfo.ReceiveProtocolLogPath "FE_Receive_Protocol_Logs"
                Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.TransportInfo.FELoggingInfo.SendProtocolLogPath "FE_Send_Protocol_Logs"
            }

            if ($PassedInfo.MailboxConnectivityLogs -and
                (-not ($Script:localServerObject.Version -eq 15 -and
                    $Script:localServerObject.CASOnly))) {
                Write-Verbose("Collecting MailboxConnectivityLogs")
                Add-LogCopyBasedOffTimeTaskAction "$($Script:localServerObject.TransportInfo.MBXLoggingInfo.ConnectivityLogPath)\Delivery" "MBX_Delivery_Connectivity_Logs"
                Add-LogCopyBasedOffTimeTaskAction "$($Script:localServerObject.TransportInfo.MBXLoggingInfo.ConnectivityLogPath)\Submission" "MBX_Submission_Connectivity_Logs"
            }

            if ($PassedInfo.MailboxProtocolLogs -and
                (-not ($Script:localServerObject.Version -eq 15 -and
                    $Script:localServerObject.CASOnly))) {
                Write-Verbose("Collecting MailboxProtocolLogs")
                Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.TransportInfo.MBXLoggingInfo.ReceiveProtocolLogPath "MBX_Receive_Protocol_Logs"
                Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.TransportInfo.MBXLoggingInfo.SendProtocolLogPath "MBX_Send_Protocol_Logs"
            }

            if ($PassedInfo.MailboxDeliveryThrottlingLogs -and
                (!($Script:localServerObject.Version -eq 15 -and
                    $Script:localServerObject.CASOnly))) {
                Write-Verbose("Collecting Mailbox Delivery Throttling Logs")
                Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.TransportInfo.MBXLoggingInfo.MailboxDeliveryThrottlingLogPath "MBX_Delivery_Throttling_Logs"
            }
        }
    }

    if ($PassedInfo.ImapLogs) {
        Write-Verbose("Collecting IMAP Logs")
        Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.ImapLogsLocation "Imap_Logs"
    }

    if ($PassedInfo.PopLogs) {
        Write-Verbose("Collecting POP Logs")
        Add-LogCopyBasedOffTimeTaskAction $Script:localServerObject.PopLogsLocation "Pop_Logs"
    }

    if ($PassedInfo.IISLogs) {

        Get-IISLogDirectory |
            ForEach-Object {
                $copyTo = "{0}\IIS_{1}_Logs" -f $Script:RootCopyToDirectory, ($_.Substring($_.LastIndexOf("\") + 1))
                Add-LogCopyBasedOffTimeTaskAction $_ $copyTo
            }

        Add-LogCopyBasedOffTimeTaskAction "$env:SystemRoot`\System32\LogFiles\HTTPERR" "HTTPERR_Logs"
    }

    if ($PassedInfo.ServerInformation) {
        Add-TaskAction "Save-ServerInfoData"
    }

    if ($PassedInfo.ExPerfWiz) {
        Add-TaskAction "Save-LogmanExPerfWizData"
    }

    if ($PassedInfo.ExMon) {
        Add-TaskAction "Save-LogmanExMonData"
    }

    Add-TaskAction "Save-WindowsEventLogs"
    #Execute the cmdlets
    foreach ($taskAction in $Script:taskActionList) {
        Write-Verbose(("Task Action: $(GetTaskActionToString $taskAction)"))

        try {
            $params = $taskAction.Parameters

            if ($null -ne $params) {
                & $taskAction.FunctionName @params -ErrorAction Stop
            } else {
                & $taskAction.FunctionName -ErrorAction Stop
            }
        } catch {
            Write-Verbose("Failed to finish running command: $(GetTaskActionToString $taskAction)")
            Invoke-CatchActions
        }
    }

    if ($Error.Count -ne 0) {
        Save-DataInfoToFile -DataIn $Error -SaveToLocation ("$Script:RootCopyToDirectory\AllErrors")
        Save-DataInfoToFile -DataIn (Get-UnhandledErrors) -SaveToLocation ("$RootCopyToDirectory\UnhandledErrors")
        Save-DataInfoToFile -DataIn (Get-HandledErrors) -SaveToLocation ("$RootCopyToDirectory\HandledErrors")
    } else {
        Write-Verbose ("No errors occurred within the script")
    }
}

        try {

            if ($PassedInfo.ByPass -ne $true) {
                $Script:RootCopyToDirectory = "{0}{1}" -f $PassedInfo.RootFilePath, $env:COMPUTERNAME
                $Script:Logger = Get-NewLoggerInstance -LogName "ExchangeLogCollector-Instance-Debug" -LogDirectory $Script:RootCopyToDirectory
                SetWriteHostManipulateObjectAction ${Function:Get-ManipulateWriteHostValue}
                SetWriteVerboseManipulateMessageAction ${Function:Get-ManipulateWriteVerboseValue}
                SetWriteHostAction ${Function:Write-DebugLog}
                SetWriteVerboseAction ${Function:Write-DebugLog}

                if ($PassedInfo.ScriptDebug) {
                    $Script:VerbosePreference = "Continue"
                }

                Write-Verbose("Root Copy To Directory: $Script:RootCopyToDirectory")
                Invoke-RemoteMain
            } else {
                Write-Verbose("Loading common functions")
            }
        } catch {
            Write-Host "An error occurred in Invoke-RemoteFunctions" -ForegroundColor "Red"
            Invoke-CatchActions
            #This is a bad place to catch the error that just occurred
            #Being that there is a try catch block around each command that we run now, we should never hit an issue here unless it is is prior to that.
            Write-Verbose "Critical Failure occurred."
        } finally {
            Write-Verbose("Exiting: Invoke-RemoteFunctions")
            Write-Verbose("[double]TotalBytesSizeCopied: {0} | [double]TotalBytesSizeCompressed: {1} | [double]AdditionalFreeSpaceCushionGB: {2} | [double]CurrentFreeSpaceGB: {3} | [double]FreeSpaceMinusCopiedAndCompressedGB: {4}" -f $Script:TotalBytesSizeCopied,
                $Script:TotalBytesSizeCompressed,
                $Script:AdditionalFreeSpaceCushionGB,
                $Script:CurrentFreeSpaceGB,
                $Script:FreeSpaceMinusCopiedAndCompressedGB)
        }
    }

    # Need to dot load the files outside of the remote functions and AFTER them to avoid issues with encapsulation and dependencies

function Confirm-Administrator {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )

    return $currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )
}


function Invoke-CatchActionError {
    [CmdletBinding()]
    param(
        [ScriptBlock]$CatchActionFunction
    )

    if ($null -ne $CatchActionFunction) {
        & $CatchActionFunction
    }
}

function Invoke-CatchActionErrorLoop {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [int]$CurrentErrors,
        [Parameter(Mandatory = $false, Position = 1)]
        [ScriptBlock]$CatchActionFunction
    )
    process {
        if ($null -ne $CatchActionFunction -and
            $Error.Count -ne $CurrentErrors) {
            $i = 0
            while ($i -lt ($Error.Count - $currentErrors)) {
                & $CatchActionFunction $Error[$i]
                $i++
            }
        }
    }
}

# Confirm that either Remote Shell or EMS is loaded from an Edge Server, Exchange Server, or a Tools box.
# It does this by also initializing the session and running Get-EventLogLevel. (Server Management RBAC right)
# All script that require Confirm-ExchangeShell should be at least using Server Management RBAC right for the user running the script.
function Confirm-ExchangeShell {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [bool]$LoadExchangeShell = $true,

        [Parameter(Mandatory = $false)]
        [ScriptBlock]$CatchActionFunction
    )

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed: LoadExchangeShell: $LoadExchangeShell"
        $currentErrors = $Error.Count
        $edgeTransportKey = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EdgeTransportRole'
        $setupKey = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup'
        $remoteShell = (-not(Test-Path $setupKey))
        $toolsServer = (Test-Path $setupKey) -and
            (-not(Test-Path $edgeTransportKey)) -and
            ($null -eq (Get-ItemProperty -Path $setupKey -Name "Services" -ErrorAction SilentlyContinue))
        Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction

        function IsExchangeManagementSession {
            [OutputType("System.Boolean")]
            param(
                [ScriptBlock]$CatchActionFunction
            )

            $getEventLogLevelCallSuccessful = $false
            $isExchangeManagementShell = $false

            try {
                $currentErrors = $Error.Count
                $attempts = 0
                do {
                    $eventLogLevel = Get-EventLogLevel -ErrorAction Stop | Select-Object -First 1
                    $attempts++
                    if ($attempts -ge 5) {
                        throw "Failed to run Get-EventLogLevel too many times."
                    }
                } while ($null -eq $eventLogLevel)
                $getEventLogLevelCallSuccessful = $true
                foreach ($e in $eventLogLevel) {
                    Write-Verbose "Type is: $($e.GetType().Name) BaseType is: $($e.GetType().BaseType)"
                    if (($e.GetType().Name -eq "EventCategoryObject") -or
                        (($e.GetType().Name -eq "PSObject") -and
                            ($null -ne $e.SerializationData))) {
                        $isExchangeManagementShell = $true
                    }
                }
                Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction
            } catch {
                Write-Verbose "Failed to run Get-EventLogLevel"
                Invoke-CatchActionError $CatchActionFunction
            }

            return [PSCustomObject]@{
                CallWasSuccessful = $getEventLogLevelCallSuccessful
                IsManagementShell = $isExchangeManagementShell
            }
        }
    }
    process {
        $isEMS = IsExchangeManagementSession $CatchActionFunction
        if ($isEMS.CallWasSuccessful) {
            Write-Verbose "Exchange PowerShell Module already loaded."
        } else {
            if (-not ($LoadExchangeShell)) { return }

            #Test 32 bit process, as we can't see the registry if that is the case.
            if (-not ([System.Environment]::Is64BitProcess)) {
                Write-Warning "Open a 64 bit PowerShell process to continue"
                return
            }

            if (Test-Path "$setupKey") {
                Write-Verbose "We are on Exchange 2013 or newer"

                try {
                    $currentErrors = $Error.Count
                    if (Test-Path $edgeTransportKey) {
                        Write-Verbose "We are on Exchange Edge Transport Server"
                        [xml]$PSSnapIns = Get-Content -Path "$env:ExchangeInstallPath\Bin\exShell.psc1" -ErrorAction Stop

                        foreach ($PSSnapIn in $PSSnapIns.PSConsoleFile.PSSnapIns.PSSnapIn) {
                            Write-Verbose ("Trying to add PSSnapIn: {0}" -f $PSSnapIn.Name)
                            Add-PSSnapin -Name $PSSnapIn.Name -ErrorAction Stop
                        }

                        Import-Module $env:ExchangeInstallPath\bin\Exchange.ps1 -ErrorAction Stop
                    } else {
                        Import-Module $env:ExchangeInstallPath\bin\RemoteExchange.ps1 -ErrorAction Stop
                        Connect-ExchangeServer -Auto -ClientApplication:ManagementShell
                    }
                    Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction

                    Write-Verbose "Imported Module. Trying Get-EventLogLevel Again"
                    $isEMS = IsExchangeManagementSession $CatchActionFunction
                    if (($isEMS.CallWasSuccessful) -and
                        ($isEMS.IsManagementShell)) {
                        Write-Verbose "Successfully loaded Exchange Management Shell"
                    } else {
                        Write-Warning "Something went wrong while loading the Exchange Management Shell"
                    }
                } catch {
                    Write-Warning "Failed to Load Exchange PowerShell Module..."
                    Invoke-CatchActionError $CatchActionFunction
                }
            } else {
                Write-Verbose "Not on an Exchange or Tools server"
            }
        }
    }
    end {

        $returnObject = [PSCustomObject]@{
            ShellLoaded = $isEMS.CallWasSuccessful
            Major       = ((Get-ItemProperty -Path $setupKey -Name "MsiProductMajor" -ErrorAction SilentlyContinue).MsiProductMajor)
            Minor       = ((Get-ItemProperty -Path $setupKey -Name "MsiProductMinor" -ErrorAction SilentlyContinue).MsiProductMinor)
            Build       = ((Get-ItemProperty -Path $setupKey -Name "MsiBuildMajor" -ErrorAction SilentlyContinue).MsiBuildMajor)
            Revision    = ((Get-ItemProperty -Path $setupKey -Name "MsiBuildMinor" -ErrorAction SilentlyContinue).MsiBuildMinor)
            EdgeServer  = $isEMS.CallWasSuccessful -and (Test-Path $setupKey) -and (Test-Path $edgeTransportKey)
            ToolsOnly   = $isEMS.CallWasSuccessful -and $toolsServer
            RemoteShell = $isEMS.CallWasSuccessful -and $remoteShell
            EMS         = $isEMS.IsManagementShell
        }

        return $returnObject
    }
}




function Get-ExchangeContainer {
    [CmdletBinding()]
    [OutputType([System.DirectoryServices.DirectoryEntry])]
    param ()

    $rootDSE = [ADSI]("LDAP://$([System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Name)/RootDSE")
    $exchangeContainerPath = ("CN=Microsoft Exchange,CN=Services," + $rootDSE.configurationNamingContext)
    $exchangeContainer = [ADSI]("LDAP://" + $exchangeContainerPath)
    Write-Verbose "Exchange Container Path: $($exchangeContainer.path)"
    return $exchangeContainer
}

function Get-OrganizationContainer {
    [CmdletBinding()]
    [OutputType([System.DirectoryServices.DirectoryEntry])]
    param ()

    $exchangeContainer = Get-ExchangeContainer
    $searcher = New-Object System.DirectoryServices.DirectorySearcher($exchangeContainer, "(objectClass=msExchOrganizationContainer)", @("distinguishedName"))
    return $searcher.FindOne().GetDirectoryEntry()
}

function Get-VirtualDirectoriesLdap {

    $authTypeEnum = @"
    namespace AuthMethods
    {
        using System;
        [Flags]
        public enum AuthenticationMethodFlags
        {
            None = 0,
            Basic = 1,
            Ntlm = 2,
            Fba = 4,
            Digest = 8,
            WindowsIntegrated = 16,
            LiveIdFba = 32,
            LiveIdBasic = 64,
            WSSecurity = 128,
            Certificate = 256,
            NegoEx = 512,
            // Exchange 2013
            OAuth = 1024,
            Adfs = 2048,
            Kerberos = 4096,
            Negotiate = 8192,
            LiveIdNegotiate = 16384,
        }
    }
"@

    Write-Host "Collecting Virtual Directory Information..."
    Add-Type -TypeDefinition $authTypeEnum -Language CSharp

    $searcher = New-Object DirectoryServices.DirectorySearcher
    $searcher.filter = "(&(objectClass=msExchVirtualDirectory)(!objectClass=container))"
    $searcher.SearchRoot = Get-OrganizationContainer
    $searcher.CacheResults = $false
    $searcher.SearchScope = "Subtree"
    $searcher.PageSize = 1000

    # Get all the results
    $colResults = $searcher.FindAll()
    $objects = @()

    # Loop through the results and
    foreach ($objResult in $colResults) {
        $objItem = $objResult.getDirectoryEntry()
        $objProps = $objItem.Properties

        $place = $objResult.Path.IndexOf("CN=Protocols,CN=")
        $ServerDN = [ADSI]("LDAP://" + $objResult.Path.SubString($place, ($objResult.Path.Length - $place)).Replace("CN=Protocols,", ""))
        [string]$Site = $serverDN.Properties.msExchServerSite.ToString().Split(",")[0].Replace("CN=", "")
        [string]$server = $serverDN.Properties.adminDisplayName.ToString()
        [string]$version = $serverDN.Properties.serialNumber.ToString()

        $obj = New-Object PSObject
        $obj | Add-Member -MemberType NoteProperty -Name Server -Value $server
        $obj | Add-Member -MemberType NoteProperty -Name Version -Value $version
        $obj | Add-Member -MemberType NoteProperty -Name Site -Value $Site
        [string]$var = $objProps.DistinguishedName.ToString().Split(",")[0].Replace("CN=", "")
        $obj | Add-Member -MemberType NoteProperty -Name VirtualDirectory -Value $var
        [string]$var = $objProps.msExchInternalHostName
        $obj | Add-Member -MemberType NoteProperty -Name InternalURL -Value $var

        if (-not [string]::IsNullOrEmpty($objProps.msExchInternalAuthenticationMethods)) {
            $obj | Add-Member -MemberType NoteProperty -Name InternalAuthenticationMethods -Value ([AuthMethods.AuthenticationMethodFlags]$objProps.msExchInternalAuthenticationMethods)
        } else {
            $obj | Add-Member -MemberType NoteProperty -Name InternalAuthenticationMethods -Value $null
        }

        [string]$var = $objProps.msExchExternalHostName
        $obj | Add-Member -MemberType NoteProperty -Name ExternalURL -Value $var

        if (-not [string]::IsNullOrEmpty($objProps.msExchExternalAuthenticationMethods)) {
            $obj | Add-Member -MemberType NoteProperty -Name ExternalAuthenticationMethods -Value ([AuthMethods.AuthenticationMethodFlags]$objProps.msExchExternalAuthenticationMethods)
        } else {
            $obj | Add-Member -MemberType NoteProperty -Name ExternalAuthenticationMethods -Value $null
        }

        if (-not [string]::IsNullOrEmpty($objProps.msExch2003Url)) {
            [string]$var = $objProps.msExch2003Url
            $obj | Add-Member -MemberType NoteProperty -Name Exchange2003URL  -Value $var
        } else {
            $obj | Add-Member -MemberType NoteProperty -Name Exchange2003URL -Value $null
        }

        [Array]$objects += $obj
    }

    return $objects
}
function Write-DataOnlyOnceOnMasterServer {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseUsingScopeModifierInNewRunSpaces', '', Justification = 'Can not use using for an env variable')]
    param()
    Write-Verbose("Enter Function: Write-DataOnlyOnceOnMasterServer")
    Write-Verbose("Writing only once data")

    if (!$Script:MasterServer.ToUpper().Contains($env:COMPUTERNAME.ToUpper())) {
        $serverName = Invoke-Command -ComputerName $Script:MasterServer -ScriptBlock { return $env:COMPUTERNAME }
        $RootCopyToDirectory = "\\{0}\{1}" -f $Script:MasterServer, (("{0}{1}" -f $Script:RootFilePath, $serverName).Replace(":", "$"))
    } else {
        $RootCopyToDirectory = "{0}{1}" -f $Script:RootFilePath, $env:COMPUTERNAME
    }

    if ($GetVDirs -and (-not($Script:EdgeRoleDetected))) {
        $target = $RootCopyToDirectory + "\ConfigNC_msExchVirtualDirectory_All.CSV"
        $data = (Get-VirtualDirectoriesLdap)
        $data | Sort-Object -Property Server | Export-Csv $target -NoTypeInformation
    }

    if ($OrganizationConfig) {
        $target = $RootCopyToDirectory + "\OrganizationConfig"
        $data = Get-OrganizationConfig
        Save-DataInfoToFile -dataIn (Get-OrganizationConfig) -SaveToLocation $target -AddServerName $false
    }

    if ($SendConnectors) {
        $create = $RootCopyToDirectory + "\Connectors"
        New-Item -ItemType Directory -Path $create -Force | Out-Null
        $saveLocation = $create + "\Send_Connectors"
        Save-DataInfoToFile -dataIn (Get-SendConnector) -SaveToLocation $saveLocation -AddServerName $false
    }

    if ($TransportConfig) {
        $target = $RootCopyToDirectory + "\TransportConfig"
        $data = Get-TransportConfig
        Save-DataInfoToFile -dataIn $data -SaveToLocation $target -AddServerName $false
    }

    if ($TransportRules) {
        $target = $RootCopyToDirectory + "\TransportRules"
        $data = Get-TransportRule

        # If no rules found, we want to report that.
        if ($null -ne $data) {
            Save-DataInfoToFile -dataIn $data -SaveToLocation $target -AddServerName $false
        } else {
            Save-DataInfoToFile -dataIn "No Transport Rules Found" -SaveXMLFile $false -SaveToLocation $target -AddServerName $false
        }
    }

    if ($AcceptedRemoteDomain) {
        $target = $RootCopyToDirectory + "\AcceptedDomain"
        $data = Get-AcceptedDomain
        Save-DataInfoToFile -dataIn $data -SaveToLocation $target -AddServerName $false

        $target = $RootCopyToDirectory + "\RemoteDomain"
        $data = Get-RemoteDomain
        Save-DataInfoToFile -dataIn $data -SaveToLocation $target -AddServerName $false
    }

    if ($Error.Count -ne 0) {
        Save-DataInfoToFile -DataIn $Error -SaveToLocation ("$RootCopyToDirectory\AllErrors")
        Save-DataInfoToFile -DataIn (Get-UnhandledErrors) -SaveToLocation ("$RootCopyToDirectory\UnhandledErrors")
        Save-DataInfoToFile -DataIn (Get-HandledErrors) -SaveToLocation ("$RootCopyToDirectory\HandledErrors")
    } else {
        Write-Verbose ("No errors occurred within the script")
    }

    Write-Verbose("Exiting Function: Write-DataOnlyOnceOnMasterServer")
}


function Get-DAGInformation {
    param(
        [Parameter(Mandatory = $true)][string]$DAGName
    )

    try {
        $dag = Get-DatabaseAvailabilityGroup $DAGName -Status -ErrorAction Stop
    } catch {
        Write-Verbose("Failed to run Get-DatabaseAvailabilityGroup on $DAGName")
        Invoke-CatchActions
    }

    try {
        $dagNetwork = Get-DatabaseAvailabilityGroupNetwork $DAGName -ErrorAction Stop
    } catch {
        Write-Verbose("Failed to run Get-DatabaseAvailabilityGroupNetwork on $DAGName")
        Invoke-CatchActions
    }

    #Now to get the Mailbox Database Information for each server in the DAG.
    $cacheDBCopyStatus = @{}
    $mailboxDatabaseInformationPerServer = @{}

    foreach ($server in $dag.Servers) {
        $serverName = $server.ToString()
        $getMailboxDatabases = Get-MailboxDatabase -Server $serverName -Status

        #Foreach of the mailbox databases on this server, we want to know the copy status
        #but we don't want to duplicate this work a lot, so we have a cache feature.
        $getMailboxDatabaseCopyStatusPerDB = @{}
        $getMailboxDatabases |
            ForEach-Object {
                $dbName = $_.Name

                if (!($cacheDBCopyStatus.ContainsKey($dbName))) {
                    $copyStatusForDB = Get-MailboxDatabaseCopyStatus $dbName\* -ErrorAction SilentlyContinue
                    $cacheDBCopyStatus.Add($dbName, $copyStatusForDB)
                } else {
                    $copyStatusForDB = $cacheDBCopyStatus[$dbName]
                }

                $getMailboxDatabaseCopyStatusPerDB.Add($dbName, $copyStatusForDB)
            }

        $serverDatabaseInformation = [PSCustomObject]@{
            MailboxDatabases                = $getMailboxDatabases
            MailboxDatabaseCopyStatusPerDB  = $getMailboxDatabaseCopyStatusPerDB
            MailboxDatabaseCopyStatusServer = (Get-MailboxDatabaseCopyStatus *\$serverName -ErrorAction SilentlyContinue)
        }

        $mailboxDatabaseInformationPerServer.Add($serverName, $serverDatabaseInformation)
    }

    return [PSCustomObject]@{
        DAGInfo             = $dag
        DAGNetworkInfo      = $dagNetwork
        MailboxDatabaseInfo = $mailboxDatabaseInformationPerServer
    }
}



# This function is used to determine the version of Exchange based off a build number or
# by providing the Exchange Version and CU and/or SU. This provides one location in the entire repository
# that is required to be updated for when a new release of Exchange is dropped.
function Get-ExchangeBuildVersionInformation {
    [CmdletBinding(DefaultParameterSetName = "AdminDisplayVersion")]
    param(
        [Parameter(ParameterSetName = "AdminDisplayVersion", Position = 1)]
        [object]$AdminDisplayVersion,

        [Parameter(ParameterSetName = "ExSetup")]
        [System.Version]$FileVersion,

        [Parameter(ParameterSetName = "VersionCU", Mandatory = $true)]
        [ValidateScript( { ValidateVersionParameter $_ } )]
        [string]$Version,

        [Parameter(ParameterSetName = "VersionCU", Mandatory = $true)]
        [ValidateScript( { ValidateCUParameter $_ } )]
        [string]$CU,

        [Parameter(ParameterSetName = "VersionCU", Mandatory = $false)]
        [ValidateScript( { ValidateSUParameter $_ } )]
        [string]$SU,

        [Parameter(ParameterSetName = "FindSUBuilds", Mandatory = $true)]
        [ValidateScript( { ValidateSUParameter $_ } )]
        [string]$FindBySUName,

        [Parameter(Mandatory = $false)]
        [ScriptBlock]$CatchActionFunction
    )
    begin {

        function GetBuildVersion {
            param(
                [Parameter(Position = 1)]
                [string]$ExchangeVersion,
                [Parameter(Position = 2)]
                [string]$CU,
                [Parameter(Position = 3)]
                [string]$SU
            )
            $cuResult = $exchangeBuildDictionary[$ExchangeVersion][$CU]

            if ((-not [string]::IsNullOrEmpty($SU)) -and
                $cuResult.SU.ContainsKey($SU)) {
                return $cuResult.SU[$SU]
            } else {
                return $cuResult.CU
            }
        }

        # Dictionary of Exchange Version/CU/SU to build number
        $exchangeBuildDictionary = GetExchangeBuildDictionary

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $exchangeMajorVersion = [string]::Empty
        $exchangeVersion = $null
        $supportedBuildNumber = $false
        $latestSUBuild = $false
        $extendedSupportDate = [string]::Empty
        $cuReleaseDate = [string]::Empty
        $friendlyName = [string]::Empty
        $cuLevel = [string]::Empty
        $suName = [string]::Empty
        $orgValue = 0
        $schemaValue = 0
        $mesoValue = 0
        $ex19 = "Exchange2019"
        $ex16 = "Exchange2016"
        $ex13 = "Exchange2013"
    }
    process {
        # Convert both input types to a [System.Version]
        try {
            if ($PSCmdlet.ParameterSetName -eq "FindSUBuilds") {
                foreach ($exchangeKey in $exchangeBuildDictionary.Keys) {
                    foreach ($cuKey in $exchangeBuildDictionary[$exchangeKey].Keys) {
                        if ($null -ne $exchangeBuildDictionary[$exchangeKey][$cuKey].SU -and
                            $exchangeBuildDictionary[$exchangeKey][$cuKey].SU.ContainsKey($FindBySUName)) {
                            Get-ExchangeBuildVersionInformation -FileVersion $exchangeBuildDictionary[$exchangeKey][$cuKey].SU[$FindBySUName]
                        }
                    }
                }
                return
            } elseif ($PSCmdlet.ParameterSetName -eq "VersionCU") {
                [System.Version]$exchangeVersion = GetBuildVersion -ExchangeVersion $Version -CU $CU -SU $SU
            } elseif ($PSCmdlet.ParameterSetName -eq "AdminDisplayVersion") {
                $AdminDisplayVersion = $AdminDisplayVersion.ToString()
                Write-Verbose "Passed AdminDisplayVersion: $AdminDisplayVersion"
                $split1 = $AdminDisplayVersion.Substring(($AdminDisplayVersion.IndexOf(" ")) + 1, 4).Split(".")
                $buildStart = $AdminDisplayVersion.LastIndexOf(" ") + 1
                $split2 = $AdminDisplayVersion.Substring($buildStart, ($AdminDisplayVersion.LastIndexOf(")") - $buildStart)).Split(".")
                [System.Version]$exchangeVersion = "$($split1[0]).$($split1[1]).$($split2[0]).$($split2[1])"
            } else {
                [System.Version]$exchangeVersion = $FileVersion
            }
        } catch {
            Write-Verbose "Failed to convert to system.version"
            Invoke-CatchActionError $CatchActionFunction
        }

        <#
            Exchange Build Numbers: https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates?view=exchserver-2019
            Exchange 2016 & 2019 AD Changes: https://learn.microsoft.com/en-us/exchange/plan-and-deploy/prepare-ad-and-domains?view=exchserver-2019
            Exchange 2013 AD Changes: https://learn.microsoft.com/en-us/exchange/prepare-active-directory-and-domains-exchange-2013-help
        #>
        if ($exchangeVersion.Major -eq 15 -and $exchangeVersion.Minor -eq 2) {
            Write-Verbose "Exchange 2019 is detected"
            $exchangeMajorVersion = "Exchange2019"
            $extendedSupportDate = "10/14/2025"
            $friendlyName = "Exchange 2019"

            #Latest Version AD Settings
            $schemaValue = 17003
            $mesoValue = 13243
            $orgValue = 16762

            switch ($exchangeVersion) {
                { $_ -ge (GetBuildVersion $ex19 "CU14") } {
                    $cuLevel = "CU14"
                    $cuReleaseDate = "02/13/2024"
                    $supportedBuildNumber = $true
                }
                (GetBuildVersion $ex19 "CU14" -SU "Mar24SU") { $latestSUBuild = $true }
                { $_ -lt (GetBuildVersion $ex19 "CU14") } {
                    $cuLevel = "CU13"
                    $cuReleaseDate = "05/03/2023"
                    $supportedBuildNumber = $true
                    $orgValue = 16761
                }
                (GetBuildVersion $ex19 "CU13" -SU "Mar24SU") { $latestSUBuild = $true }
                { $_ -lt (GetBuildVersion $ex19 "CU13") } {
                    $cuLevel = "CU12"
                    $cuReleaseDate = "04/20/2022"
                    $supportedBuildNumber = $false
                    $orgValue = 16760
                }
                { $_ -lt (GetBuildVersion $ex19 "CU12") } {
                    $cuLevel = "CU11"
                    $cuReleaseDate = "09/28/2021"
                    $mesoValue = 13242
                    $orgValue = 16759
                }
                (GetBuildVersion $ex19 "CU11" -SU "May22SU") { $mesoValue = 13243 }
                { $_ -lt (GetBuildVersion $ex19 "CU11") } {
                    $cuLevel = "CU10"
                    $cuReleaseDate = "06/29/2021"
                    $mesoValue = 13241
                    $orgValue = 16758
                }
                { $_ -lt (GetBuildVersion $ex19 "CU10") } {
                    $cuLevel = "CU9"
                    $cuReleaseDate = "03/16/2021"
                    $schemaValue = 17002
                    $mesoValue = 13240
                    $orgValue = 16757
                }
                { $_ -lt (GetBuildVersion $ex19 "CU9") } {
                    $cuLevel = "CU8"
                    $cuReleaseDate = "12/15/2020"
                    $mesoValue = 13239
                    $orgValue = 16756
                }
                { $_ -lt (GetBuildVersion $ex19 "CU8") } {
                    $cuLevel = "CU7"
                    $cuReleaseDate = "09/15/2020"
                    $schemaValue = 17001
                    $mesoValue = 13238
                    $orgValue = 16755
                }
                { $_ -lt (GetBuildVersion $ex19 "CU7") } {
                    $cuLevel = "CU6"
                    $cuReleaseDate = "06/16/2020"
                    $mesoValue = 13237
                    $orgValue = 16754
                }
                { $_ -lt (GetBuildVersion $ex19 "CU6") } {
                    $cuLevel = "CU5"
                    $cuReleaseDate = "03/17/2020"
                }
                { $_ -lt (GetBuildVersion $ex19 "CU5") } {
                    $cuLevel = "CU4"
                    $cuReleaseDate = "12/17/2019"
                }
                { $_ -lt (GetBuildVersion $ex19 "CU4") } {
                    $cuLevel = "CU3"
                    $cuReleaseDate = "09/17/2019"
                }
                { $_ -lt (GetBuildVersion $ex19 "CU3") } {
                    $cuLevel = "CU2"
                    $cuReleaseDate = "06/18/2019"
                }
                { $_ -lt (GetBuildVersion $ex19 "CU2") } {
                    $cuLevel = "CU1"
                    $cuReleaseDate = "02/12/2019"
                    $schemaValue = 17000
                    $mesoValue = 13236
                    $orgValue = 16752
                }
                { $_ -lt (GetBuildVersion $ex19 "CU1") } {
                    $cuLevel = "RTM"
                    $cuReleaseDate = "10/22/2018"
                    $orgValue = 16751
                }
            }
        } elseif ($exchangeVersion.Major -eq 15 -and $exchangeVersion.Minor -eq 1) {
            Write-Verbose "Exchange 2016 is detected"
            $exchangeMajorVersion = "Exchange2016"
            $extendedSupportDate = "10/14/2025"
            $friendlyName = "Exchange 2016"

            #Latest Version AD Settings
            $schemaValue = 15334
            $mesoValue = 13243
            $orgValue = 16223

            switch ($exchangeVersion) {
                { $_ -ge (GetBuildVersion $ex16 "CU23") } {
                    $cuLevel = "CU23"
                    $cuReleaseDate = "04/20/2022"
                    $supportedBuildNumber = $true
                }
                (GetBuildVersion $ex16 "CU23" -SU "Mar24SU") { $latestSUBuild = $true }
                { $_ -lt (GetBuildVersion $ex16 "CU23") } {
                    $cuLevel = "CU22"
                    $cuReleaseDate = "09/28/2021"
                    $supportedBuildNumber = $false
                    $mesoValue = 13242
                    $orgValue = 16222
                }
                (GetBuildVersion $ex16 "CU22" -SU "May22SU") { $mesoValue = 13243 }
                { $_ -lt (GetBuildVersion $ex16 "CU22") } {
                    $cuLevel = "CU21"
                    $cuReleaseDate = "06/29/2021"
                    $mesoValue = 13241
                    $orgValue = 16221
                }
                { $_ -lt (GetBuildVersion $ex16 "CU21") } {
                    $cuLevel = "CU20"
                    $cuReleaseDate = "03/16/2021"
                    $schemaValue = 15333
                    $mesoValue = 13240
                    $orgValue = 16220
                }
                { $_ -lt (GetBuildVersion $ex16 "CU20") } {
                    $cuLevel = "CU19"
                    $cuReleaseDate = "12/15/2020"
                    $mesoValue = 13239
                    $orgValue = 16219
                }
                { $_ -lt (GetBuildVersion $ex16 "CU19") } {
                    $cuLevel = "CU18"
                    $cuReleaseDate = "09/15/2020"
                    $schemaValue = 15332
                    $mesoValue = 13238
                    $orgValue = 16218
                }
                { $_ -lt (GetBuildVersion $ex16 "CU18") } {
                    $cuLevel = "CU17"
                    $cuReleaseDate = "06/16/2020"
                    $mesoValue = 13237
                    $orgValue = 16217
                }
                { $_ -lt (GetBuildVersion $ex16 "CU17") } {
                    $cuLevel = "CU16"
                    $cuReleaseDate = "03/17/2020"
                }
                { $_ -lt (GetBuildVersion $ex16 "CU16") } {
                    $cuLevel = "CU15"
                    $cuReleaseDate = "12/17/2019"
                }
                { $_ -lt (GetBuildVersion $ex16 "CU15") } {
                    $cuLevel = "CU14"
                    $cuReleaseDate = "09/17/2019"
                }
                { $_ -lt (GetBuildVersion $ex16 "CU14") } {
                    $cuLevel = "CU13"
                    $cuReleaseDate = "06/18/2019"
                }
                { $_ -lt (GetBuildVersion $ex16 "CU13") } {
                    $cuLevel = "CU12"
                    $cuReleaseDate = "02/12/2019"
                    $mesoValue = 13236
                    $orgValue = 16215
                }
                { $_ -lt (GetBuildVersion $ex16 "CU12") } {
                    $cuLevel = "CU11"
                    $cuReleaseDate = "10/16/2018"
                    $orgValue = 16214
                }
                { $_ -lt (GetBuildVersion $ex16 "CU11") } {
                    $cuLevel = "CU10"
                    $cuReleaseDate = "06/19/2018"
                    $orgValue = 16213
                }
                { $_ -lt (GetBuildVersion $ex16 "CU10") } {
                    $cuLevel = "CU9"
                    $cuReleaseDate = "03/20/2018"
                }
                { $_ -lt (GetBuildVersion $ex16 "CU9") } {
                    $cuLevel = "CU8"
                    $cuReleaseDate = "12/19/2017"
                }
                { $_ -lt (GetBuildVersion $ex16 "CU8") } {
                    $cuLevel = "CU7"
                    $cuReleaseDate = "09/16/2017"
                }
                { $_ -lt (GetBuildVersion $ex16 "CU7") } {
                    $cuLevel = "CU6"
                    $cuReleaseDate = "06/24/2017"
                    $schemaValue = 15330
                }
                { $_ -lt (GetBuildVersion $ex16 "CU6") } {
                    $cuLevel = "CU5"
                    $cuReleaseDate = "03/21/2017"
                    $schemaValue = 15326
                }
                { $_ -lt (GetBuildVersion $ex16 "CU5") } {
                    $cuLevel = "CU4"
                    $cuReleaseDate = "12/13/2016"
                }
                { $_ -lt (GetBuildVersion $ex16 "CU4") } {
                    $cuLevel = "CU3"
                    $cuReleaseDate = "09/20/2016"
                    $orgValue = 16212
                }
                { $_ -lt (GetBuildVersion $ex16 "CU3") } {
                    $cuLevel = "CU2"
                    $cuReleaseDate = "06/21/2016"
                    $schemaValue = 15325
                }
                { $_ -lt (GetBuildVersion $ex16 "CU2") } {
                    $cuLevel = "CU1"
                    $cuReleaseDate = "03/15/2016"
                    $schemaValue = 15323
                    $orgValue = 16211
                }
            }
        } elseif ($exchangeVersion.Major -eq 15 -and $exchangeVersion.Minor -eq 0) {
            Write-Verbose "Exchange 2013 is detected"
            $exchangeMajorVersion = "Exchange2013"
            $extendedSupportDate = "04/11/2023"
            $friendlyName = "Exchange 2013"

            #Latest Version AD Settings
            $schemaValue = 15312
            $mesoValue = 13237
            $orgValue = 16133

            switch ($exchangeVersion) {
                { $_ -ge (GetBuildVersion $ex13 "CU23") } {
                    $cuLevel = "CU23"
                    $cuReleaseDate = "06/18/2019"
                    $supportedBuildNumber = $true
                }
                (GetBuildVersion $ex13 "CU23" -SU "May22SU") { $mesoValue = 13238 }
                { $_ -lt (GetBuildVersion $ex13 "CU23") } {
                    $cuLevel = "CU22"
                    $cuReleaseDate = "02/12/2019"
                    $mesoValue = 13236
                    $orgValue = 16131
                    $supportedBuildNumber = $false
                }
                { $_ -lt (GetBuildVersion $ex13 "CU22") } {
                    $cuLevel = "CU21"
                    $cuReleaseDate = "06/19/2018"
                    $orgValue = 16130
                }
                { $_ -lt (GetBuildVersion $ex13 "CU21") } {
                    $cuLevel = "CU20"
                    $cuReleaseDate = "03/20/2018"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU20") } {
                    $cuLevel = "CU19"
                    $cuReleaseDate = "12/19/2017"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU19") } {
                    $cuLevel = "CU18"
                    $cuReleaseDate = "09/16/2017"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU18") } {
                    $cuLevel = "CU17"
                    $cuReleaseDate = "06/24/2017"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU17") } {
                    $cuLevel = "CU16"
                    $cuReleaseDate = "03/21/2017"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU16") } {
                    $cuLevel = "CU15"
                    $cuReleaseDate = "12/13/2016"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU15") } {
                    $cuLevel = "CU14"
                    $cuReleaseDate = "09/20/2016"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU14") } {
                    $cuLevel = "CU13"
                    $cuReleaseDate = "06/21/2016"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU13") } {
                    $cuLevel = "CU12"
                    $cuReleaseDate = "03/15/2016"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU12") } {
                    $cuLevel = "CU11"
                    $cuReleaseDate = "12/15/2015"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU11") } {
                    $cuLevel = "CU10"
                    $cuReleaseDate = "09/15/2015"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU10") } {
                    $cuLevel = "CU9"
                    $cuReleaseDate = "06/17/2015"
                    $orgValue = 15965
                }
                { $_ -lt (GetBuildVersion $ex13 "CU9") } {
                    $cuLevel = "CU8"
                    $cuReleaseDate = "03/17/2015"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU8") } {
                    $cuLevel = "CU7"
                    $cuReleaseDate = "12/09/2014"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU7") } {
                    $cuLevel = "CU6"
                    $cuReleaseDate = "08/26/2014"
                    $schemaValue = 15303
                }
                { $_ -lt (GetBuildVersion $ex13 "CU6") } {
                    $cuLevel = "CU5"
                    $cuReleaseDate = "05/27/2014"
                    $schemaValue = 15300
                    $orgValue = 15870
                }
                { $_ -lt (GetBuildVersion $ex13 "CU5") } {
                    $cuLevel = "CU4"
                    $cuReleaseDate = "02/25/2014"
                    $schemaValue = 15292
                    $orgValue = 15844
                }
                { $_ -lt (GetBuildVersion $ex13 "CU4") } {
                    $cuLevel = "CU3"
                    $cuReleaseDate = "11/25/2013"
                    $schemaValue = 15283
                    $orgValue = 15763
                }
                { $_ -lt (GetBuildVersion $ex13 "CU3") } {
                    $cuLevel = "CU2"
                    $cuReleaseDate = "07/09/2013"
                    $schemaValue = 15281
                    $orgValue = 15688
                }
                { $_ -lt (GetBuildVersion $ex13 "CU2") } {
                    $cuLevel = "CU1"
                    $cuReleaseDate = "04/02/2013"
                    $schemaValue = 15254
                    $orgValue = 15614
                }
            }
        } else {
            Write-Verbose "Unknown version of Exchange is detected."
        }

        # Now get the SU Name
        if ([string]::IsNullOrEmpty($exchangeMajorVersion) -or
            [string]::IsNullOrEmpty($cuLevel)) {
            Write-Verbose "Can't lookup when keys aren't set"
            return
        }

        $currentSUInfo = $exchangeBuildDictionary[$exchangeMajorVersion][$cuLevel].SU
        $compareValue = $exchangeVersion.ToString()
        if ($null -ne $currentSUInfo -and
            $currentSUInfo.ContainsValue($compareValue)) {
            foreach ($key in $currentSUInfo.Keys) {
                if ($compareValue -eq $currentSUInfo[$key]) {
                    $suName = $key
                }
            }
        }
    }
    end {

        if ($PSCmdlet.ParameterSetName -eq "FindSUBuilds") {
            Write-Verbose "Return nothing here, results were already returned on the pipeline"
            return
        }

        $friendlyName = "$friendlyName $cuLevel $suName".Trim()
        Write-Verbose "Determined Build Version $friendlyName"
        return [PSCustomObject]@{
            MajorVersion        = $exchangeMajorVersion
            FriendlyName        = $friendlyName
            BuildVersion        = $exchangeVersion
            CU                  = $cuLevel
            ReleaseDate         = if (-not([System.String]::IsNullOrEmpty($cuReleaseDate))) { ([System.Convert]::ToDateTime([DateTime]$cuReleaseDate, [System.Globalization.DateTimeFormatInfo]::InvariantInfo)) } else { $null }
            ExtendedSupportDate = if (-not([System.String]::IsNullOrEmpty($extendedSupportDate))) { ([System.Convert]::ToDateTime([DateTime]$extendedSupportDate, [System.Globalization.DateTimeFormatInfo]::InvariantInfo)) } else { $null }
            Supported           = $supportedBuildNumber
            LatestSU            = $latestSUBuild
            ADLevel             = [PSCustomObject]@{
                SchemaValue = $schemaValue
                MESOValue   = $mesoValue
                OrgValue    = $orgValue
            }
        }
    }
}

function GetExchangeBuildDictionary {

    function NewCUAndSUObject {
        param(
            [string]$CUBuildNumber,
            [Hashtable]$SUBuildNumber
        )
        return @{
            "CU" = $CUBuildNumber
            "SU" = $SUBuildNumber
        }
    }

    @{
        "Exchange2013" = @{
            "CU1"  = (NewCUAndSUObject "15.0.620.29")
            "CU2"  = (NewCUAndSUObject "15.0.712.24")
            "CU3"  = (NewCUAndSUObject "15.0.775.38")
            "CU4"  = (NewCUAndSUObject "15.0.847.32")
            "CU5"  = (NewCUAndSUObject "15.0.913.22")
            "CU6"  = (NewCUAndSUObject "15.0.995.29")
            "CU7"  = (NewCUAndSUObject "15.0.1044.25")
            "CU8"  = (NewCUAndSUObject "15.0.1076.9")
            "CU9"  = (NewCUAndSUObject "15.0.1104.5")
            "CU10" = (NewCUAndSUObject "15.0.1130.7")
            "CU11" = (NewCUAndSUObject "15.0.1156.6")
            "CU12" = (NewCUAndSUObject "15.0.1178.4")
            "CU13" = (NewCUAndSUObject "15.0.1210.3")
            "CU14" = (NewCUAndSUObject "15.0.1236.3")
            "CU15" = (NewCUAndSUObject "15.0.1263.5")
            "CU16" = (NewCUAndSUObject "15.0.1293.2")
            "CU17" = (NewCUAndSUObject "15.0.1320.4")
            "CU18" = (NewCUAndSUObject "15.0.1347.2" @{
                    "Mar18SU" = "15.0.1347.5"
                })
            "CU19" = (NewCUAndSUObject "15.0.1365.1" @{
                    "Mar18SU" = "15.0.1365.3"
                    "May18SU" = "15.0.1365.7"
                })
            "CU20" = (NewCUAndSUObject "15.0.1367.3" @{
                    "May18SU" = "15.0.1367.6"
                    "Aug18SU" = "15.0.1367.9"
                })
            "CU21" = (NewCUAndSUObject "15.0.1395.4" @{
                    "Aug18SU" = "15.0.1395.7"
                    "Oct18SU" = "15.0.1395.8"
                    "Jan19SU" = "15.0.1395.10"
                    "Mar21SU" = "15.0.1395.12"
                })
            "CU22" = (NewCUAndSUObject "15.0.1473.3" @{
                    "Feb19SU" = "15.0.1473.3"
                    "Apr19SU" = "15.0.1473.4"
                    "Jun19SU" = "15.0.1473.5"
                    "Mar21SU" = "15.0.1473.6"
                })
            "CU23" = (NewCUAndSUObject "15.0.1497.2" @{
                    "Jul19SU" = "15.0.1497.3"
                    "Nov19SU" = "15.0.1497.4"
                    "Feb20SU" = "15.0.1497.6"
                    "Oct20SU" = "15.0.1497.7"
                    "Nov20SU" = "15.0.1497.8"
                    "Dec20SU" = "15.0.1497.10"
                    "Mar21SU" = "15.0.1497.12"
                    "Apr21SU" = "15.0.1497.15"
                    "May21SU" = "15.0.1497.18"
                    "Jul21SU" = "15.0.1497.23"
                    "Oct21SU" = "15.0.1497.24"
                    "Nov21SU" = "15.0.1497.26"
                    "Jan22SU" = "15.0.1497.28"
                    "Mar22SU" = "15.0.1497.33"
                    "May22SU" = "15.0.1497.36"
                    "Aug22SU" = "15.0.1497.40"
                    "Oct22SU" = "15.0.1497.42"
                    "Nov22SU" = "15.0.1497.44"
                    "Jan23SU" = "15.0.1497.45"
                    "Feb23SU" = "15.0.1497.47"
                    "Mar23SU" = "15.0.1497.48"
                })
        }
        "Exchange2016" = @{
            "CU1"  = (NewCUAndSUObject "15.1.396.30")
            "CU2"  = (NewCUAndSUObject "15.1.466.34")
            "CU3"  = (NewCUAndSUObject "15.1.544.27")
            "CU4"  = (NewCUAndSUObject "15.1.669.32")
            "CU5"  = (NewCUAndSUObject "15.1.845.34")
            "CU6"  = (NewCUAndSUObject "15.1.1034.26")
            "CU7"  = (NewCUAndSUObject "15.1.1261.35" @{
                    "Mar18SU" = "15.1.1261.39"
                })
            "CU8"  = (NewCUAndSUObject "15.1.1415.2" @{
                    "Mar18SU" = "15.1.1415.4"
                    "May18SU" = "15.1.1415.7"
                    "Mar21SU" = "15.1.1415.8"
                })
            "CU9"  = (NewCUAndSUObject "15.1.1466.3" @{
                    "May18SU" = "15.1.1466.8"
                    "Aug18SU" = "15.1.1466.9"
                    "Mar21SU" = "15.1.1466.13"
                })
            "CU10" = (NewCUAndSUObject "15.1.1531.3" @{
                    "Aug18SU" = "15.1.1531.6"
                    "Oct18SU" = "15.1.1531.8"
                    "Jan19SU" = "15.1.1531.10"
                    "Mar21SU" = "15.1.1531.12"
                })
            "CU11" = (NewCUAndSUObject "15.1.1591.10" @{
                    "Dec18SU" = "15.1.1591.11"
                    "Jan19SU" = "15.1.1591.13"
                    "Apr19SU" = "15.1.1591.16"
                    "Jun19SU" = "15.1.1591.17"
                    "Mar21SU" = "15.1.1591.18"
                })
            "CU12" = (NewCUAndSUObject "15.1.1713.5" @{
                    "Feb19SU" = "15.1.1713.5"
                    "Apr19SU" = "15.1.1713.6"
                    "Jun19SU" = "15.1.1713.7"
                    "Jul19SU" = "15.1.1713.8"
                    "Sep19SU" = "15.1.1713.9"
                    "Mar21SU" = "15.1.1713.10"
                })
            "CU13" = (NewCUAndSUObject "15.1.1779.2" @{
                    "Jul19SU" = "15.1.1779.4"
                    "Sep19SU" = "15.1.1779.5"
                    "Nov19SU" = "15.1.1779.7"
                    "Mar21SU" = "15.1.1779.8"
                })
            "CU14" = (NewCUAndSUObject "15.1.1847.3" @{
                    "Nov19SU" = "15.1.1847.5"
                    "Feb20SU" = "15.1.1847.7"
                    "Mar20SU" = "15.1.1847.10"
                    "Mar21SU" = "15.1.1847.12"
                })
            "CU15" = (NewCUAndSUObject "15.1.1913.5" @{
                    "Feb20SU" = "15.1.1913.7"
                    "Mar20SU" = "15.1.1913.10"
                    "Mar21SU" = "15.1.1913.12"
                })
            "CU16" = (NewCUAndSUObject "15.1.1979.3" @{
                    "Sep20SU" = "15.1.1979.6"
                    "Mar21SU" = "15.1.1979.8"
                })
            "CU17" = (NewCUAndSUObject "15.1.2044.4" @{
                    "Sep20SU" = "15.1.2044.6"
                    "Oct20SU" = "15.1.2044.7"
                    "Nov20SU" = "15.1.2044.8"
                    "Dec20SU" = "15.1.2044.12"
                    "Mar21SU" = "15.1.2044.13"
                })
            "CU18" = (NewCUAndSUObject "15.1.2106.2" @{
                    "Oct20SU" = "15.1.2106.3"
                    "Nov20SU" = "15.1.2106.4"
                    "Dec20SU" = "15.1.2106.6"
                    "Feb21SU" = "15.1.2106.8"
                    "Mar21SU" = "15.1.2106.13"
                })
            "CU19" = (NewCUAndSUObject "15.1.2176.2" @{
                    "Feb21SU" = "15.1.2176.4"
                    "Mar21SU" = "15.1.2176.9"
                    "Apr21SU" = "15.1.2176.12"
                    "May21SU" = "15.1.2176.14"
                })
            "CU20" = (NewCUAndSUObject "15.1.2242.4" @{
                    "Apr21SU" = "15.1.2242.8"
                    "May21SU" = "15.1.2242.10"
                    "Jul21SU" = "15.1.2242.12"
                })
            "CU21" = (NewCUAndSUObject "15.1.2308.8" @{
                    "Jul21SU" = "15.1.2308.14"
                    "Oct21SU" = "15.1.2308.15"
                    "Nov21SU" = "15.1.2308.20"
                    "Jan22SU" = "15.1.2308.21"
                    "Mar22SU" = "15.1.2308.27"
                })
            "CU22" = (NewCUAndSUObject "15.1.2375.7" @{
                    "Oct21SU" = "15.1.2375.12"
                    "Nov21SU" = "15.1.2375.17"
                    "Jan22SU" = "15.1.2375.18"
                    "Mar22SU" = "15.1.2375.24"
                    "May22SU" = "15.1.2375.28"
                    "Aug22SU" = "15.1.2375.31"
                    "Oct22SU" = "15.1.2375.32"
                    "Nov22SU" = "15.1.2375.37"
                })
            "CU23" = (NewCUAndSUObject "15.1.2507.6" @{
                    "May22SU"   = "15.1.2507.9"
                    "Aug22SU"   = "15.1.2507.12"
                    "Oct22SU"   = "15.1.2507.13"
                    "Nov22SU"   = "15.1.2507.16"
                    "Jan23SU"   = "15.1.2507.17"
                    "Feb23SU"   = "15.1.2507.21"
                    "Mar23SU"   = "15.1.2507.23"
                    "Jun23SU"   = "15.1.2507.27"
                    "Aug23SU"   = "15.1.2507.31"
                    "Aug23SUv2" = "15.1.2507.32"
                    "Oct23SU"   = "15.1.2507.34"
                    "Nov23SU"   = "15.1.2507.35"
                    "Mar24SU"   = "15.1.2507.37"
                })
        }
        "Exchange2019" = @{
            "CU1"  = (NewCUAndSUObject "15.2.330.5" @{
                    "Feb19SU" = "15.2.330.5"
                    "Apr19SU" = "15.2.330.7"
                    "Jun19SU" = "15.2.330.8"
                    "Jul19SU" = "15.2.330.9"
                    "Sep19SU" = "15.2.330.10"
                    "Mar21SU" = "15.2.330.11"
                })
            "CU2"  = (NewCUAndSUObject "15.2.397.3" @{
                    "Jul19SU" = "15.2.397.5"
                    "Sep19SU" = "15.2.397.6"
                    "Nov19SU" = "15.2.397.9"
                    "Mar21SU" = "15.2.397.11"
                })
            "CU3"  = (NewCUAndSUObject "15.2.464.5" @{
                    "Nov19SU" = "15.2.464.7"
                    "Feb20SU" = "15.2.464.11"
                    "Mar20SU" = "15.2.464.14"
                    "Mar21SU" = "15.2.464.15"
                })
            "CU4"  = (NewCUAndSUObject "15.2.529.5" @{
                    "Feb20SU" = "15.2.529.8"
                    "Mar20SU" = "15.2.529.11"
                    "Mar21SU" = "15.2.529.13"
                })
            "CU5"  = (NewCUAndSUObject "15.2.595.3" @{
                    "Sep20SU" = "15.2.595.6"
                    "Mar21SU" = "15.2.595.8"
                })
            "CU6"  = (NewCUAndSUObject "15.2.659.4" @{
                    "Sep20SU" = "15.2.659.6"
                    "Oct20SU" = "15.2.659.7"
                    "Nov20SU" = "15.2.659.8"
                    "Dec20SU" = "15.2.659.11"
                    "Mar21SU" = "15.2.659.12"
                })
            "CU7"  = (NewCUAndSUObject "15.2.721.2" @{
                    "Oct20SU" = "15.2.721.3"
                    "Nov20SU" = "15.2.721.4"
                    "Dec20SU" = "15.2.721.6"
                    "Feb21SU" = "15.2.721.8"
                    "Mar21SU" = "15.2.721.13"
                })
            "CU8"  = (NewCUAndSUObject "15.2.792.3" @{
                    "Feb21SU" = "15.2.792.5"
                    "Mar21SU" = "15.2.792.10"
                    "Apr21SU" = "15.2.792.13"
                    "May21SU" = "15.2.792.15"
                })
            "CU9"  = (NewCUAndSUObject "15.2.858.5" @{
                    "Apr21SU" = "15.2.858.10"
                    "May21SU" = "15.2.858.12"
                    "Jul21SU" = "15.2.858.15"
                })
            "CU10" = (NewCUAndSUObject "15.2.922.7" @{
                    "Jul21SU" = "15.2.922.13"
                    "Oct21SU" = "15.2.922.14"
                    "Nov21SU" = "15.2.922.19"
                    "Jan22SU" = "15.2.922.20"
                    "Mar22SU" = "15.2.922.27"
                })
            "CU11" = (NewCUAndSUObject "15.2.986.5" @{
                    "Oct21SU" = "15.2.986.9"
                    "Nov21SU" = "15.2.986.14"
                    "Jan22SU" = "15.2.986.15"
                    "Mar22SU" = "15.2.986.22"
                    "May22SU" = "15.2.986.26"
                    "Aug22SU" = "15.2.986.29"
                    "Oct22SU" = "15.2.986.30"
                    "Nov22SU" = "15.2.986.36"
                    "Jan23SU" = "15.2.986.37"
                    "Feb23SU" = "15.2.986.41"
                    "Mar23SU" = "15.2.986.42"
                })
            "CU12" = (NewCUAndSUObject "15.2.1118.7" @{
                    "May22SU"   = "15.2.1118.9"
                    "Aug22SU"   = "15.2.1118.12"
                    "Oct22SU"   = "15.2.1118.15"
                    "Nov22SU"   = "15.2.1118.20"
                    "Jan23SU"   = "15.2.1118.21"
                    "Feb23SU"   = "15.2.1118.25"
                    "Mar23SU"   = "15.2.1118.26"
                    "Jun23SU"   = "15.2.1118.30"
                    "Aug23SU"   = "15.2.1118.36"
                    "Aug23SUv2" = "15.2.1118.37"
                    "Oct23SU"   = "15.2.1118.39"
                    "Nov23SU"   = "15.2.1118.40"
                })
            "CU13" = (NewCUAndSUObject "15.2.1258.12" @{
                    "Jun23SU"   = "15.2.1258.16"
                    "Aug23SU"   = "15.2.1258.23"
                    "Aug23SUv2" = "15.2.1258.25"
                    "Oct23SU"   = "15.2.1258.27"
                    "Nov23SU"   = "15.2.1258.28"
                    "Mar24SU"   = "15.2.1258.32"
                })
            "CU14" = (NewCUAndSUObject "15.2.1544.4" @{
                    "Mar24SU" = "15.2.1544.9"
                })
        }
    }
}

# Must be outside function to use it as a validate script
function GetValidatePossibleParameters {
    $exchangeBuildDictionary = GetExchangeBuildDictionary
    $suNames = New-Object 'System.Collections.Generic.HashSet[string]'
    $cuNames = New-Object 'System.Collections.Generic.HashSet[string]'
    $versionNames = New-Object 'System.Collections.Generic.HashSet[string]'

    foreach ($exchangeKey in $exchangeBuildDictionary.Keys) {
        [void]$versionNames.Add($exchangeKey)
        foreach ($cuKey in $exchangeBuildDictionary[$exchangeKey].Keys) {
            [void]$cuNames.Add($cuKey)
            if ($null -eq $exchangeBuildDictionary[$exchangeKey][$cuKey].SU) { continue }
            foreach ($suKey in $exchangeBuildDictionary[$exchangeKey][$cuKey].SU.Keys) {
                [void]$suNames.Add($suKey)
            }
        }
    }
    return [PSCustomObject]@{
        Version = $versionNames
        CU      = $cuNames
        SU      = $suNames
    }
}

function ValidateSUParameter {
    param($name)

    $possibleParameters = GetValidatePossibleParameters
    $possibleParameters.SU.Contains($Name)
}

function ValidateCUParameter {
    param($Name)

    $possibleParameters = GetValidatePossibleParameters
    $possibleParameters.CU.Contains($Name)
}

function ValidateVersionParameter {
    param($Name)

    $possibleParameters = GetValidatePossibleParameters
    $possibleParameters.Version.Contains($Name)
}
#TODO: Create Pester Testing on this
# Used to get the Exchange Version information and what roles are set on the server.
function Get-ExchangeBasicServerObject {
    param(
        [Parameter(Mandatory = $true)][string]$ServerName,
        [Parameter(Mandatory = $false)][bool]$AddGetServerProperty = $false
    )
    Write-Verbose("Function Enter: $($MyInvocation.MyCommand)")
    Write-Verbose("Passed: [string]ServerName: {0}" -f $ServerName)
    try {
        $getExchangeServer = Get-ExchangeServer $ServerName -Status -ErrorAction Stop
    } catch {
        Write-Host "Failed to detect server $ServerName as an Exchange Server" -ForegroundColor "Red"
        Invoke-CatchActions
        return $null
    }

    $exchAdminDisplayVersion = $getExchangeServer.AdminDisplayVersion
    $exchServerRole = $getExchangeServer.ServerRole
    Write-Verbose("AdminDisplayVersion: {0} | ServerRole: {1}" -f $exchAdminDisplayVersion.ToString(), $exchServerRole.ToString())
    $buildVersionInformation = Get-ExchangeBuildVersionInformation $exchAdminDisplayVersion

    if ($buildVersionInformation.BuildVersion.Major -eq 15) {
        if ($buildVersionInformation.BuildVersion.Minor -eq 0) {
            $exchVersion = 15
        } elseif ($buildVersionInformation.BuildVersion.Minor -eq 1) {
            $exchVersion = 16
        } else {
            $exchVersion = 19
        }
    }

    $mailbox = $exchServerRole -like "*Mailbox*"
    $dagName = [string]::Empty
    $exchangeServer = $null

    if ($mailbox) {
        $getMailboxServer = Get-MailboxServer $ServerName

        if (-not([string]::IsNullOrEmpty($getMailboxServer.DatabaseAvailabilityGroup))) {
            $dagName = $getMailboxServer.DatabaseAvailabilityGroup.ToString()
        }
    }

    if ($AddGetServerProperty) {
        $exchangeServer = $getExchangeServer
    }

    $exchServerObject = [PSCustomObject]@{
        ServerName     = $getExchangeServer.Name.ToUpper()
        Mailbox        = $mailbox
        MailboxOnly    = $exchServerRole -eq "Mailbox"
        Hub            = $exchVersion -ge 15 -and (-not ($exchServerRole -eq "ClientAccess"))
        CAS            = $exchVersion -ge 16 -or $exchServerRole -like "*ClientAccess*"
        CASOnly        = $exchServerRole -eq "ClientAccess"
        Edge           = $exchServerRole -eq "Edge"
        Version        = $exchVersion
        DAGMember      = (-not ([string]::IsNullOrEmpty($dagName)))
        DAGName        = $dagName
        ExchangeServer = $exchangeServer
    }

    Write-Verbose("Mailbox: {0} | CAS: {1} | Hub: {2} | CASOnly: {3} | MailboxOnly: {4} | Edge: {5} | DAGMember {6} | Version: {7} | AnyTransportSwitchesEnabled: {8} | DAGName: {9}" -f $exchServerObject.Mailbox,
        $exchServerObject.CAS,
        $exchServerObject.Hub,
        $exchServerObject.CASOnly,
        $exchServerObject.MailboxOnly,
        $exchServerObject.Edge,
        $exchServerObject.DAGMember,
        $exchServerObject.Version,
        $Script:AnyTransportSwitchesEnabled,
        $exchServerObject.DAGName
    )

    return $exchServerObject
}


# Injects Verbose and Debug Preferences and other passed variables into the script block
# It will also inject any additional script blocks into the main script block.
# This allows for an Invoke-Command to work as intended if multiple functions/script blocks are required.
function Add-ScriptBlockInjection {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [ScriptBlock]$PrimaryScriptBlock,

        [string[]]$IncludeUsingParameter,

        [ScriptBlock[]]$IncludeScriptBlock,

        [ScriptBlock]
        $CatchActionFunction
    )
    process {
        try {
            # In Remote Execution if you want Write-Verbose to work or add in additional
            # Script Blocks to your code to be executed, like a custom Write-Verbose, you need to inject it into the script block
            # that is passed to Invoke-Command.
            Write-Verbose "Calling: $($MyInvocation.MyCommand)"
            $scriptBlockInjectLines = @()
            $scriptBlockFinalized = [string]::Empty
            $adjustedScriptBlock = $PrimaryScriptBlock
            $injectedLinesHandledInBeginBlock = $false
            $adjustInject = $false

            if ($null -ne $IncludeUsingParameter) {
                $lines = @()
                $lines += 'if ($PSSenderInfo) {'
                $IncludeUsingParameter | ForEach-Object {
                    $lines += '$name=$Using:name'.Replace("name", "$_")
                }
                $lines += "}" + [System.Environment]::NewLine
                $usingLines = $lines -join [System.Environment]::NewLine
            } else {
                $usingLines = [System.Environment]::NewLine
            }

            if ($null -ne $IncludeScriptBlock) {
                $lines = @()
                $IncludeScriptBlock | ForEach-Object {
                    $lines += "Function $($_.Ast.Name) { $([System.Environment]::NewLine)"
                    $lines += "$($_.ToString().Trim()) $([System.Environment]::NewLine) } $([System.Environment]::NewLine)"
                }
                $scriptBlockIncludeLines = $lines -join [System.Environment]::NewLine
            } else {
                $scriptBlockIncludeLines = [System.Environment]::NewLine
            }

            # There are a few different ways to create a script block
            # [ScriptBlock]::Create(string) and ${Function:Write-Verbose}
            # each one ends up adding in the ParamBlock at different locations
            # You need to add in the injected code after the params, if that is the only thing that is passed
            # If you provide a script block that contains a begin or a process section,
            # you need to add the injected code into the begin block.
            # Here you need to find the ParamBlock and add it to the inject lines to be at the top of the script block.
            # Then you need to recreate the adjustedScriptBlock to be where the ParamBlock ended.

            if ($null -ne $PrimaryScriptBlock.Ast.ParamBlock) {
                Write-Verbose "Ast ParamBlock detected"
                $adjustLocation = $PrimaryScriptBlock.Ast
            } elseif ($null -ne $PrimaryScriptBlock.Ast.Body.ParamBlock) {
                Write-Verbose "Ast Body ParamBlock detected"
                $adjustLocation = $PrimaryScriptBlock.Ast.Body
            }

            $adjustInject = $null -ne $PrimaryScriptBlock.Ast.ParamBlock -or $null -ne $PrimaryScriptBlock.Ast.Body.ParamBlock

            if ($adjustInject) {
                $scriptBlockInjectLines += $adjustLocation.ParamBlock.ToString()
                $startIndex = $adjustLocation.ParamBlock.Extent.EndOffSet - $adjustLocation.Extent.StartOffset
                $adjustedScriptBlock = [ScriptBlock]::Create($PrimaryScriptBlock.ToString().Substring($startIndex))
            }

            # Inject the script blocks and using parameters in the begin block when required.
            if ($null -ne $adjustedScriptBlock.Ast.BeginBlock) {
                Write-Verbose "Ast BeginBlock detected"
                $replaceMatch = $adjustedScriptBlock.Ast.BeginBlock.Extent.ToString()
                $addString = [string]::Empty + [System.Environment]::NewLine
                $addString += {
                    if ($PSSenderInfo) {
                        $VerbosePreference = $Using:VerbosePreference
                        $DebugPreference = $Using:DebugPreference
                    }
                }
                $addString += [System.Environment]::NewLine + $usingLines + $scriptBlockIncludeLines
                $startIndex = $replaceMatch.IndexOf("{")
                #insert the adding context to one character after the begin curl bracket
                $replaceWith = $replaceMatch.Insert($startIndex + 1, $addString)
                $adjustedScriptBlock = [ScriptBlock]::Create($adjustedScriptBlock.ToString().Replace($replaceMatch, $replaceWith))
                $injectedLinesHandledInBeginBlock = $true
            } elseif ($null -ne $adjustedScriptBlock.Ast.ProcessBlock) {
                # Add in a begin block that contains all information that we are wanting.
                Write-Verbose "Ast Process Block detected"
                $addString = [string]::Empty + [System.Environment]::NewLine
                $addString += {
                    begin {
                        if ($PSScriptRoot) {
                            $VerbosePreference = $Using:VerbosePreference
                            $DebugPreference = $Using:DebugPreference
                        }
                    }
                }
                $endIndex = $addString.LastIndexOf("}") - 1
                $addString = $addString.insert($endIndex, [System.Environment]::NewLine + $usingLines + $scriptBlockIncludeLines + [System.Environment]::NewLine )
                $startIndex = $adjustedScriptBlock.Ast.ProcessBlock.Extent.StartOffset - 1
                $adjustedScriptBlock = [ScriptBlock]::Create($adjustedScriptBlock.ToString().Insert($startIndex, $addString))
                $injectedLinesHandledInBeginBlock = $true
            } else {
                Write-Verbose "No Begin or Process Blocks detected, normal injection"
                $scriptBlockInjectLines += {
                    if ($PSSenderInfo) {
                        $VerbosePreference = $Using:VerbosePreference
                        $DebugPreference = $Using:DebugPreference
                    }
                }
            }

            if (-not $injectedLinesHandledInBeginBlock) {
                $scriptBlockInjectLines += $usingLines + $scriptBlockIncludeLines + [System.Environment]::NewLine
            }

            # Combined the injected lines and the main script block together
            # then create a new script block from finalized result
            $scriptBlockInjectLines += $adjustedScriptBlock
            $scriptBlockInjectLines | ForEach-Object {
                $scriptBlockFinalized += $_.ToString() + [System.Environment]::NewLine
            }

            #Need to return a string type otherwise run into issues.
            return $scriptBlockFinalized
        } catch {
            Write-Verbose "Failed to add to the script block"
            Invoke-CatchActionError $CatchActionFunction
        }
    }
}

function New-PipelineObject {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Caller knows that this is an action')]
    [CmdletBinding()]
    param(
        [object]$Object,
        [string]$Type
    )
    process {
        return [PSCustomObject]@{
            Object = $Object
            Type   = $Type
        }
    }
}

function Invoke-PipelineHandler {
    [CmdletBinding()]
    param(
        [object[]]$Object
    )
    process {
        foreach ($instance in $Object) {
            if ($instance.Type -eq "Verbose") {
                Write-Verbose "$($instance.PSComputerName) - $($instance.Object)"
            } elseif ($instance.Type -eq "Host") {
                Write-Host $instance.Object
            } else {
                return $instance
            }
        }
    }
}

function New-VerbosePipelineObject {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Caller knows that this is an action')]
    [CmdletBinding()]
    param(
        [Parameter(Position = 1)]
        [string]$Message
    )
    process {
        New-PipelineObject $Message "Verbose"
    }
}

function Start-JobManager {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'I prefer Start here')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$ServersWithArguments,

        [Parameter(Mandatory = $true)]
        [ScriptBlock]$ScriptBlock,

        [string]$JobBatchName,

        [bool]$DisplayReceiveJob = $true,

        [bool]$NeedReturnData = $false,

        [ScriptBlock]$RemotePipelineHandler
    )
    <# It needs to be this way incase of different arguments being passed to different machines
        [array]ServersWithArguments
            [string]ServerName
            [object]ArgumentList #customized for your ScriptBlock
    #>

    function Wait-JobsCompleted {
        Write-Verbose "Calling Wait-JobsCompleted"
        [System.Diagnostics.Stopwatch]$timer = [System.Diagnostics.Stopwatch]::StartNew()
        # Data returned is a Hash Table that matches to the Server the Script Block ran against
        $returnData = @{}
        do {
            $completedJobs = Get-Job | Where-Object { $_.State -ne "Running" }
            if ($null -eq $completedJobs) {
                Start-Sleep 1
                continue
            }

            foreach ($job in $completedJobs) {
                $jobName = $job.Name
                Write-Verbose "Job $($job.Name) received. State: $($job.State) | HasMoreData: $($job.HasMoreData)"
                if ($NeedReturnData -eq $false -and $DisplayReceiveJob -eq $false -and $job.HasMoreData -eq $true) {
                    Write-Verbose "This job has data and you provided you didn't want to return it or display it."
                }
                $receiveJob = Receive-Job $job
                Remove-Job $job
                if ($null -eq $receiveJob) {
                    Write-Verbose "Job $jobName didn't have any receive job data"
                }

                # If more things are added to the pipeline than just the desired result (like custom Write-Verbose data to the pipeline)
                # The caller needs to handle this by having a custom ScriptBlock to process the data
                # Then return the desired result back
                if ($null -ne $RemotePipelineHandler -and $receiveJob) {
                    Write-Verbose "Starting to call RemotePipelineHandler"
                    $returnJobData = & $RemotePipelineHandler $receiveJob
                    Write-Verbose "Finished RemotePipelineHandler"
                    if ($null -ne $returnJobData) {
                        $returnData.Add($jobName, $returnJobData)
                    } else {
                        Write-Verbose "Nothing came back from the RemotePipelineHandler"
                    }
                } elseif ($NeedReturnData) {
                    $returnData.Add($jobName, $receiveJob)
                }
            }
        } while ($true -eq (Get-Job))
        $timer.Stop()
        Write-Verbose "Waiting for jobs to complete took $($timer.Elapsed.TotalSeconds) seconds"
        if ($NeedReturnData) {
            return $returnData
        }
        return $null
    }

    [System.Diagnostics.Stopwatch]$timerMain = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Verbose "Calling Start-JobManager"
    Write-Verbose "Passed: [bool]DisplayReceiveJob: $DisplayReceiveJob | [string]JobBatchName: $JobBatchName | [bool]NeedReturnData:$NeedReturnData"

    foreach ($serverObject in $ServersWithArguments) {
        $server = $serverObject.ServerName
        $argumentList = $serverObject.ArgumentList
        Write-Verbose "Starting job on server $server"
        Invoke-Command -ComputerName $server -ScriptBlock $ScriptBlock -ArgumentList $argumentList -AsJob -JobName $server | Out-Null
    }

    $data = Wait-JobsCompleted
    $timerMain.Stop()
    Write-Verbose "Exiting: Start-JobManager | Time in Start-JobManager: $($timerMain.Elapsed.TotalSeconds) seconds"
    if ($NeedReturnData) {
        return $data
    }
    return $null
}
#This function job is to write out the Data that is too large to pass into the main script block
#This is for mostly Exchange Related objects.
#To handle this, we export the data locally and copy the data over the correct server.
function Write-LargeDataObjectsOnMachine {

    Write-Verbose("Function Enter Write-LargeDataObjectsOnMachine")

    [array]$serverNames = $Script:ArgumentList.ServerObjects |
        ForEach-Object {
            return $_.ServerName
        }

    #Collect the Exchange Data that resides on their own machine.
    function Invoke-ExchangeResideDataCollectionWrite {
        param(
            [Parameter(Mandatory = $true, Position = 1)]
            [string]$SaveToLocation,

            [Parameter(Mandatory = $true, Position = 2)]
            [string]$InstallDirectory
        )

        $location = $SaveToLocation
        $exchBin = "{0}\Bin" -f $InstallDirectory
        $configFiles = Get-ChildItem $exchBin | Where-Object { $_.Name -like "*.config" }
        $copyTo = "{0}\Config" -f $location
        $configFiles | ForEach-Object { Copy-Item $_.VersionInfo.FileName $copyTo }

        $copyServerComponentStatesRegistryTo = "{0}\regServerComponentStates.TXT" -f $location
        reg query HKLM\SOFTWARE\Microsoft\ExchangeServer\v15\ServerComponentStates /s > $copyServerComponentStatesRegistryTo

        Get-Command ExSetup | ForEach-Object { $_.FileVersionInfo } > ("{0}\{1}_GCM.txt" -f $location, $env:COMPUTERNAME)

        #Exchange Web App Pools
        $windir = $env:windir
        $appCmd = "{0}\system32\inetSrv\appCmd.exe" -f $windir
        if (Test-Path $appCmd) {
            $appPools = &$appCmd list appPool
            $sites = &$appCmd list sites

            $exchangeAppPools = $appPools |
                ForEach-Object {
                    $startIndex = $_.IndexOf('"') + 1
                    $appPoolName = $_.Substring($startIndex,
                        ($_.Substring($startIndex).IndexOf('"')))
                    return $appPoolName
                } |
                Where-Object {
                    $_.StartsWith("MSExchange")
                }

            $sitesContent = @{}
            $sites |
                ForEach-Object {
                    $startIndex = $_.IndexOf('"') + 1
                    $siteName = $_.Substring($startIndex,
                        ($_.Substring($startIndex).IndexOf('"')))
                    $sitesContent.Add($siteName, (&$appCmd list site $siteName /text:*))
                }

            $webAppPoolsSaveRoot = "{0}\WebAppPools" -f $location
            $cacheConfigFileListLocation = @()
            $exchangeAppPools |
                ForEach-Object {
                    $config = &$appCmd list appPool $_ /text:CLRConfigFile
                    $allInfo = &$appCmd list appPool $_ /text:*

                    if (![string]::IsNullOrEmpty($config) -and
                        (Test-Path $config) -and
                        (!($cacheConfigFileListLocation.Contains($config.ToLower())))) {

                        $cacheConfigFileListLocation += $config.ToLower()
                        $saveConfigLocation = "{0}\{1}_{2}" -f $webAppPoolsSaveRoot, $env:COMPUTERNAME,
                        $config.Substring($config.LastIndexOf("\") + 1)
                        #Copy item to keep the date modify time
                        Copy-Item $config -Destination $saveConfigLocation
                    }
                    $saveAllInfoLocation = "{0}\{1}_{2}.txt" -f $webAppPoolsSaveRoot, $env:COMPUTERNAME, $_
                    $allInfo | Format-List * > $saveAllInfoLocation
                }

            $sitesContent.Keys |
                ForEach-Object {
                    $sitesContent[$_] > ("{0}\{1}_{2}_Site.config" -f $webAppPoolsSaveRoot, $env:COMPUTERNAME, ($_.Replace(" ", "")))
                    $slsResults = $sitesContent[$_] | Select-String applicationPool:, physicalPath:
                    $appPoolName = [string]::Empty
                    foreach ($matchInfo in $slsResults) {
                        $line = $matchInfo.Line

                        if ($line.Trim().StartsWith("applicationPool:")) {
                            $correctAppPoolSection = $false
                        }

                        if ($line.Trim().StartsWith("applicationPool:`"MSExchange")) {
                            $correctAppPoolSection = $true
                            $startIndex = $line.IndexOf('"') + 1
                            $appPoolName = $line.Substring($startIndex,
                                ($line.Substring($startIndex).IndexOf('"')))
                        }

                        if ($correctAppPoolSection -and
                            (!($line.Trim() -eq 'physicalPath:""')) -and
                            $line.Trim().StartsWith("physicalPath:")) {
                            $startIndex = $line.IndexOf('"') + 1
                            $path = $line.Substring($startIndex,
                                ($line.Substring($startIndex).IndexOf('"')))
                            $fullPath = "{0}\web.config" -f $path

                            if ((Test-Path $path) -and
                                (Test-Path $fullPath)) {
                                $saveFileName = "{0}\{1}_{2}_{3}_web.config" -f $webAppPoolsSaveRoot, $env:COMPUTERNAME, $appPoolName, ($_.Replace(" ", ""))
                                #Use Copy-Item to keep date modified
                                Copy-Item $fullPath -Destination $saveFileName
                                $bakFullPath = "$fullPath.bak"

                                if (Test-Path $bakFullPath) {
                                    Copy-Item $bakFullPath -Destination ("$saveFileName.bak")
                                }
                            }
                        }
                    }
                }

            $machineConfig = [System.Runtime.InteropServices.RuntimeEnvironment]::SystemConfigurationFile

            if (Test-Path $machineConfig) {
                Copy-Item $machineConfig -Destination ("{0}\{1}_machine.config" -f $webAppPoolsSaveRoot, $env:COMPUTERNAME)
            }

            $siteConfigs = @{}
            # always try to get the hardcoded default
            $siteConfigs.Add("applicationHost.config", "$($env:WINDIR)\System32\inetSrv\config\applicationHost.config")

            try {
                # default location normally your applicationHost.config
                try {
                    $defaultLocation = Get-WebConfigFile

                    if (-not $siteConfigs.ContainsKey($defaultLocation.Name)) {
                        $siteConfigs.Add($defaultLocation.Name, $defaultLocation.FullName)
                    }
                } catch {
                    Write-Verbose "Failed to get default web config file path. $_"
                }

                $sitesContent.Keys |
                    ForEach-Object {
                        try {
                            $name = $_
                            $siteWebFileConfig = Get-WebConfigFile "IIS:\Sites\$($name)"

                            $keyName = if ($siteWebFileConfig.Name -eq "web.config") { "$name`_web.config" } else { $siteWebFileConfig.Name }

                            if (-not $siteConfigs.ContainsKey($keyName)) {
                                $siteConfigs.Add($keyName, $siteWebFileConfig.FullName)
                            }
                        } catch {
                            Write-Verbose "Failed to get web config for $name. $_"
                        }
                    }
            } catch {
                Write-Verbose "Failed to get the web config file for the sites. $_"
                # remote context, cant call catch actions
            } finally {
                if ($null -ne $siteConfigs -and
                    $siteConfigs.Count -gt 0) {
                    $siteConfigs.Keys |
                        ForEach-Object {
                            if ((Test-Path $siteConfigs[$_])) {
                                Copy-Item $siteConfigs[$_] -Destination ("{0}\{1}_{2}" -f $webAppPoolsSaveRoot, $env:COMPUTERNAME, $_)
                            }
                        }
                }
            }

            # list the app pools ids
            $ids = & $appCmd list wp
            $fileName = ("{0}\{1}_Web_App_IDs.txt" -f $webAppPoolsSaveRoot, $env:COMPUTERNAME)

            if ($null -ne $ids) {
                $ids > $fileName
            } else {
                "No Data" > $fileName
            }
        }
    }

    #Write the Exchange Object Information locally first to then allow it to be copied over to the remote machine.
    #Exchange objects can be rather large preventing them to be passed within an Invoke-Command -ArgumentList
    #In order to get around this and to avoid going through a loop of doing an Invoke-Command per server per object,
    #Write the data out locally, copy that directory over to the remote location.
    function Write-ExchangeObjectDataLocal {
        param(
            [object]$ServerData,
            [string]$Location
        )
        $tempLocation = "{0}\{1}" -f $Location, $ServerData.ServerName
        Save-DataToFile -DataIn $ServerData.ExchangeServer -SaveToLocation ("{0}_ExchangeServer" -f $tempLocation)
        Save-DataToFile -DataIn $ServerData.HealthReport -SaveToLocation ("{0}_HealthReport" -f $tempLocation)
        Save-DataToFile -DataIn $ServerData.ServerComponentState -SaveToLocation ("{0}_ServerComponentState" -f $tempLocation)
        Save-DataToFile -DataIn $ServerData.ServerMonitoringOverride -SaveToLocation ("{0}_serverMonitoringOverride" -f $tempLocation)
        Save-DataToFile -DataIn $ServerData.ServerHealth -SaveToLocation ("{0}_ServerHealth" -f $tempLocation)

        if ($ServerData.Hub) {
            Save-DataToFile -DataIn $ServerData.TransportServerInfo -SaveToLocation ("{0}_TransportServer" -f $tempLocation)
            Save-DataToFile -DataIn $ServerData.ReceiveConnectors -SaveToLocation ("{0}_ReceiveConnectors" -f $tempLocation)
            Save-DataToFile -DataIn $ServerData.QueueData -SaveToLocation ("{0}_InstantQueueInfo" -f $tempLocation)
        }

        if ($ServerData.CAS) {
            Save-DataToFile -DataIn $ServerData.CAServerInfo -SaveToLocation ("{0}_ClientAccessServer" -f $tempLocation)
            Save-DataToFile -DataIn $ServerData.FrontendTransportServiceInfo -SaveToLocation ("{0}_FrontendTransportService" -f $tempLocation)
        }

        if ($ServerData.Mailbox) {
            Save-DataToFile -DataIn $ServerData.MailboxServerInfo -SaveToLocation ("{0}_MailboxServer" -f $tempLocation)
            Save-DataToFile -DataIn $ServerData.MailboxTransportServiceInfo -SaveToLocation ("{0}_MailboxTransportService" -f $tempLocation)
        }
    }

    function Write-DatabaseAvailabilityGroupDataLocal {
        param(
            [object]$DAGWriteInfo
        )
        $dagName = $DAGWriteInfo.DAGInfo.Name
        $serverName = $DAGWriteInfo.ServerName
        $rootSaveToLocation = $DAGWriteInfo.RootSaveToLocation
        $mailboxDatabaseSaveToLocation = "{0}\MailboxDatabase\" -f $rootSaveToLocation
        $copyStatusSaveToLocation = "{0}\MailboxDatabaseCopyStatus\" -f $rootSaveToLocation
        New-Item -ItemType Directory -Path @($mailboxDatabaseSaveToLocation, $copyStatusSaveToLocation) -Force | Out-Null
        Save-DataToFile -DataIn $DAGWriteInfo.DAGInfo -SaveToLocation ("{0}{1}_DatabaseAvailabilityGroup" -f $rootSaveToLocation, $dagName)
        Save-DataToFile -DataIn $DAGWriteInfo.DAGNetworkInfo -SaveToLocation ("{0}{1}_DatabaseAvailabilityGroupNetwork" -f $rootSaveToLocation, $dagName)
        Save-DataToFile -DataIn $DAGWriteInfo.MailboxDatabaseCopyStatusServer -SaveToLocation ("{0}{1}_MailboxDatabaseCopyStatus" -f $copyStatusSaveToLocation, $serverName)

        $DAGWriteInfo.MailboxDatabases |
            ForEach-Object {
                Save-DataToFile -DataIn $_ -SaveToLocation ("{0}{1}_MailboxDatabase" -f $mailboxDatabaseSaveToLocation, $_.Name)
            }

        $DAGWriteInfo.MailboxDatabaseCopyStatusPerDB.Keys |
            ForEach-Object {
                $data = $DAGWriteInfo.MailboxDatabaseCopyStatusPerDB[$_]
                Save-DataToFile -DataIn $data -SaveToLocation ("{0}{1}_MailboxDatabaseCopyStatus" -f $copyStatusSaveToLocation, $_)
            }
    }

    $dagNameGroup = $argumentList.ServerObjects |
        Group-Object DAGName |
        Where-Object { ![string]::IsNullOrEmpty($_.Name) }

    if ($DAGInformation -and
        !$Script:EdgeRoleDetected -and
        $null -ne $dagNameGroup -and
        $dagNameGroup.Count -ne 0) {

        $dagWriteInformation = @()
        $dagNameGroup |
            ForEach-Object {
                $dagName = $_.Name
                $getDAGInformation = Get-DAGInformation -DAGName $dagName
                $_.Group |
                    ForEach-Object {
                        $dagWriteInformation += [PSCustomObject]@{
                            ServerName                      = $_.ServerName
                            DAGInfo                         = $getDAGInformation.DAGInfo
                            DAGNetworkInfo                  = $getDAGInformation.DAGNetworkInfo
                            MailboxDatabaseCopyStatusServer = $getDAGInformation.MailboxDatabaseInfo[$_.ServerName].MailboxDatabaseCopyStatusServer
                            MailboxDatabases                = $getDAGInformation.MailboxDatabaseInfo[$_.ServerName].MailboxDatabases
                            MailboxDatabaseCopyStatusPerDB  = $getDAGInformation.MailboxDatabaseInfo[$_.ServerName].MailboxDatabaseCopyStatusPerDB
                        }
                    }
                }

        $localServerTempLocation = "{0}{1}\Exchange_DAG_Temp\" -f $Script:RootFilePath, $env:COMPUTERNAME
        $dagWriteInformation |
            ForEach-Object {
                $location = "{0}{1}" -f $Script:RootFilePath, $_.ServerName
                Write-Verbose("Location of the data should be at: $location")
                $remoteLocation = "\\{0}\{1}" -f $_.ServerName, $location.Replace(":", "$")
                Write-Verbose("Remote Copy Location: $remoteLocation")
                $rootTempLocation = "{0}{1}\{2}_Exchange_DAG_Information\" -f $localServerTempLocation, $_.ServerName, $_.DAGInfo.Name
                Write-Verbose("Local Root Temp Location: $rootTempLocation")
                New-Item -ItemType Directory -Path $rootTempLocation -Force | Out-Null
                $_ | Add-Member -MemberType NoteProperty -Name RootSaveToLocation -Value $rootTempLocation
                Write-DatabaseAvailabilityGroupDataLocal -DAGWriteInfo $_

                $zipCopyLocation = Compress-Folder -Folder $rootTempLocation -ReturnCompressedLocation $true
                try {
                    Copy-Item $zipCopyLocation $remoteLocation
                } catch {
                    Write-Verbose("Failed to copy data to $remoteLocation. This is likely due to file sharing permissions.")
                    Invoke-CatchActions
                }
            }
        #Remove the temp data location
        Remove-Item $localServerTempLocation -Force -Recurse
    }

    # Can not invoke CollectOverMetrics.ps1 script inside of a script block against a different machine.
    if ($CollectFailoverMetrics -and
        !$Script:LocalExchangeShell.RemoteShell -and
        !$Script:EdgeRoleDetected -and
        $null -ne $dagNameGroup -and
        $dagNameGroup.Count -ne 0) {

        $localServerTempLocation = "{0}{1}\Temp_Exchange_Failover_Reports" -f $Script:RootFilePath, $env:COMPUTERNAME
        $argumentList.ServerObjects |
            Group-Object DAGName |
            Where-Object { ![string]::IsNullOrEmpty($_.Name) } |
            ForEach-Object {
                $failed = $false
                $reportPath = "{0}\{1}_FailoverMetrics" -f $localServerTempLocation, $_.Name
                New-Item -ItemType Directory -Path $reportPath -Force | Out-Null

                try {
                    Write-Host "Attempting to run CollectOverMetrics.ps1 against $($_.Name)"
                    &"$Script:localExInstall\Scripts\CollectOverMetrics.ps1" -DatabaseAvailabilityGroup $_.Name `
                        -IncludeExtendedEvents `
                        -GenerateHtmlReport `
                        -ReportPath $reportPath
                } catch {
                    Write-Verbose("Failed to collect failover metrics")
                    Invoke-CatchActions
                    $failed = $true
                }

                if (!$failed) {
                    $zipCopyLocation = Compress-Folder -Folder $reportPath -ReturnCompressedLocation $true
                    $_.Group |
                        ForEach-Object {
                            $location = "{0}{1}" -f $Script:RootFilePath, $_.ServerName
                            Write-Verbose("Location of the data should be at: $location")
                            $remoteLocation = "\\{0}\{1}" -f $_.ServerName, $location.Replace(":", "$")
                            Write-Verbose("Remote Copy Location: $remoteLocation")

                            try {
                                Copy-Item $zipCopyLocation $remoteLocation
                            } catch {
                                Write-Verbose("Failed to copy data to $remoteLocation. This is likely due to file sharing permissions.")
                                Invoke-CatchActions
                            }
                        }
                    } else {
                        Write-Verbose("Not compressing or copying over this folder.")
                    }
                }

        Remove-Item $localServerTempLocation -Recurse -Force
    } elseif ($null -eq $dagNameGroup -or
        $dagNameGroup.Count -eq 0) {
        Write-Verbose("No DAGs were found. Didn't run CollectOverMetrics.ps1")
    } elseif ($Script:EdgeRoleDetected) {
        Write-Host "Unable to run CollectOverMetrics.ps1 script from an edge server" -ForegroundColor Yellow
    } elseif ($CollectFailoverMetrics) {
        Write-Host "Unable to run CollectOverMetrics.ps1 script from a remote shell session not on an Exchange Server or Tools box." -ForegroundColor Yellow
    }

    if ($ExchangeServerInformation) {

        #Create a list that contains all the information that we need to dump out locally then copy over to each respective server within "Exchange_Server_Data"
        $exchangeServerData = @()
        foreach ($server in $serverNames) {
            $basicServerObject = Get-ExchangeBasicServerObject -ServerName $server -AddGetServerProperty $true

            if ($basicServerObject.Hub) {
                $basicServerObject | Add-Member -MemberType NoteProperty -Name "TransportServerInfo" -Value (Get-TransportService $server)
                $basicServerObject | Add-Member -MemberType NoteProperty -Name "ReceiveConnectors" -Value (Get-ReceiveConnector -Server $server)
                $basicServerObject | Add-Member -MemberType NoteProperty -Name "QueueData" -Value (Get-Queue -Server $server)
            }

            if ($basicServerObject.CAS) {

                if ($basicServerObject.Version -ge 16) {
                    $getClientAccessService = Get-ClientAccessService $server -IncludeAlternateServiceAccountCredentialStatus
                } else {
                    $getClientAccessService = Get-ClientAccessServer $server -IncludeAlternateServiceAccountCredentialStatus
                }
                $basicServerObject | Add-Member -MemberType NoteProperty -Name "CAServerInfo" -Value $getClientAccessService
                $basicServerObject | Add-Member -MemberType NoteProperty -Name "FrontendTransportServiceInfo" -Value (Get-FrontendTransportService -Identity $server)
            }

            if ($basicServerObject.Mailbox) {
                $basicServerObject | Add-Member -MemberType NoteProperty -Name "MailboxServerInfo" -Value (Get-MailboxServer $server)
                $basicServerObject | Add-Member -MemberType NoteProperty -Name "MailboxTransportServiceInfo" -Value (Get-MailboxTransportService -Identity $server)
            }

            $basicServerObject | Add-Member -MemberType NoteProperty -Name "HealthReport" -Value (Get-HealthReport $server)
            $basicServerObject | Add-Member -MemberType NoteProperty -Name "ServerComponentState" -Value (Get-ServerComponentState $server)
            $basicServerObject | Add-Member -MemberType NoteProperty -Name "ServerMonitoringOverride" -Value (Get-ServerMonitoringOverride -Server $server -ErrorAction SilentlyContinue)
            $basicServerObject | Add-Member -MemberType NoteProperty -Name "ServerHealth" -Value (Get-ServerHealth $server)

            $exchangeServerData += $basicServerObject
        }

        #if single server or Exchange 2010 where invoke-command doesn't work
        if (!($serverNames.count -eq 1 -and
                $serverNames[0].ToUpper().Contains($env:COMPUTERNAME.ToUpper()))) {

            <#
            To pass an action to Start-JobManager, need to create objects like this.
                Where ArgumentList is the arguments for the ScriptBlock that we are running
            [array]
                [PSCustom]
                    [string]ServerName
                    [object]ArgumentList

            Need to do the following:
                Collect Exchange Install Directory Location
                Create directories where data is being stored with the upcoming requests
                Write out the Exchange Server Object Data and copy them over to the correct server
            #>

            # Set remote version action to be able to return objects on the pipeline to log and handle them.
            SetWriteRemoteVerboseAction "New-VerbosePipelineObject"
            $scriptBlockInjectParams = @{
                IncludeScriptBlock    = @(${Function:Write-Verbose}, ${Function:New-PipelineObject}, ${Function:New-VerbosePipelineObject})
                IncludeUsingParameter = "WriteRemoteVerboseDebugAction"
            }
            #Setup all the Script blocks that we are going to use.
            Write-Verbose("Getting Get-ExchangeInstallDirectory string to create Script Block")
            $getExchangeInstallDirectoryString = Add-ScriptBlockInjection @scriptBlockInjectParams `
                -PrimaryScriptBlock ${Function:Get-ExchangeInstallDirectory} `
                -CatchActionFunction ${Function:Invoke-CatchActions}
            Write-Verbose("Creating Script Block")
            $getExchangeInstallDirectoryScriptBlock = [ScriptBlock]::Create($getExchangeInstallDirectoryString)

            Write-Verbose("New-Item create Script Block")
            $newFolderScriptBlock = { param($path) New-Item -ItemType Directory -Path $path -Force | Out-Null }

            $serverArgListExchangeInstallDirectory = @()
            $serverArgListDirectoriesToCreate = @()
            $serverArgListExchangeResideData = @()
            $localServerTempLocation = "{0}{1}\Exchange_Server_Data_Temp\" -f $Script:RootFilePath, $env:COMPUTERNAME

            #Need to do two loops as both of these actions are required before we can do actions in the next loop.
            foreach ($serverData in $exchangeServerData) {
                $serverName = $serverData.ServerName

                $serverArgListExchangeInstallDirectory += [PSCustomObject]@{
                    ServerName   = $serverName
                    ArgumentList = $null
                }

                # Use , prior to the array to make sure it doesn't unwrap
                $serverArgListDirectoriesToCreate += [PSCustomObject]@{
                    ServerName   = $serverName
                    ArgumentList = (, @("$Script:RootFilePath$serverName\Exchange_Server_Data\Config", "$Script:RootFilePath$serverName\Exchange_Server_Data\WebAppPools"))
                }
            }

            Write-Verbose ("Calling job for Get Exchange Install Directory")
            $serverInstallDirectories = Start-JobManager -ServersWithArguments $serverArgListExchangeInstallDirectory `
                -ScriptBlock $getExchangeInstallDirectoryScriptBlock `
                -NeedReturnData $true `
                -JobBatchName "Exchange Install Directories for Write-LargeDataObjectsOnMachine" `
                -RemotePipelineHandler ${Function:Invoke-PipelineHandler}

            Write-Verbose("Calling job for folder creation")
            Start-JobManager -ServersWithArguments $serverArgListDirectoriesToCreate -ScriptBlock $newFolderScriptBlock `
                -JobBatchName "Creating folders for Write-LargeDataObjectsOnMachine" `
                -RemotePipelineHandler ${Function:Invoke-PipelineHandler}

            #Now do the rest of the actions
            foreach ($serverData in $exchangeServerData) {
                $serverName = $serverData.ServerName

                $saveToLocation = "{0}{1}\Exchange_Server_Data" -f $Script:RootFilePath, $serverName
                $serverArgListExchangeResideData += [PSCustomObject]@{
                    ServerName   = $serverName
                    ArgumentList = @($saveToLocation, $serverInstallDirectories[$serverName])
                }

                #Write out the Exchange object data locally as a temp and copy it over to the remote server
                $location = "{0}{1}\Exchange_Server_Data" -f $Script:RootFilePath, $serverName
                Write-Verbose("Location of data should be at: {0}" -f $location)
                $remoteLocation = "\\{0}\{1}" -f $serverName, $location.Replace(":", "$")
                Write-Verbose("Remote Copy Location: {0}" -f $remoteLocation)
                $rootTempLocation = "{0}{1}" -f $localServerTempLocation, $serverName
                Write-Verbose("Local Root Temp Location: {0}" -f $rootTempLocation)
                #Create the temp location and write out the data
                New-Item -ItemType Directory -Path $rootTempLocation -Force | Out-Null
                Write-ExchangeObjectDataLocal -ServerData $serverData -Location $rootTempLocation
                Get-ChildItem $rootTempLocation |
                    ForEach-Object {
                        try {
                            Copy-Item $_.VersionInfo.FileName $remoteLocation
                        } catch {
                            Write-Verbose("Failed to copy data to $remoteLocation. This is likely due to file sharing permissions.")
                            Invoke-CatchActions
                        }
                    }
            }

            #Remove the temp data location right away
            Remove-Item $localServerTempLocation -Force -Recurse

            Write-Verbose("Calling Invoke-ExchangeResideDataCollectionWrite")
            Start-JobManager -ServersWithArguments $serverArgListExchangeResideData -ScriptBlock ${Function:Invoke-ExchangeResideDataCollectionWrite} `
                -DisplayReceiveJob $false `
                -JobBatchName "Write the data for Write-LargeDataObjectsOnMachine"
        } else {

            if ($null -eq $ExInstall) {
                $ExInstall = Get-ExchangeInstallDirectory
            }
            $location = "{0}{1}\Exchange_Server_Data" -f $Script:RootFilePath, $exchangeServerData.ServerName
            [array]$createFolders = @(("{0}\Config" -f $location), ("{0}\WebAppPools" -f $location))
            New-Item -ItemType Directory -Path $createFolders -Force | Out-Null
            Write-ExchangeObjectDataLocal -Location $location -ServerData $exchangeServerData

            $passInfo = @{
                SaveToLocation   = $location
                InstallDirectory = $ExInstall
            }

            Write-Verbose("Writing out the Exchange data")
            Invoke-ExchangeResideDataCollectionWrite @passInfo
        }
    }
}



function Get-TransportLoggingInformationPerServer {
    param(
        [string]$Server,
        [int]$Version,
        [bool]$EdgeServer,
        [bool]$CASOnly,
        [bool]$MailboxOnly
    )
    Write-Verbose("Function Enter: Get-TransportLoggingInformationPerServer")
    Write-Verbose("Passed: [string]Server: {0} | [int]Version: {1} | [bool]EdgeServer: {2} | [bool]CASOnly: {3} | [bool]MailboxOnly: {4}" -f $Server, $Version, $EdgeServer, $CASOnly, $MailboxOnly)
    $transportLoggingObject = New-Object PSCustomObject

    if ($Version -ge 15) {

        if (-not($CASOnly)) {
            #Hub Transport Layer
            $data = Get-TransportService -Identity $Server
            $hubObject = [PSCustomObject]@{
                ConnectivityLogPath    = $data.ConnectivityLogPath.ToString()
                MessageTrackingLogPath = $data.MessageTrackingLogPath.ToString()
                PipelineTracingPath    = $data.PipelineTracingPath.ToString()
                ReceiveProtocolLogPath = $data.ReceiveProtocolLogPath.ToString()
                SendProtocolLogPath    = $data.SendProtocolLogPath.ToString()
                WlmLogPath             = $data.WlmLogPath.ToString()
                RoutingTableLogPath    = $data.RoutingTableLogPath.ToString()
                AgentLogPath           = $data.AgentLogPath.ToString()
            }

            if (![string]::IsNullOrEmpty($data.QueueLogPath)) {
                $hubObject | Add-Member -MemberType NoteProperty -Name "QueueLogPath" -Value ($data.QueueLogPath.ToString())
            }

            $transportLoggingObject | Add-Member -MemberType NoteProperty -Name HubLoggingInfo -Value $hubObject
        }

        if (-not ($EdgeServer)) {
            #Front End Transport Layer
            if (($Version -eq 15 -and (-not ($MailboxOnly))) -or $Version -ge 16) {
                $data = Get-FrontendTransportService -Identity $Server

                if ($Version -ne 15 -and (-not([string]::IsNullOrEmpty($data.RoutingTableLogPath)))) {
                    $routingTableLogPath = $data.RoutingTableLogPath.ToString()
                }

                $FETransObject = [PSCustomObject]@{
                    ConnectivityLogPath    = $data.ConnectivityLogPath.ToString()
                    ReceiveProtocolLogPath = $data.ReceiveProtocolLogPath.ToString()
                    SendProtocolLogPath    = $data.SendProtocolLogPath.ToString()
                    AgentLogPath           = $data.AgentLogPath.ToString()
                    RoutingTableLogPath    = $routingTableLogPath
                }
                $transportLoggingObject | Add-Member -MemberType NoteProperty -Name FELoggingInfo -Value $FETransObject
            }

            if (($Version -eq 15 -and (-not ($CASOnly))) -or $Version -ge 16) {
                #Mailbox Transport Layer
                $data = Get-MailboxTransportService -Identity $Server

                if ($Version -ne 15 -and (-not([string]::IsNullOrEmpty($data.RoutingTableLogPath)))) {
                    $routingTableLogPath = $data.RoutingTableLogPath.ToString()
                }

                $mbxObject = [PSCustomObject]@{
                    ConnectivityLogPath              = $data.ConnectivityLogPath.ToString()
                    ReceiveProtocolLogPath           = $data.ReceiveProtocolLogPath.ToString()
                    SendProtocolLogPath              = $data.SendProtocolLogPath.ToString()
                    PipelineTracingPath              = $data.PipelineTracingPath.ToString()
                    MailboxDeliveryThrottlingLogPath = $data.MailboxDeliveryThrottlingLogPath.ToString()
                    MailboxDeliveryAgentLogPath      = $data.MailboxDeliveryAgentLogPath.ToString()
                    MailboxSubmissionAgentLogPath    = $data.MailboxSubmissionAgentLogPath.ToString()
                    RoutingTableLogPath              = $routingTableLogPath
                }
                $transportLoggingObject | Add-Member -MemberType NoteProperty -Name MBXLoggingInfo -Value $mbxObject
            }
        }
    } elseif ($Version -eq 14) {
        $data = Get-TransportServer -Identity $Server
        $hubObject = New-Object PSCustomObject #TODO Remove because we shouldn't support 2010 any longer
        $hubObject | Add-Member -MemberType NoteProperty -Name ConnectivityLogPath -Value ($data.ConnectivityLogPath.PathName)
        $hubObject | Add-Member -MemberType NoteProperty -Name MessageTrackingLogPath -Value ($data.MessageTrackingLogPath.PathName)
        $hubObject | Add-Member -MemberType NoteProperty -Name PipelineTracingPath -Value ($data.PipelineTracingPath.PathName)
        $hubObject | Add-Member -MemberType NoteProperty -Name ReceiveProtocolLogPath -Value ($data.ReceiveProtocolLogPath.PathName)
        $hubObject | Add-Member -MemberType NoteProperty -Name SendProtocolLogPath -Value ($data.SendProtocolLogPath.PathName)
        $transportLoggingObject | Add-Member -MemberType NoteProperty -Name HubLoggingInfo -Value $hubObject
    } else {
        Write-Host "trying to determine transport information for server $Server and wasn't able to determine the correct version type"
        return
    }

    Write-Verbose("Function Exit: Get-TransportLoggingInformationPerServer")
    return $transportLoggingObject
}
function Get-ServerObjects {
    param(
        [Parameter(Mandatory = $true)][Array]$ValidServers
    )

    Write-Verbose ("Function Enter: Get-ServerObjects")
    Write-Verbose ("Passed: {0} number of Servers" -f $ValidServers.Count)
    $serversObject = @()
    $validServersList = @()
    foreach ($svr in $ValidServers) {
        Write-Verbose ("Working on Server {0}" -f $svr)

        $serverObj = Get-ExchangeBasicServerObject -ServerName $svr
        if ($serverObj -eq $true) {
            Write-Host "Removing Server $svr from the list" -ForegroundColor "Red"
            continue
        } else {
            $validServersList += $svr
        }

        if ($Script:AnyTransportSwitchesEnabled -and ($serverObj.Hub -or $serverObj.Version -ge 15)) {
            $serverObj | Add-Member -Name TransportInfoCollect -MemberType NoteProperty -Value $true
            $serverObj | Add-Member -Name TransportInfo -MemberType NoteProperty -Value `
            (Get-TransportLoggingInformationPerServer -Server $svr `
                    -version $serverObj.Version `
                    -EdgeServer $serverObj.Edge `
                    -CASOnly $serverObj.CASOnly `
                    -MailboxOnly $serverObj.MailboxOnly)
        } else {
            $serverObj | Add-Member -Name TransportInfoCollect -MemberType NoteProperty -Value $false
        }

        if ($PopLogs -and
            !$Script:EdgeRoleDetected) {
            $serverObj | Add-Member -Name PopLogsLocation -MemberType NoteProperty -Value ((Get-PopSettings -Server $svr).LogFileLocation)
        }

        if ($ImapLogs -and
            !$Script:EdgeRoleDetected) {
            $serverObj | Add-Member -Name ImapLogsLocation -MemberType NoteProperty -Value ((Get-ImapSettings -Server $svr).LogFileLocation)
        }

        $serversObject += $serverObj
    }

    if (($null -eq $serversObject) -or
        ($serversObject.Count -eq 0)) {
        Write-Host "Something wrong happened in Get-ServerObjects stopping script" -ForegroundColor "Red"
        exit
    }
    #Set the valid servers
    $Script:ValidServers = $validServersList
    Write-Verbose("Function Exit: Get-ServerObjects")
    return $serversObject
}
function Get-ArgumentList {
    param(
        [Parameter(Mandatory = $true)][array]$Servers
    )

    #First we need to verify if the local computer is in the list or not. If it isn't we need to pick a master server to store
    #the additional information vs having a small amount of data dumped into the local directory.
    $localServerInList = $false
    $Script:MasterServer = $env:COMPUTERNAME
    foreach ($server in $Servers) {

        if ($server.ToUpper().Contains($env:COMPUTERNAME.ToUpper())) {
            $localServerInList = $true
            break
        }
    }

    if (!$localServerInList) {
        $Script:MasterServer = $Servers[0]
    }

    return [PSCustomObject]@{
        AcceptedRemoteDomain           = $AcceptedRemoteDomain
        ADDriverLogs                   = $ADDriverLogs
        AnyTransportSwitchesEnabled    = $Script:AnyTransportSwitchesEnabled
        AppSysLogs                     = $AppSysLogs
        AppSysLogsToXml                = $AppSysLogsToXml
        AutoDLogs                      = $AutoDLogs
        CollectAllLogsBasedOnLogAge    = $CollectAllLogsBasedOnLogAge
        DAGInformation                 = $DAGInformation
        DailyPerformanceLogs           = $DailyPerformanceLogs
        DefaultTransportLogging        = $DefaultTransportLogging
        EASLogs                        = $EASLogs
        ECPLogs                        = $ECPLogs
        EWSLogs                        = $EWSLogs
        ExchangeServerInformation      = $ExchangeServerInformation
        ExMon                          = $ExMon
        ExMonLogmanName                = $ExMonLogmanName
        ExPerfWiz                      = $ExPerfWiz
        ExPerfWizLogmanName            = $ExPerfWizLogmanName
        FilePath                       = $FilePath
        FrontEndConnectivityLogs       = $FrontEndConnectivityLogs
        FrontEndProtocolLogs           = $FrontEndProtocolLogs
        GetVDirs                       = $GetVDirs
        HighAvailabilityLogs           = $HighAvailabilityLogs
        HostExeServerName              = $env:COMPUTERNAME
        HubConnectivityLogs            = $HubConnectivityLogs
        HubProtocolLogs                = $HubProtocolLogs
        IISLogs                        = $IISLogs
        ImapLogs                       = $ImapLogs
        TimeSpan                       = $LogAge
        EndTimeSpan                    = $LogEndAge
        MailboxAssistantsLogs          = $MailboxAssistantsLogs
        MailboxConnectivityLogs        = $MailboxConnectivityLogs
        MailboxDeliveryThrottlingLogs  = $MailboxDeliveryThrottlingLogs
        MailboxProtocolLogs            = $MailboxProtocolLogs
        ManagedAvailabilityLogs        = $ManagedAvailabilityLogs
        MapiLogs                       = $MapiLogs
        MasterServer                   = $Script:MasterServer
        MessageTrackingLogs            = $MessageTrackingLogs
        MitigationService              = $MitigationService
        OABLogs                        = $OABLogs
        OWALogs                        = $OWALogs
        PipelineTracingLogs            = $PipelineTracingLogs
        PopLogs                        = $PopLogs
        PowerShellLogs                 = $PowerShellLogs
        QueueInformation               = $QueueInformation
        RootFilePath                   = $Script:RootFilePath
        RPCLogs                        = $RPCLogs
        SearchLogs                     = $SearchLogs
        SendConnectors                 = $SendConnectors
        ServerInformation              = $ServerInformation
        ServerObjects                  = (Get-ServerObjects -ValidServers $Servers)
        ScriptDebug                    = $ScriptDebug
        StandardFreeSpaceInGBCheckSize = $Script:StandardFreeSpaceInGBCheckSize
        TransportAgentLogs             = $TransportAgentLogs
        TransportConfig                = $TransportConfig
        TransportRoutingTableLogs      = $TransportRoutingTableLogs
        TransportRules                 = $TransportRules
        WindowsSecurityLogs            = $WindowsSecurityLogs
    }
}

#This function is to handle all root zipping capabilities and copying of the data over.
function Invoke-ServerRootZipAndCopy {
    param(
        [bool]$RemoteExecute = $true
    )

    $serverNames = $Script:ArgumentList.ServerObjects |
        ForEach-Object {
            return $_.ServerName
        }

    function Write-CollectFilesFromLocation {
        Write-Host ""
        Write-Host "Please collect the following files from these servers and upload them: "
        $LogPaths |
            ForEach-Object {
                Write-Host "Server: $($_.ServerName) Path: $($_.ZipFolder)"
            }
    }

    if ($RemoteExecute) {
        $Script:ErrorsFromStartOfCopy = $Error.Count
        $Script:Logger = Get-NewLoggerInstance -LogName "ExchangeLogCollector-ZipAndCopy-Debug" -LogDirectory $Script:RootFilePath

        Write-Verbose("Getting Compress-Folder string to create Script Block")
        $compressFolderString = (${Function:Compress-Folder}).ToString()
        Write-Verbose("Creating script block")
        $compressFolderScriptBlock = [ScriptBlock]::Create($compressFolderString)

        $serverArgListZipFolder = @()

        foreach ($serverName in $serverNames) {

            $folder = "{0}{1}" -f $Script:RootFilePath, $serverName
            $serverArgListZipFolder += [PSCustomObject]@{
                ServerName   = $serverName
                ArgumentList = @($folder, $true, $true)
            }
        }

        Write-Verbose("Calling Compress-Folder")
        Start-JobManager -ServersWithArguments $serverArgListZipFolder -ScriptBlock $compressFolderScriptBlock `
            -JobBatchName "Zipping up the data for Invoke-ServerRootZipAndCopy"

        $LogPaths = Invoke-Command -ComputerName $serverNames -ScriptBlock {

            $item = $Using:RootFilePath + (Get-ChildItem $Using:RootFilePath |
                    Where-Object { $_.Name -like ("*-{0}*.zip" -f (Get-Date -Format Md)) } |
                    Sort-Object CreationTime -Descending |
                    Select-Object -First 1)

            return [PSCustomObject]@{
                ServerName = $env:COMPUTERNAME
                ZipFolder  = $item
                Size       = ((Get-Item $item).Length)
            }
        }

        if (!($SkipEndCopyOver)) {
            #Check to see if we have enough free space.
            $LogPaths |
                ForEach-Object {
                    $totalSizeToCopyOver += $_.Size
                }

            $freeSpace = Get-FreeSpace -FilePath $Script:RootFilePath
            $totalSizeGB = $totalSizeToCopyOver / 1GB

            if ($freeSpace -gt ($totalSizeGB + $Script:StandardFreeSpaceInGBCheckSize)) {
                Write-Host "Looks like we have enough free space at the path to copy over the data"
                Write-Host "FreeSpace: $freeSpace TestSize: $(($totalSizeGB + $Script:StandardFreeSpaceInGBCheckSize)) Path: $RootPath"
                Write-Host ""
                Write-Host "Copying over the data may take some time depending on the network"

                $LogPaths |
                    ForEach-Object {
                        if ($_.ServerName -ne $env:COMPUTERNAME) {
                            $remoteCopyLocation = "\\{0}\{1}" -f $_.ServerName, ($_.ZipFolder.Replace(":", "$"))
                            Write-Host "[$($_.ServerName)] : Copying File $remoteCopyLocation...."
                            Copy-Item -Path $remoteCopyLocation -Destination $Script:RootFilePath
                            Write-Host "[$($_.ServerName)] : Done copying file"
                        }
                    }
            } else {
                Write-Host "Looks like we don't have enough free space to copy over the data" -ForegroundColor "Yellow"
                Write-Host "FreeSpace: $FreeSpace TestSize: $(($totalSizeGB + $Script:StandardFreeSpaceInGBCheckSize)) Path: $RootPath"
                Write-CollectFilesFromLocation
            }
        } else {
            Write-CollectFilesFromLocation
        }
    } else {
        Invoke-ZipFolder -Folder $Script:RootCopyToDirectory -ZipItAll $true -AddCompressedSize $false
    }
}

function Test-DiskSpace {
    param(
        [Parameter(Mandatory = $true)][array]$Servers,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][int]$CheckSize
    )
    Write-Verbose("Function Enter: Test-DiskSpace")
    Write-Verbose("Passed: [string]Path: {0} | [int]CheckSize: {1}" -f $Path, $CheckSize)
    Write-Host "Checking the free space on the servers before collecting the data..."
    if (-not ($Path.EndsWith("\"))) {
        $Path = "{0}\" -f $Path
    }

    function Test-ServerDiskSpace {
        param(
            [Parameter(Mandatory = $true)][string]$Server,
            [Parameter(Mandatory = $true)][int]$FreeSpace,
            [Parameter(Mandatory = $true)][int]$CheckSize
        )
        Write-Verbose("Calling Test-ServerDiskSpace")
        Write-Verbose("Passed: [string]Server: {0} | [int]FreeSpace: {1} | [int]CheckSize: {2}" -f $Server, $FreeSpace, $CheckSize)

        if ($FreeSpace -gt $CheckSize) {
            Write-Host "[Server: $Server] : We have more than $CheckSize GB of free space."
            return $true
        } else {
            Write-Host "[Server: $Server] : We have less than $CheckSize GB of free space."
            return $false
        }
    }

    if ($Servers.Count -eq 1 -and $Servers[0] -eq $env:COMPUTERNAME) {
        Write-Verbose("Local server only check. Not going to invoke Start-JobManager")
        $freeSpace = Get-FreeSpace -FilePath $Path
        if (Test-ServerDiskSpace -Server $Servers[0] -FreeSpace $freeSpace -CheckSize $CheckSize) {
            return $Servers[0]
        } else {
            return $null
        }
    }

    $serverArgs = @()
    foreach ($server in $Servers) {
        $serverArgs += [PSCustomObject]@{
            ServerName   = $server
            ArgumentList = $Path
        }
    }

    Write-Verbose("Getting Get-FreeSpace string to create Script Block")
    SetWriteRemoteVerboseAction "New-VerbosePipelineObject"
    $getFreeSpaceString = Add-ScriptBlockInjection -PrimaryScriptBlock ${Function:Get-FreeSpace} `
        -IncludeScriptBlock @(${Function:Write-Verbose}, ${Function:New-PipelineObject}, ${Function:New-VerbosePipelineObject}) `
        -IncludeUsingParameter "WriteRemoteVerboseDebugAction" `
        -CatchActionFunction ${Function:Invoke-CatchActions}
    Write-Verbose("Creating Script Block")
    $getFreeSpaceScriptBlock = [ScriptBlock]::Create($getFreeSpaceString)
    $serversData = Start-JobManager -ServersWithArguments $serverArgs -ScriptBlock $getFreeSpaceScriptBlock `
        -NeedReturnData $true `
        -JobBatchName "Getting the free space for test disk space" `
        -RemotePipelineHandler ${Function:Invoke-PipelineHandler}
    $passedServers = @()
    foreach ($server in $Servers) {

        $freeSpace = $serversData[$server]
        if (Test-ServerDiskSpace -Server $server -FreeSpace $freeSpace -CheckSize $CheckSize) {
            $passedServers += $server
        }
    }

    if ($passedServers.Count -eq 0) {
        Write-Host "Looks like all the servers didn't pass the disk space check."
        Write-Host "Because there are no servers left, we will stop the script."
        exit
    } elseif ($passedServers.Count -ne $Servers.Count) {
        Write-Host "Looks like all the servers didn't pass the disk space check."
        Write-Host "We will only collect data from these servers: "
        foreach ($svr in $passedServers) {
            Write-Host $svr
        }
        Enter-YesNoLoopAction -Question "Collect data only from servers that passed the disk space check?" -YesAction {} -NoAction { exit }
    }
    Write-Verbose("Function Exit: Test-DiskSpace")
    return $passedServers
}

function Test-NoSwitchesProvided {
    if ($EWSLogs -or
        $IISLogs -or
        $DailyPerformanceLogs -or
        $ManagedAvailabilityLogs -or
        $ExPerfWiz -or
        $RPCLogs -or
        $EASLogs -or
        $ECPLogs -or
        $AutoDLogs -or
        $SearchLogs -or
        $OWALogs -or
        $ADDriverLogs -or
        $HighAvailabilityLogs -or
        $MapiLogs -or
        $Script:AnyTransportSwitchesEnabled -or
        $DAGInformation -or
        $GetVDirs -or
        $OrganizationConfig -or
        $ExMon -or
        $ServerInformation -or
        $PopLogs -or
        $ImapLogs -or
        $OABLogs -or
        $PowerShellLogs -or
        $WindowsSecurityLogs -or
        $MailboxAssistantsLogs -or
        $ExchangeServerInformation -or
        $MitigationService
    ) {
        return
    } else {
        Write-Host "`r`nWARNING: Doesn't look like any parameters were provided, are you sure you are running the correct command? This is ONLY going to collect the Application and System Logs." -ForegroundColor "Yellow"
        Enter-YesNoLoopAction -Question "Would you like to continue?" -YesAction { Write-Host "Okay moving on..." } -NoAction { exit }
    }
}

function Test-PossibleCommonScenarios {

    #all possible logs
    if ($AllPossibleLogs) {
        $Script:EWSLogs = $true
        $Script:IISLogs = $true
        $Script:DailyPerformanceLogs = $true
        $Script:ManagedAvailabilityLogs = $true
        $Script:RPCLogs = $true
        $Script:EASLogs = $true
        $Script:AutoDLogs = $true
        $Script:OWALogs = $true
        $Script:ADDriverLogs = $true
        $Script:SearchLogs = $true
        $Script:HighAvailabilityLogs = $true
        $Script:ServerInformation = $true
        $Script:GetVDirs = $true
        $Script:DAGInformation = $true
        $Script:DefaultTransportLogging = $true
        $Script:MapiLogs = $true
        $Script:OrganizationConfig = $true
        $Script:ECPLogs = $true
        $Script:ExchangeServerInformation = $true
        $Script:PopLogs = $true
        $Script:ImapLogs = $true
        $Script:ExPerfWiz = $true
        $Script:OABLogs = $true
        $Script:PowerShellLogs = $true
        $Script:WindowsSecurityLogs = $true
        $Script:CollectFailoverMetrics = $true
        $Script:ConnectivityLogs = $true
        $Script:ProtocolLogs = $true
        $Script:MitigationService = $true
        $Script:MailboxAssistantsLogs = $true
    }

    if ($DefaultTransportLogging) {
        $Script:HubConnectivityLogs = $true
        $Script:MessageTrackingLogs = $true
        $Script:QueueInformation = $true
        $Script:SendConnectors = $true
        $Script:ReceiveConnectors = $true
        $Script:TransportAgentLogs = $true
        $Script:TransportConfig = $true
        $Script:TransportRoutingTableLogs = $true
        $Script:FrontEndConnectivityLogs = $true
        $Script:MailboxConnectivityLogs = $true
        $Script:FrontEndProtocolLogs = $true
        $Script:MailboxDeliveryThrottlingLogs = $true
        $Script:PipelineTracingLogs = $true
        $Script:TransportRules = $true
        $Script:AcceptedRemoteDomain = $true
    }

    if ($ConnectivityLogs) {
        $Script:FrontEndConnectivityLogs = $true
        $Script:HubConnectivityLogs = $true
        $Script:MailboxConnectivityLogs = $true
    }

    if ($ProtocolLogs) {
        $Script:FrontEndProtocolLogs = $true
        $Script:HubProtocolLogs = $true
        $Script:MailboxProtocolLogs = $true
    }

    if ($DatabaseFailoverIssue) {
        $Script:DailyPerformanceLogs = $true
        $Script:HighAvailabilityLogs = $true
        $Script:ManagedAvailabilityLogs = $true
        $Script:DAGInformation = $true
        $Script:ExPerfWiz = $true
        $Script:ServerInformation = $true
        $Script:CollectFailoverMetrics = $true
    }

    if ($PerformanceIssues) {
        $Script:DailyPerformanceLogs = $true
        $Script:ManagedAvailabilityLogs = $true
        $Script:ExPerfWiz = $true
    }

    if ($PerformanceMailFlowIssues) {
        $Script:DailyPerformanceLogs = $true
        $Script:ExPerfWiz = $true
        $Script:MessageTrackingLogs = $true
        $Script:QueueInformation = $true
        $Script:TransportConfig = $true
        $Script:TransportRules = $true
        $Script:AcceptedRemoteDomain = $true
    }

    if ($OutlookConnectivityIssues) {
        $Script:DailyPerformanceLogs = $true
        $Script:ExPerfWiz = $true
        $Script:IISLogs = $true
        $Script:MapiLogs = $true
        $Script:RPCLogs = $true
        $Script:AutoDLogs = $true
        $Script:EWSLogs = $true
        $Script:ServerInformation = $true
    }

    #Because we right out our Receive Connector information in Exchange Server Info now
    if ($ReceiveConnectors -or
        $QueueInformation) {
        $Script:ExchangeServerInformation = $true
    }

    #See if any transport logging is enabled.
    $Script:AnyTransportSwitchesEnabled = $false
    if ($HubProtocolLogs -or
        $HubConnectivityLogs -or
        $MessageTrackingLogs -or
        $QueueInformation -or
        $SendConnectors -or
        $ReceiveConnectors -or
        $TransportConfig -or
        $FrontEndConnectivityLogs -or
        $FrontEndProtocolLogs -or
        $MailboxConnectivityLogs -or
        $MailboxProtocolLogs -or
        $MailboxDeliveryThrottlingLogs -or
        $TransportAgentLogs -or
        $TransportRoutingTableLogs -or
        $DefaultTransportLogging -or
        $PipelineTracingLogs -or
        $TransportRules -or
        $AcceptedRemoteDomain) {
        $Script:AnyTransportSwitchesEnabled = $true
    }

    if ($ServerInformation -or $ManagedAvailabilityLogs) {
        $Script:ExchangeServerInformation = $true
    }
}

function Test-RemoteExecutionOfServers {
    param(
        [Parameter(Mandatory = $true)][Array]$ServerList
    )
    Write-Verbose("Function Enter: Test-RemoteExecutionOfServers")
    Write-Host "Checking to see if the servers are up in this list:"
    $ServerList | ForEach-Object { Write-Host $_ }
    #Going to just use Invoke-Command to see if the servers are up. As ICMP might be disabled in the environment.
    Write-Host ""
    Write-Host "For all the servers in the list, checking to see if Invoke-Command will work against them."
    #shouldn't need to test if they are Exchange servers, as we should be doing that locally as well.
    $validServers = @()
    foreach ($server in $ServerList) {

        try {
            Write-Host "Checking Server $server....." -NoNewline
            Invoke-Command -ComputerName $server -ScriptBlock { Get-Process | Out-Null } -ErrorAction Stop
            #if that doesn't fail, we should be okay to add it to the working list
            Write-Host "Passed" -ForegroundColor "Green"
            $validServers += $server
        } catch {
            Write-Host "Failed" -ForegroundColor "Red"
            Write-Host "Removing Server $server from the list to collect data from"
            Invoke-CatchActions
        }
    }

    if ($validServers.Count -gt 0) {
        $validServers = Test-DiskSpace -Servers $validServers -Path $FilePath -CheckSize $Script:StandardFreeSpaceInGBCheckSize
    }

    #all servers in teh list weren't able to do Invoke-Command or didn't have enough free space. Try to do against local server.
    if ($null -ne $validServers -and
        $validServers.Count -eq 0) {

        #Can't do this on a tools or remote shell
        if ($Script:LocalExchangeShell.ToolsOnly -or
            $Script:LocalExchangeShell.RemoteShell) {
            Write-Host "Failed to invoke against the machines to do remote collection from a tools box or a remote machine." -ForegroundColor "Red"
            exit
        }

        Write-Host "Failed to do remote collection for all the servers in the list..." -ForegroundColor "Yellow"

        if ((Enter-YesNoLoopAction -Question "Do you want to collect from the local server only?" -YesAction { return $true } -NoAction { return $false })) {
            $validServers = @($env:COMPUTERNAME)
        } else {
            exit
        }

        #want to test local server's free space first before moving to just collecting the data
        if ($null -eq (Test-DiskSpace -Servers $validServers -Path $FilePath -CheckSize $Script:StandardFreeSpaceInGBCheckSize)) {
            Write-Host "Failed to have enough space available locally. We can't continue with the data collection" -ForegroundColor "Yellow"
            exit
        }
    }

    Write-Verbose("Function Exit: Test-RemoteExecutionOfServers")
    return $validServers
}

    function Main {

        Test-PossibleCommonScenarios
        Test-NoSwitchesProvided

        if ( $PSCmdlet.ParameterSetName -eq "LogPeriod" -and ( $LogAge.CompareTo($LogEndAge) -ne 1 ) ) {
            Write-Host "LogStartDate time should smaller than LogEndDate time." -ForegroundColor "Yellow"
            exit
        }

        if (-not (Confirm-Administrator)) {
            Write-Host "Hey! The script needs to be executed in elevated mode. Start the Exchange Management Shell as an Administrator." -ForegroundColor "Yellow"
            exit
        }

        $Script:LocalExchangeShell = Confirm-ExchangeShell

        if (!($Script:LocalExchangeShell.ShellLoaded)) {
            Write-Host "It appears that you are not on an Exchange 2010 or newer server. Sorry I am going to quit."
            exit
        }

        if (!$Script:LocalExchangeShell.RemoteShell) {
            $Script:localExInstall = Get-ExchangeInstallDirectory
        }

        if ($Script:LocalExchangeShell.EdgeServer) {
            #If we are on an Exchange Edge Server, we are going to treat it like a single server on purpose as we recommend that the Edge Server is a non domain joined computer.
            #Because it isn't a domain joined computer, we can't use remote execution
            Write-Host "Determined that we are on an Edge Server, we can only use locally collection for this role." -ForegroundColor "Yellow"
            $Script:EdgeRoleDetected = $true
            $serversToProcess = @($env:COMPUTERNAME)
        }

        if ($null -ne $serversToProcess -and
            !($serversToProcess.Count -eq 1 -and
                $serversToProcess[0].ToUpper().Equals($env:COMPUTERNAME.ToUpper()))) {
            [array]$Script:ValidServers = Test-RemoteExecutionOfServers -ServerList $serversToProcess
        } else {
            [array]$Script:ValidServers = $serversToProcess
        }

        #possible to return null or only a single server back (localhost)
        if (!($null -ne $Script:ValidServers -and
                $Script:ValidServers.Count -eq 1 -and
                $Script:ValidServers[0].ToUpper().Equals($env:COMPUTERNAME.ToUpper()))) {

            $Script:ArgumentList = Get-ArgumentList -Servers $Script:ValidServers
            #I can do a try catch here, but i also need to do a try catch in the remote so i don't end up failing here and assume the wrong failure location
            try {
                Invoke-Command -ComputerName $Script:ValidServers -ScriptBlock ${Function:Invoke-RemoteFunctions} -ArgumentList $argumentList -ErrorAction Stop
            } catch {
                Write-Host "An error has occurred attempting to call Invoke-Command to do a remote collect all at once. Please notify ExToolsFeedback@microsoft.com of this issue. Stopping the script." -ForegroundColor "Red"
                Invoke-CatchActions
                exit
            }

            Write-DataOnlyOnceOnMasterServer
            Write-LargeDataObjectsOnMachine
            Invoke-ServerRootZipAndCopy
        } else {

            if ($null -eq (Test-DiskSpace -Servers $env:COMPUTERNAME -Path $FilePath -CheckSize $Script:StandardFreeSpaceInGBCheckSize)) {
                Write-Host "Failed to have enough space available locally. We can't continue with the data collection" -ForegroundColor "Yellow"
                exit
            }
            if (-not($Script:EdgeRoleDetected)) {
                Write-Host "Note: Remote Collection is now possible for Windows Server 2012 and greater on the remote machine. Just use the -Servers parameter with a list of Exchange Server names" -ForegroundColor "Yellow"
                Write-Host "Going to collect the data locally"
            }
            $Script:ArgumentList = (Get-ArgumentList -Servers $env:COMPUTERNAME)
            Invoke-RemoteFunctions -PassedInfo $Script:ArgumentList
            # Don't manipulate the host object when running locally after the Invoke-RemoteFunctions to
            # make it the same as when having multiple servers executing the script against.
            SetWriteHostManipulateObjectAction $null
            Write-DataOnlyOnceOnMasterServer
            Write-LargeDataObjectsOnMachine
            Invoke-ServerRootZipAndCopy -RemoteExecute $false
        }

        Write-Host "`r`n`r`n`r`nLooks like the script is done. If you ran into any issues or have additional feedback, please feel free to reach out ExToolsFeedback@microsoft.com."
    }

    try {
        <#
        Added the ability to call functions from within a bundled function so i don't have to duplicate work.
        Loading the functions into memory by using the '.' allows me to do this,
        providing that the calling of that function doesn't do anything of value when doing this.
        #>
        . Invoke-RemoteFunctions -PassedInfo ([PSCustomObject]@{
                ByPass = $true
            })

        Invoke-ErrorMonitoring

        $Script:RootFilePath = "{0}\{1}\" -f $FilePath, (Get-Date -Format yyyyMd)
        $Script:Logger = Get-NewLoggerInstance -LogName "ExchangeLogCollector-Main-Debug" -LogDirectory ("$RootFilePath$env:COMPUTERNAME")
        SetWriteVerboseAction ${Function:Write-DebugLog}
        SetWriteHostAction ${Function:Write-DebugLog}

        Main
    } finally {

        if ($Script:VerboseEnabled -or
        ($Error.Count -ne $Script:ErrorsFromStartOfCopy)) {
            #$Script:Logger.RemoveLatestLogFile()
        }
    }
}

# SIG # Begin signature block
# MIIoRQYJKoZIhvcNAQcCoIIoNjCCKDICAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBDgmG9V3r3oeNZ
# AMu88xDPzaP2Q2tqJDbvod9fAe7/jqCCDXYwggX0MIID3KADAgECAhMzAAADrzBA
# DkyjTQVBAAAAAAOvMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMxMTE2MTkwOTAwWhcNMjQxMTE0MTkwOTAwWjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDOS8s1ra6f0YGtg0OhEaQa/t3Q+q1MEHhWJhqQVuO5amYXQpy8MDPNoJYk+FWA
# hePP5LxwcSge5aen+f5Q6WNPd6EDxGzotvVpNi5ve0H97S3F7C/axDfKxyNh21MG
# 0W8Sb0vxi/vorcLHOL9i+t2D6yvvDzLlEefUCbQV/zGCBjXGlYJcUj6RAzXyeNAN
# xSpKXAGd7Fh+ocGHPPphcD9LQTOJgG7Y7aYztHqBLJiQQ4eAgZNU4ac6+8LnEGAL
# go1ydC5BJEuJQjYKbNTy959HrKSu7LO3Ws0w8jw6pYdC1IMpdTkk2puTgY2PDNzB
# tLM4evG7FYer3WX+8t1UMYNTAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQURxxxNPIEPGSO8kqz+bgCAQWGXsEw
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMTgyNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAISxFt/zR2frTFPB45Yd
# mhZpB2nNJoOoi+qlgcTlnO4QwlYN1w/vYwbDy/oFJolD5r6FMJd0RGcgEM8q9TgQ
# 2OC7gQEmhweVJ7yuKJlQBH7P7Pg5RiqgV3cSonJ+OM4kFHbP3gPLiyzssSQdRuPY
# 1mIWoGg9i7Y4ZC8ST7WhpSyc0pns2XsUe1XsIjaUcGu7zd7gg97eCUiLRdVklPmp
# XobH9CEAWakRUGNICYN2AgjhRTC4j3KJfqMkU04R6Toyh4/Toswm1uoDcGr5laYn
# TfcX3u5WnJqJLhuPe8Uj9kGAOcyo0O1mNwDa+LhFEzB6CB32+wfJMumfr6degvLT
# e8x55urQLeTjimBQgS49BSUkhFN7ois3cZyNpnrMca5AZaC7pLI72vuqSsSlLalG
# OcZmPHZGYJqZ0BacN274OZ80Q8B11iNokns9Od348bMb5Z4fihxaBWebl8kWEi2O
# PvQImOAeq3nt7UWJBzJYLAGEpfasaA3ZQgIcEXdD+uwo6ymMzDY6UamFOfYqYWXk
# ntxDGu7ngD2ugKUuccYKJJRiiz+LAUcj90BVcSHRLQop9N8zoALr/1sJuwPrVAtx
# HNEgSW+AKBqIxYWM4Ev32l6agSUAezLMbq5f3d8x9qzT031jMDT+sUAoCw0M5wVt
# CUQcqINPuYjbS1WgJyZIiEkBMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGiUwghohAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAOvMEAOTKNNBUEAAAAAA68wDQYJYIZIAWUDBAIB
# BQCggcYwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIBy4uUTnZjwHuZgzhIyB8vgN
# gMWdhyE1aenTXPiYqaSNMFoGCisGAQQBgjcCAQwxTDBKoBqAGABDAFMAUwAgAEUA
# eABjAGgAYQBuAGcAZaEsgCpodHRwczovL2dpdGh1Yi5jb20vbWljcm9zb2Z0L0NT
# Uy1FeGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAEggEAmiuOt3ws7M/DY0CFUh1RKfaL
# Vv7ztYgiYfHt1ravZnIvSFGalMb2AYzSJ0XUrrqPhmrSoYTJuRgMgeVZBgkRIeFm
# x6w5k3H6xMi8y9IGG1mZ8yhIVcyP5sRw5qFrQwJk6+6Cmy6r1j+YKzmrWOekCQmg
# FNtTs4HRQUNzj6NbNlxoF8eAHzGufNuNOKQ4wyegue+7dKrNVNAHHbTbgHEXw0yw
# tgEoawgZrg49XmOUKMWq+AEMo0zWlWC8cgp0EzM0Y1ctC/lFqjinvXNZ+FLDCZaj
# JApjc0U8uzr7HTE1nTkTO8eoMKvhQ+K5GVBJa5dNbKjGzdP4myj3c8vcMAfiNaGC
# F5cwgheTBgorBgEEAYI3AwMBMYIXgzCCF38GCSqGSIb3DQEHAqCCF3AwghdsAgED
# MQ8wDQYJYIZIAWUDBAIBBQAwggFSBgsqhkiG9w0BCRABBKCCAUEEggE9MIIBOQIB
# AQYKKwYBBAGEWQoDATAxMA0GCWCGSAFlAwQCAQUABCCw3SsKqXjNTtOhwbnqPbyY
# oiqt4P1JxvV2+s4nX0hpKwIGZfxmbwj8GBMyMDI0MDQwNTE4MjkzNy45OTlaMASA
# AgH0oIHRpIHOMIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQL
# Ex5uU2hpZWxkIFRTUyBFU046N0YwMC0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jv
# c29mdCBUaW1lLVN0YW1wIFNlcnZpY2WgghHtMIIHIDCCBQigAwIBAgITMwAAAfAq
# fB1ZO+YfrQABAAAB8DANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFt
# cCBQQ0EgMjAxMDAeFw0yMzEyMDYxODQ1NTFaFw0yNTAzMDUxODQ1NTFaMIHLMQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNy
# b3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBF
# U046N0YwMC0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFNlcnZpY2UwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQC1Hi1Tozh3
# O0czE8xfRnrymlJNCaGWommPy0eINf+4EJr7rf8tSzlgE8Il4Zj48T5fTTOAh6nI
# TRf2lK7+upcnZ/xg0AKoDYpBQOWrL9ObFShylIHfr/DQ4PsRX8GRtInuJsMkwSg6
# 3bfB4Q2UikMEP/CtZHi8xW5XtAKp95cs3mvUCMvIAA83Jr/UyADACJXVU4maYisc
# zUz7J111eD1KrG9mQ+ITgnRR/X2xTDMCz+io8ZZFHGwEZg+c3vmPp87m4OqOKWyh
# cqMUupPveO/gQC9Rv4szLNGDaoePeK6IU0JqcGjXqxbcEoS/s1hCgPd7Ux6YWeWr
# UXaxbb+JosgOazUgUGs1aqpnLjz0YKfUqn8i5TbmR1dqElR4QA+OZfeVhpTonrM4
# sE/MlJ1JLpR2FwAIHUeMfotXNQiytYfRBUOJHFeJYEflZgVk0Xx/4kZBdzgFQPOW
# fVd2NozXlC2epGtUjaluA2osOvQHZzGOoKTvWUPX99MssGObO0xJHd0DygP/JAVp
# +bRGJqa2u7AqLm2+tAT26yI5veccDmNZsg3vDh1HcpCJa9QpRW/MD3a+AF2ygV1s
# RnGVUVG3VODX3BhGT8TMU/GiUy3h7ClXOxmZ+weCuIOzCkTDbK5OlAS8qSPpgp+X
# GlOLEPaM31Mgf6YTppAaeP0ophx345ohtwIDAQABo4IBSTCCAUUwHQYDVR0OBBYE
# FNCCsqdXRy/MmjZGVTAvx7YFWpslMB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWn
# G1M1GelyMF8GA1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNv
# bS9wa2lvcHMvY3JsL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEw
# KDEpLmNybDBsBggrBgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFt
# cCUyMFBDQSUyMDIwMTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAww
# CgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQA4
# IvSbnr4jEPgo5W4xj3/+0dCGwsz863QGZ2mB9Z4SwtGGLMvwfsRUs3NIlPD/LsWA
# xdVYHklAzwLTwQ5M+PRdy92DGftyEOGMHfut7Gq8L3RUcvrvr0AL/NNtfEpbAEkC
# FzseextY5s3hzj3rX2wvoBZm2ythwcLeZmMgHQCmjZp/20fHWJgrjPYjse6RDJtU
# TlvUsjr+878/t+vrQEIqlmebCeEi+VQVxc7wF0LuMTw/gCWdcqHoqL52JotxKzY8
# jZSQ7ccNHhC4eHGFRpaKeiSQ0GXtlbGIbP4kW1O3JzlKjfwG62NCSvfmM1iPD90X
# YiFm7/8mgR16AmqefDsfjBCWwf3qheIMfgZzWqeEz8laFmM8DdkXjuOCQE/2L0Tx
# hrjUtdMkATfXdZjYRlscBDyr8zGMlprFC7LcxqCXlhxhtd2CM+mpcTc8RB2D3Eor
# 0UdoP36Q9r4XWCVV/2Kn0AXtvWxvIfyOFm5aLl0eEzkhfv/XmUlBeOCElS7jdddW
# pBlQjJuHHUHjOVGXlrJT7X4hicF1o23x5U+j7qPKBceryP2/1oxfmHc6uBXlXBKu
# kV/QCZBVAiBMYJhnktakWHpo9uIeSnYT6Qx7wf2RauYHIER8SLRmblMzPOs+JHQz
# rvh7xStx310LOp+0DaOXs8xjZvhpn+WuZij5RmZijDCCB3EwggVZoAMCAQICEzMA
# AAAVxedrngKbSZkAAAAAABUwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVT
# MRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQK
# ExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eSAyMDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMw
# MDkzMDE4MzIyNVowfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggIiMA0G
# CSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDk4aZM57RyIQt5osvXJHm9DtWC0/3u
# nAcH0qlsTnXIyjVX9gF/bErg4r25PhdgM/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1
# jRPPdzLAEBjoYH1qUoNEt6aORmsHFPPFdvWGUNzBRMhxXFExN6AKOG6N7dcP2CZT
# fDlhAnrEqv1yaa8dq6z2Nr41JmTamDu6GnszrYBbfowQHJ1S/rboYiXcag/PXfT+
# jlPP1uyFVk3v3byNpOORj7I5LFGc6XBpDco2LXCOMcg1KL3jtIckw+DJj361VI/c
# +gVVmG1oO5pGve2krnopN6zL64NF50ZuyjLVwIYwXE8s4mKyzbnijYjklqwBSru+
# cakXW2dg3viSkR4dPf0gz3N9QZpGdc3EXzTdEonW/aUgfX782Z5F37ZyL9t9X4C6
# 26p+Nuw2TPYrbqgSUei/BQOj0XOmTTd0lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV
# 2HM9Q07BMzlMjgK8QmguEOqEUUbi0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoS
# CtdjbwzJNmSLW6CmgyFdXzB0kZSU2LlQ+QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxS
# UV0S2yW6r1AFemzFER1y7435UsSFF5PAPBXbGjfHCBUYP3irRbb1Hode2o+eFnJp
# xq57t7c+auIurQIDAQABo4IB3TCCAdkwEgYJKwYBBAGCNxUBBAUCAwEAATAjBgkr
# BgEEAYI3FQIEFgQUKqdS/mTEmr6CkTxGNSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0A
# XmJdg/Tl0mWnG1M1GelyMFwGA1UdIARVMFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYI
# KwYBBQUHAgEWM2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9S
# ZXBvc2l0b3J5Lmh0bTATBgNVHSUEDDAKBggrBgEFBQcDCDAZBgkrBgEEAYI3FAIE
# DB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNV
# HSMEGDAWgBTV9lbLj+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEugSaBHhkVo
# dHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29D
# ZXJBdXRfMjAxMC0wNi0yMy5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAC
# hj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1
# dF8yMDEwLTA2LTIzLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwEx
# JFvhnnJL/Klv6lwUtj5OR2R4sQaTlz0xM7U518JxNj/aZGx80HU5bbsPMeTCj/ts
# 0aGUGCLu6WZnOlNN3Zi6th542DYunKmCVgADsAW+iehp4LoJ7nvfam++Kctu2D9I
# dQHZGN5tggz1bSNU5HhTdSRXud2f8449xvNo32X2pFaq95W2KFUn0CS9QKC/GbYS
# EhFdPSfgQJY4rPf5KYnDvBewVIVCs/wMnosZiefwC2qBwoEZQhlSdYo2wh3DYXMu
# LGt7bj8sCXgU6ZGyqVvfSaN0DLzskYDSPeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT9
# 9kxybxCrdTDFNLB62FD+CljdQDzHVG2dY3RILLFORy3BFARxv2T5JL5zbcqOCb2z
# AVdJVGTZc9d/HltEAY5aGZFrDZ+kKNxnGSgkujhLmm77IVRrakURR6nxt67I6Ile
# T53S0Ex2tVdUCbFpAUR+fKFhbHP+CrvsQWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6l
# MVGEvL8CwYKiexcdFYmNcP7ntdAoGokLjzbaukz5m/8K6TT4JDVnK+ANuOaMmdbh
# IurwJ0I9JZTmdHRbatGePu1+oDEzfbzL6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3u
# gm2lBRDBcQZqELQdVTNYs6FwZvKhggNQMIICOAIBATCB+aGB0aSBzjCByzELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9z
# b2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNO
# OjdGMDAtMDVFMC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNloiMKAQEwBwYFKw4DAhoDFQDCKAZKKv5lsdC2yoMGKYiQy79p/6CBgzCB
# gKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNV
# BAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUA
# AgUA6bqqmjAiGA8yMDI0MDQwNTE2NTEzOFoYDzIwMjQwNDA2MTY1MTM4WjB3MD0G
# CisGAQQBhFkKBAExLzAtMAoCBQDpuqqaAgEAMAoCAQACAh77AgH/MAcCAQACAhMz
# MAoCBQDpu/waAgEAMDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAI
# AgEAAgMHoSChCjAIAgEAAgMBhqAwDQYJKoZIhvcNAQELBQADggEBABTfA86rm8FF
# Fudexe2cqhd8RKoJbIH6aaIG5sZaa2FO/MEANTuF8x1DOpVCD2HFrvv7W+bf0e6x
# +lhykKkrhlt+3KkBashKSPysA+gbtKPaTLTHO2KQ2QSPb0Vh5QXKxGVXVO0ojSlS
# qpJvZYcRfEArtcNECTIgjhhZEjF3RSRad3rtXcIL/E9x1sT5f/fDqP6WnLALGAkV
# Zee2VYpgRxANODMy2ZdblBvti26n/kCIBEEwER4Pf50P3vd8LfmgV9LEu2P2Q3Ai
# nEulyfvm3oOmi7KeTO3HC88rY8ZOg7WIBCJq7ZyuZ3rlmtPQxLv43+GK3z4KuRCE
# Rsf7Z0rshEQxggQNMIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0Eg
# MjAxMAITMwAAAfAqfB1ZO+YfrQABAAAB8DANBglghkgBZQMEAgEFAKCCAUowGgYJ
# KoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCCvAnh9is/9
# hLS8q6twMh/T2qHjutgBTsN92Ko90xhDRTCB+gYLKoZIhvcNAQkQAi8xgeowgecw
# geQwgb0EIFwBmqOlcv3kU7mAB5sWR74QFAiS6mb+CM6asnFAZUuLMIGYMIGApH4w
# fDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1Jl
# ZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMd
# TWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAHwKnwdWTvmH60AAQAA
# AfAwIgQgZ6IyoCK6qBXqLuSyvA4I1G4MR+uV5p8Ho+9yaWDOg3EwDQYJKoZIhvcN
# AQELBQAEggIAe879q4eEqNHeXmTSFEFl5r05vST8zJlAgvxfwRUb4S7+RrmMp166
# UXeJR3n04gv+6O+c++nuqoLdskxeKpsBjSStpHArX7fcwP4lWXPfbZhyp9RYCRN0
# cWrSSU6a74UIbI4yymWJnTRhWeuNO1mpkhJ0EPckPBgfbneo35b9F8vKkIWAr4PI
# tBaYMmRPMceNbsuQSeTqYXT/WJFvTsP+A8uZUE/o1ryeAsw6t/pB2dgYL4KNEkCg
# DfiMnlytizthx2OHYTh6w5t0+gghrgpjCGEoYehw+6kPsxpo3V3V21WzVzDexpmc
# 5ljcVdzWQ9VpF4u0/unkY6jzc9Ot43YLNlcJDuNFjBIrFJ1F+aqtPG46EbV78Y4x
# NKPryzUuacTJ/pD0fR7TA6QyN7y7jxWqGyVgaxn22kWFi9kuLkpmjDhzpkB1Iusz
# evy4jh6YBNWVKsrGO2CU2iG0HtcGCLyaj9BpSHgQ7sffSmsj39favK1VWPwLzC+x
# 49lTSJmYSIzN/5ncFubawAweKcu1VmvWerxsGYhVO9uewxXF8YJeX9gLHGcNhqHs
# K3gXVJtrDeVkb3zv1yNFViVG1NAURCsh5tmmP6GPDh1Kmrq/Fl+w4VjD46WIHn9L
# p3sPBkycjo3oaBIimzJxcb8B+4NkDssPGXl6VxsuS2PI4xwZQMkdCdg=
# SIG # End signature block
