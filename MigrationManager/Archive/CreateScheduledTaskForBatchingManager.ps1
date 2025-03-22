$action = New-ScheduledTaskAction -Execute 'Powershell.exe' -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden C:\scripts\BatchingManagerV2.ps1"
# Calculate the next top or bottom of the hour from the current time
$CurrentTime = Get-Date
$NextHalfHour = if ($CurrentTime.Minute -lt 30) { 
    Get-Date $CurrentTime -Minute 30 -Second 0 
} else { 
    Get-Date $CurrentTime.AddHours(1) -Minute 0 -Second 0 
}

# Define the trigger to start at the calculated time and repeat every 30 minutes for a duration of 10 years
$Duration = New-TimeSpan -Days (10 * 365) # 10 years
$Trigger = New-ScheduledTaskTrigger -Once -At $NextHalfHour -RepetitionInterval (New-TimeSpan -Minutes 30) -RepetitionDuration $Duration


# Define the trigger to start at the calculated time and repeat every 30 minutes


Register-ScheduledTask -Action $action -Trigger $trigger -TaskName "BatchingManager" -Description "30 minute batching manager"