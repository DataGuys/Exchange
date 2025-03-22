# Define the action to be scheduled
$action = New-ScheduledTaskAction -Execute 'Powershell.exe' -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -File C:\scripts\EndUserComms.ps1"

# Define the trigger for the task to start daily
$trigger = New-ScheduledTaskTrigger -At (Get-Date).Date -Daily

# Register the scheduled task
Register-ScheduledTask -Action $action -Trigger $trigger -TaskName "EndUserComms" -Description "Sends EndUser Communications"
