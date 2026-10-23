$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$py = (& python -c "import sys,os;print(os.path.join(os.path.dirname(sys.executable),'pythonw.exe'))").Trim()
$action = New-ScheduledTaskAction -Execute $py -Argument "`"$dir\mail_collector.py`"" -WorkingDirectory $dir
$trigger = New-ScheduledTaskTrigger -Daily -At 17:00
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
Register-ScheduledTask -TaskName "Kerem Mail Feed" -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null
Write-Host "Task updated: runs on battery too, and catches up if the PC was asleep at 17:00."
Start-ScheduledTask -TaskName "Kerem Mail Feed"
Start-Sleep -Seconds 20
Get-ScheduledTask "Kerem Mail Feed" | Get-ScheduledTaskInfo | Format-List LastRunTime, LastTaskResult, NextRunTime
Get-Content "$dir\run_log.txt" -Tail 3
