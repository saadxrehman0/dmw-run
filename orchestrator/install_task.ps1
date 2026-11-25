# Register the 8:55 AM edition run in Windows Task Scheduler (Sunday to Thursday).
# Run once on the RUNNER (the office server), in PowerShell as the dedicated DMW user:
#   powershell -ExecutionPolicy Bypass -File orchestrator\install_task.ps1 [-Python C:\Python315\python.exe]
# The task runs only while that user is logged on, because PowerPoint needs a desktop session
# (plan doc, Build steps 1). The PC's clock must be on Pakistan time (UTC+5).
# Remove with:  Unregister-ScheduledTask -TaskName "DMW edition run" -Confirm:$false
param(
  [string]$Python = "C:\Python315\python.exe",
  [string]$Repo = (Split-Path -Parent $PSScriptRoot),
  [string]$At = "08:55"
)
$script = Join-Path $Repo "scripts\run_edition.py"
$log = Join-Path $Repo "out\run_edition.log"
New-Item -ItemType Directory -Force -Path (Join-Path $Repo "out") | Out-Null
$action = New-ScheduledTaskAction -Execute "cmd.exe" `
  -Argument "/c `"`"$Python`" `"$script`" >> `"$log`" 2>&1`"" -WorkingDirectory $Repo
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday,Monday,Tuesday,Wednesday,Thursday -At $At
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopIfGoingOnBatteries `
  -AllowStartIfOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 4) -MultipleInstances IgnoreNew `
  -WakeToRun
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName "DMW edition run" -Action $action -Trigger $trigger `
  -Settings $settings -Principal $principal -Force
Write-Host "Registered 'DMW edition run' at $At, Sunday to Thursday, for $env:USERNAME."
