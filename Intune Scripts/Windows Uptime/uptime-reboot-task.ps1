$taskName = "WeeklyUptimeRebootCheck"

$scriptBlock = {
    $uptimeDays = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
    $uptime = (New-TimeSpan -Start `$uptimeDays -End (Get-Date)).Days
    if ($uptime -ge 7) {
        shutdown.exe /r /t 900 /c 'IT Notice: Automatic Restart in 15 minute(s). Save your work now.' /d p:5:19
    }
}

$encodedScript = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($scriptBlock))
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $encodedScript"

$trigger = New-ScheduledTaskTrigger -AtLogon
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest

Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Force
