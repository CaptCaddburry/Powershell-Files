$taskName = "WeeklyUptimeRebootCheck"

$scriptBlock = {
    $uptimeDays = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
    $uptime = (New-TimeSpan -Start $uptimeDays -End (Get-Date)).Days
    if ($uptime -ge 7) {
        $Minutes = 15
        $Seconds = [Math]::Max(60, $Minutes * 60)
        $ETA = (Get-Date).AddSeconds($Seconds).ToString("hh:mm tt")
        $ShutdownComment = "IT Notice: Your computer needs to be restarted in order to install some needed security patches. Your computer will be restarted at $ETA. Please save your work now."
        & shutdown /r /t $Seconds /c $ShutdownComment /d p:5:19
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
