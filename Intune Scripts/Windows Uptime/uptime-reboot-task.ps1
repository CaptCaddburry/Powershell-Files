$taskName = "WeeklyUptimeRebootCheck"

$scriptBlock = {
    $Minutes = 15
    $Seconds = [Math]::Max(60, $Minutes * 60)
    $ETA = (Get-Date).AddSeconds($Seconds).ToString("hh:mm tt")
    $XML = @"
        <toast duration="long">
            <visual>
                <binding template="ToastGeneric">
                    <text>Windows Restart Notification</text>
                    <text>Your computer needs to be rebooted for security purposes. Your computer will restart at $ETA. Please save your work now.</text>
                </binding>
            </visual>
        </toast>
"@
    $XMLDocument = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]::New()
    $XMLDocument.LoadXml($XML)
    $AppId = (Get-StartApps | Where -Property Name -EQ "Windows PowerShell").AppId
    [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]::CreateToastNotifier($AppId).Show($XMLDocument)
    $ShutdownComment = "IT Notice: Automatic Restart in $Minutes minute(s). Save your work now."
    & shutdown /r /t $Seconds /c $ShutdownComment /d p:5:19
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
