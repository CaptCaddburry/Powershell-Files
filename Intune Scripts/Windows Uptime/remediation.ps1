$Minutes = 15
$Seconds = [Math]::Max(60, $Minutes * 60)
$ETA = (Get-Date).AddSeconds($Seconds).ToString("hh:mm tt")

try {
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
    Write-Output "Reboot scheduled in $Minutes minute(s) (ETA $ETA). A system shutdown notification has been displayed."
} catch {
    Write-Error "Failed to schedule reboot: $($_.Exception.Message)"
    exit 1
}
