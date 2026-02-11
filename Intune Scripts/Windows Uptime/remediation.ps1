$Minutes = 15
$Seconds = [Math]::Max(60, $Minutes * 60)
$ETA = (Get-Date).AddSeconds($Seconds).ToString("hh:mm tt")

try {
    $ShutdownComment = "IT Notice: Your computer needs to be restarted in order to install some needed security patches. Your computer will be restarted at $ETA. Please save your work now."
    & shutdown /r /t $Seconds /c $ShutdownComment /d p:5:19
    Write-Output "Reboot scheduled in $Minutes minute(s) (ETA $ETA). A system shutdown notification has been displayed."
} catch {
    Write-Error "Failed to schedule reboot: $($_.Exception.Message)"
    exit 1
}
