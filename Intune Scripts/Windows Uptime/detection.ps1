$Threshold = 7
try {
    $OS = Get-CimInstance -ClassName Win32_OperatingSystem
    $LastBoot = $OS.LastBootUpTime
    $UpTime = (Get-Date) - $LastBoot
    if($UpTime.TotalDays -ge $Threshold) {
        exit 1
    } else {
        exit 0
    }
} catch {
    Write-Error "Failed to determine last reboot time: $($_.Exception.Message)"
    exit 1
}