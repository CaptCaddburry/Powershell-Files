$RegistryKeys = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
$M365Apps = "Microsoft Visio"
$M365AppsCheck = $RegistryKeys | Where-Object { $_.GetValue("DisplayName") -match $M365Apps }
if ($M365AppsCheck) {
    Write-Output "Visio Detected"
	Exit 0
   } else {
    Exit 1
}