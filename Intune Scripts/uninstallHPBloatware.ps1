# Apps to be uninstalled silently
#
# HP Documentation
# HP Notifications
# HP PC Hardware Diagnostics
# HP Privacy Settings
# HP Security Update Service
# HP Sure Recover
# HP Sure Run Module
# HP Wolf Security
# HP Wolf Security - Console
# myHP
# 

# List of HP Applications installed as Appx Packages
$UninstalledPackages = @(
    "AD2F1837.HPPCHardwareDiagnosticsWindows"
    "AD2F1837.HPPrivacySettings"
    "AD2F1837.myHP"
)

# List of HP Applications installed as Programs
$UninstalledPrograms = @(
    "HP Notifications"
    "HP Security Update Service"
    "HP Sure Recover"
    "HP Sure Run Module"
    "HP Wolf Security"
    "HP Wolf Security - Console"
)

# Locates all installed applications on machine and stores them in distinct variables
$InstalledPackages = Get-AppxPackage -AllUsers | Where-Object {$UninstalledPackages -contains $_.Name}
$InstalledPrograms = Get-Package | Where-Object {$UninstalledPrograms -contains $_.Name}

# Uninstalls each Appx Package
ForEach ($AppX in $InstalledPackages) {
    Try {
        $Null = Remove-AppxPackage -Package $AppX.PackageFullName -AllUsers -ErrorAction Stop
    }
    Catch {}
}

# Uninstalls each program
$InstalledPrograms | ForEach-Object {
    Try {
        $Null = $_ | Uninstall-Package -AllVersions -Force -ErrorAction Stop
    }
    Catch {}
}

# Remove HP Documentation if it exists
if (Test-Path -Path "C:\Program Files\HP\Documentation\Doc_uninstall.cmd") {
    Start-Process -FilePath "C:\Program Files\HP\Documentation\Doc_uninstall.cmd" -Wait -passthru -NoNewWindow
}
