If ($ENV:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    Try {
        &amp;"$ENV:WINDIR\SysNative\WindowsPowershell\v1.0\PowerShell.exe" -File $PSCOMMANDPATH
    }
    Catch {
        Throw "Failed to start $PSCOMMANDPATH"
    }
    Exit
}

$configPathUninstall = Get-ChildItem -Path "C:\Windows\IMECache\*\visioConfigUninstall.xml" -Recurse  | ForEach-Object{$_.FullName}
$officeSetupPath = Get-ChildItem -Path "C:\Windows\IMECache\*\VisioInstaller.exe" -Recurse  | ForEach-Object{$_.FullName}

if (Test-Path -Path "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE" -PathType Leaf) {
    C:\Windows\IMECache\*\ServiceUI.exe -process:explorer.exe $($officeSetupPath) /configure $($configPathUninstall)
} elseif (Test-Path -Path "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE" -PathType Leaf) {
    C:\Windows\IMECache\*\ServiceUI.exe -process:explorer.exe $($officeSetupPath) /configure $($configPathUninstall)
} else {
    C:\Windows\IMECache\*\ServiceUI.exe -process:explorer.exe $($officeSetupPath) /configure $($configPathUninstall)
}