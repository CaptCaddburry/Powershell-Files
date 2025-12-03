$RegistryPath = 'HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32'
$Name = '(Default)'
$Value = 'default'

If (!(Test-Path $RegistryPath)) {
    New-Item -Path $RegistryPath -Force
}

New-ItemProperty -Path $RegistryPath -Name $Name -Value $Value -PropertyType String -Force