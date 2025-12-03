$RegistryPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer'
$Name = 'DisableSearchBoxSuggestions'
$Value = '1'

If (!(Test-Path $RegistryPath)) {
    New-Item -Path $RegistryPath -Force
}

New-ItemProperty -Path $RegistryPath -Name $Name -Value $Value -PropertyType DWORD -Force