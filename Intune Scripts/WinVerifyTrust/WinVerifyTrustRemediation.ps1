# Defines the proper paths and name/value of the registry keys to set
$64BitRegistryPath = 'HKLM:\Software\Microsoft\Cryptography\Wintrust\Config'
$32BitRegistryPath = 'HKLM:\Software\WOW6432Node\Microsoft\Cryptography\Wintrust\Config'
$Name = 'EnableCertPaddingCheck'
$Value = '1'

# Checks to see if there is already a key in place
$64BitExists = (Get-ItemProperty $64BitRegistryPath).PSObject.Properties.Name -contains $Name
$32BitExists = (Get-ItemProperty $32BitRegistryPath).PSObject.Properties.Name -contains $Name

# If the key exists, reset it as a REG_DWORD, instead of a REG_SZ
# Else, create the new path and REG_DWORD key
If ($64BitExists) {
    Set-ItemProperty -Path $64BitRegistryPath -Name $Name -Value $Value -Type DWORD -Force
} else {
    New-Item -Path $64BitRegistryPath -Force
    New-ItemProperty -Path $64BitRegistryPath -Name $Name -Value $Value -PropertyType DWORD -Force
}

If ($32BitExists) {
    Set-ItemProperty -Path $32BitRegistryPath -Name $Name -Value $Value -Type DWORD -Force
} else {
    New-Item -Path $32BitRegistryPath -Force
    New-ItemProperty -Path $32BitRegistryPath -Name $Name -Value $Value -PropertyType DWORD -Force
}