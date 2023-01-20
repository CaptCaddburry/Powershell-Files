param($vpn)
if($vpn -eq "Marlette") {
    Get-VpnConnection -Name "Marlette VPN Server" -ErrorAction SilentlyContinue | Remove-VpnConnection -Force -ErrorAction SilentlyContinue
    Add-VpnConnection -Name "Marlette VPN Server" -ServerAddress "vpn.marlettefunding.com" -TunnelType L2tp -L2tpPsk '5steJe&BUzZ!s#%a' -EncryptionLevel NoEncryption -AuthenticationMethod Pap -AllUserConnection -Force 
}
if($vpn -eq "Card") {
    Get-VpnConnection -Name "Marlette Card VPN" -ErrorAction SilentlyContinue | Remove-VpnConnection -Force -ErrorAction SilentlyContinue
    Add-VpnConnection -Name "Marlette Card VPN" -ServerAddress "vpn2.marlettefunding.com" -TunnelType L2tp -L2tpPsk '3BPQeJEm5QA7qnd!@Xu%' -EncryptionLevel NoEncryption -AuthenticationMethod Pap -AllUserConnection -Force
    Set-VpnConnectionIPsecConfiguration -ConnectionName "Marlette Card VPN" -AuthenticationTransformConstants SHA196 -CipherTransformConstants AES128 -EncryptionMethod AES128 -DHGroup Group14 -PfsGroup None -IntegrityCheckMethod SHA1 -Force
}
else {
    Write-Host("Incorrect VPN Name")
}