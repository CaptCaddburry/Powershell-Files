param($dl, $filename)
$fileLocation = "{LOCATION}" + $filename + ".csv"
Get-DistributionGroupMember -Identity $dl | Sort-Object manager, name | Select-Object -Property @{Name="Member";Expression={$_.name}}, Title, Manager, @{Name="Email Address";Expression={$_.primarysmtpaddress}} | Export-Csv -Path $fileLocation -NoTypeInformation
