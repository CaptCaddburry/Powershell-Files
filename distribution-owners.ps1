param($dl)
Get-DistributionGroup -Identity $dl | Format-List -Property @{Name="Distribution List";Expression={$_.displayname}}, @{Name="Owners";Expression={$_.managedby}}