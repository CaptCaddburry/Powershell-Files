param($distroList, $csv)

$newUsers = @(Get-Content -Path $csv)

$currentUsers = @((Get-DistributionGroupMember -Identity $distroList).PrimarySmtpAddress)

$addUsers = @($newUsers | Where-Object {$currentUsers -NotContains $_})
$removeUsers = @($currentUsers | Where-Object {$newUsers -NotContains $_})

foreach ($member in $addUsers) {
    Add-DistributionGroupMember -Identity $distro -Member $member
    Write-Host "Added" $member "to" $distro "Distribution List"
}

foreach ($member in $removeUsers) {
    Remove-DistributionGroupMember -Identity $distro -Member $member
    Write-Host "Removed" $member "from" $distro "Distribution List"
}
