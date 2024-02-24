param($distroList, $csv)

$updatedUsers = @(Get-Content -Path $csv)

$currentUsers = @((Get-DistributionGroupMember -Identity $distroList).PrimarySmtpAddress)

$addUsers = @($updatedUsers | Where-Object {$currentUsers -NotContains $_})
$removeUsers = @($currentUsers | Where-Object {$updatedUsers -NotContains $_})

Write-Host "`nUsers to be Added:"
$addUsers
Write-Host "`nUsers to be Removed:"
$removeUsers

$confirmation = Read-Host "`nDo you wish to continue with this process?? (Y/N)"

if($confirmation -eq "Y" -or $confirmation -eq "y" -or $confirmation -eq "Yes" -or $confirmation -eq "yes" -or $confirmation -eq "YES") {
    if([string]::IsNullOrEmpty($addUsers)) {
        Write-Host "No users to add"
    } else {
        foreach ($member in $addUsers) {
            Add-DistributionGroupMember -Identity $distro -Member $member
            Write-Host "Added" $member "to" $distroList "Distribution List"
        }
    }

    if([string]::IsNullOrEmpty($removeUsers)) {
        Write-Host "No users to remove"
    } else {
        foreach ($member in $removeUsers) {
            Remove-DistributionGroupMember -Identity $distro -Member $member
            Write-Host "Removed" $member "from" $distroList "Distribution List"
        }
    }
} else {
    Write-Host "No Actions Taken"
    break
}
