Clear-Host
Write-Host "Checking for a password expiration, input exit to close"
Write-Host "*******************************************************" -ForegroundColor Green

While ($true) {
    $username = Read-Host -Prompt "Enter a username"
    if ($username -eq 'exit') {
        Break
    }
    #$manager = (get-aduser $username -Properties reportsTo).reportsTo
    $currentTime = Get-Date -UFormat "%m/%d/%Y %r"
    $lockedOut = (get-aduser $username -Properties "LockedOut")."LockedOut"
    $passwordChanged = (get-aduser $username -Properties passwordlastset).PasswordLastSet
    $tempNewPassword = (get-aduser $username -Properties "msDS-UserPasswordExpiryTimeComputed")."msDS-UserPasswordExpiryTimeComputed"
    $newPassword = ([datetime]::FromFileTime($tempNewPassword))
        #Write-Host "User's Manager:" $manager -ForegroundColor Yellow
    Write-Host "Current Time:" $currentTime
    if ($lockedOut -eq "True") {
        Write-Host "User Locked Out:" $lockedOut -ForegroundColor Red
    } else {
        Write-Host "User Locked Out:" $lockedOut -ForegroundColor Yellow
    }
    Write-Host "Password Last Set:" $passwordChanged
    Write-Host "Password Expires:" $newPassword
    Write-Host "*******************************************************" -ForegroundColor Green
}