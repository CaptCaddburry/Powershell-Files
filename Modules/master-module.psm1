function Get-Distro($DistributionList) {
    (Get-DistributionGroupMember -Identity $DistributionList).PrimarySmtpAddress
}

function Get-SharedMailbox($SharedMailbox) {
    Get-MailboxPermission -Identity $SharedMailbox | Sort-Object User | Format-Table Identity, User, AccessRights
}

function Get-UserInfo($User) {
    $currentTime = Get-Date -UFormat "%m/%d/%Y %r"
    $lockedOut = (get-aduser $User -Properties "LockedOut")."LockedOut"
    $passwordChanged = (get-aduser $User -Properties passwordlastset).PasswordLastSet
    $tempNewPassword = (get-aduser $User -Properties "msDS-UserPasswordExpiryTimeComputed")."msDS-UserPasswordExpiryTimeComputed"
    $newPassword = ([datetime]::FromFileTime($tempNewPassword))
    $adName = (get-aduser $User -Properties Name).Name
    $treeLocation = (get-aduser $User -Properties CanonicalName).CanonicalName
    $ouIndex = $treeLocation.IndexOf('Active-Users/')
    $ouLocation = (($treeLocation.substring($ouIndex)).Replace('Active-Users/', '')).Replace('/' + $adName, '')
    Write-Host "`nCurrent Time:" $currentTime
    if ($lockedOut -eq "True") {
        Write-Host "User Locked Out:" $lockedOut -ForegroundColor Red
    } else {
        Write-Host "User Locked Out:" $lockedOut -ForegroundColor Yellow
    }
    Write-Host "Password Last Set:" $passwordChanged
    if($newPassword -lt $currentTime) {
        Write-Host "Password Expires:" $newPassword -ForegroundColor Red
    } else {
        Write-Host "Password Expires:" $newPassword
    }
    Write-Host "Location: " $ouLocation
}

function Get-ADLockouts {
    Clear-Host

    $CurrentTime = Get-Date
    $PreviousQuarter = (Get-Date).AddDays(-91)
    $AccountPasswords = Get-ADUser -Filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} -Properties "DisplayName", "SamAccountName", "msDS-UserPasswordExpiryTimeComputed" | Select-Object -Property "DisplayName", "SamAccountName", @{Name="Expiration Date";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}} | Where-Object {$_."Expiration Date" -lt $CurrentTime -and $_."Expiration Date" -gt $PreviousQuarter} | Sort-Object "DisplayName" | Format-Table
    $AccountLockouts = Get-ADUser -Filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} -Properties "DisplayName", "SamAccountName", "lockedOut", "lockoutTime" | Select-Object -Property "DisplayName", "SamAccountName", "lockedOut", @{Name="Lockout Time";Expression={[datetime]::FromFileTime($_."lockoutTime")}} | Where-Object {$_."lockedOut" -eq $True} | Sort-Object "DisplayName" | Format-Table -Property "DisplayName", "SamAccountName", "Lockout Time"
    $AccountDuplicates = Get-ADUser -Filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} -Properties "DisplayName", "SamAccountName", "WhenCreated", "LastLogon" | Where-Object SamAccountName -like "*[0-9]" | Sort-Object "SamAccountName" | Format-Table -Property "DisplayName", "SamAccountName", "WhenCreated", "LastLogon"

    Write-Host "Accounts with Expired Passwords:" -ForegroundColor Yellow
    CheckString $AccountPasswords
    Write-Host "Locked Out Accounts:" -ForegroundColor Yellow
    CheckString $AccountLockouts
    Write-Host "Duplicate Accounts:" -ForegroundColor Yellow
    CheckString $AccountDuplicates

    Write-Host "Current Time:" $CurrentTime -ForegroundColor Yellow
}

function CheckString($accounts) {
    $nothing_message = @("`n*****************************", "*                           *", "*     Nothing was found     *", "*                           *", "*****************************`n")
    if([string]::IsNullOrEmpty($accounts)) {
        foreach ($member in $nothing_message) {
            Write-Host $member
        }
    } else {
        $accounts
    }
}
