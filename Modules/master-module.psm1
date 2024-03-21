# This makes the default value of any command with -AD in the name to use the provided server address
#
$PSDefaultParameterValues['*-AD*:Server'] = '{SERVER_NAME}'

# This will output all the email addresses of users who are members of the specified distribution list
#
function Get-DistroMembers($DistributionList) {
    (Get-DistributionGroupMember -Identity $DistributionList).PrimarySmtpAddress
}

# This will output all users that have access to the specified shared mailbox
#
function Get-SharedMailbox($SharedMailbox) {
    Get-MailboxPermission -Identity $SharedMailbox | Sort-Object User | Format-Table Identity, User, AccessRights
}

# This will output a table of all emails sent to the specified user, during the specified time frame
#
function Get-EmailTrace($User, $StartTime, $EndTime) {
    Clear-Host

    Write-Host "Recipient Address:" $User
    Get-MessageTrace -RecipientAddress $User -StartDate $StartTime -EndDate $EndTime | Format-Table Received, SenderAddress, Subject, Status
}

# This will create a new dynamic distribution list, based on who their direct report is
#
function New-DynamicDistro($DirectManager) {
    $managerName = ((Get-Culture).TextInfo.ToTitleCase($DirectManager)).replace("."," ")
    $emailManagerName = $DirectManager.replace(".","")
    $DistributionName = $managerName + " Direct Reports"
    $DistributionEmail = $emailManagerName + "-directs@{DOMAIN}"
    $DistinguishedName = (Get-Recipient -Identity $DirectManager).DistinguishedName
    $Filter = "Manager -eq '$DistinguishedName'"

    New-DynamicDistributionGroup -Name $DistributionName -PrimarySmtpAddress $DistributionEmail -RecipientFilter $Filter
}

# This will pull information based on the specified user from the active directory server
#
function Get-UserInfo($User) {
    $currentTime = Get-Date -UFormat "%m/%d/%Y %r"
    $lockedOut = (Get-ADUser $User -Properties "LockedOut")."LockedOut"
    $passwordChanged = (Get-ADUser $User -Properties passwordlastset).PasswordLastSet
    $tempNewPassword = (Get-ADUser $User -Properties "msDS-UserPasswordExpiryTimeComputed")."msDS-UserPasswordExpiryTimeComputed"
    $newPassword = ([datetime]::FromFileTime($tempNewPassword))
    $adName = (Get-ADUser $User -Properties Name).Name
    $treeLocation = (Get-ADUser $User -Properties CanonicalName).CanonicalName
    $ouIndex = $treeLocation.IndexOf('Active-Users/')
    $ouLocation = (($treeLocation.substring($ouIndex)).Replace('Active-Users/', '')).Replace('/' + $adName, '')
    Write-Host "*******************************************************"
    Write-Host "Username:" $User
    Write-Host "Current Time:" $currentTime
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
    Write-Host "Location:" $ouLocation
    Write-Host "*******************************************************"
}

# This will output every account with an actively expired password from the past 91 days, every account that is actively locked out, and every duplicate account that ends with a number
#
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

# This will update the specified distribution list based on users stored in the specified CSV file
# Users that are in the file and not the distro will be added. Users that are in the file and the distro, nothing will happen to them. Users that are not in the file and are in the distro will be removed.
#
function UpdateDistros($DistributionList, $CSV) {
    $distroStatus = (Get-DistributionGroup -Identity $DistributionList).DisplayName

    if([string]::IsNullOrEmpty($distroStatus)) {
        Write-Host "The requested distribution list doesn't exist"
        Exit
    } else {
        Write-Host $distroStatus "exists. Continuing..."
        Start-Sleep -Seconds 2
    }

    $updatedUsers = @(Get-Content -Path $CSV)
    $currentUsers = @((Get-DistributionGroupMember -Identity $DistributionList).PrimarySmtpAddress)
    $addUsers = @($updatedUsers | Where-Object {$currentUsers -NotContains $_})
    $removeUsers = @($currentUsers | Where-Object {$updatedUsers -NotContains $_})

    Write-Host "`nIdentity:" $distroStatus
    Write-Host "Email Address:" $DistributionList
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
                Add-DistributionGroupMember -Identity $DistributionList -Member $member
                Write-Host "Added" $member "to" $distroStatus
            }
        }

        if([string]::IsNullOrEmpty($removeUsers)) {
            Write-Host "No users to remove"
        } else {
            foreach ($member in $removeUsers) {
                Remove-DistributionGroupMember -Identity $DistributionList -Member $member
                Write-Host "Removed" $member "from" $distroStatus
            }
        }
    } else {
        Write-Host "No Actions Taken"
        Exit
    }
}

# This will look through every mailbox in the exchange server and store all the specified stats for items that were archived, due to the retention policy
#
function Get-OutlookRetentionReport($ExportLocation) {
    $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited)
    $report = @()
    
    foreach ($mailbox in $mailboxes) {
        $inboxstats = Get-EXOMailboxFolderStatistics $mailbox.UserPrincipalName -FolderScope Archive | Where-Object {$_.FolderPath -eq "/Archive"}

        $mbObj = New-Object PSObject
        $mbObj | Add-Member -MemberType NoteProperty -Name "UPN" -Value $mailbox.UserPrincipalName
        $mbObj | Add-Member -MemberType NoteProperty -Name "Display Name" -Value $mailbox.DisplayName
        $mbObj | Add-Member -MemberType NoteProperty -Name "Folder" -Value $inboxstats.FolderPath
        $mbObj | Add-Member -MemberType NoteProperty -Name "Folder Size (MB)" -Value $inboxstats.FolderandSubFolderSize
        $mbObj | Add-Member -MemberType NoteProperty -Name "Number of Items" -Value $inboxstats.ItemsinFolderandSubfolders
        $report += $mbObj
    }
    $report | Export-CSV $ExportLocation
}

# Custom internal function used to verify if the specified variable is empty or not
#
function CheckString($Accounts) {
    $nothing_message = @("`n*****************************", "*                           *", "*     Nothing was found     *", "*                           *", "*****************************`n")
    if([string]::IsNullOrEmpty($Accounts)) {
        foreach ($member in $nothing_message) {
            Write-Host $member
        }
    } else {
        $Accounts
    }
}
