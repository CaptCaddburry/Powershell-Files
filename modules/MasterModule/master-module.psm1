# This will make it so anytime you use a command that includes '-AD' and has a Server Parameter, it will be set to the supplied address
# Currently, this is set to the FQDN of one of the two domain controllers
# The FQDN of the domain controllers may change from time to time, so always double check what's being used currently
#
$PSDefaultParameterValues['*-AD*:Server'] = {SERVER_NAME}

function Get-AssignedDistros(
    [Parameter(Mandatory)] $Username) {
        <#
        .SYNOPSIS
        Display all distribution lists that a user is assigned to.

        .DESCRIPTION
        This will list all the distribution lists that the specified user is a member of.
        You must first be connected to the server by using the Connect-ExchangeOnline command.

        .PARAMETER Username
        Specifies the user's email address

        .INPUTS
        PS> Get-AssignedDistros -Username "james.cadd@domain.com"

        .OUTPUTS
        Username: james.cadd@domain.com

        DisplayName                 PrimarySmtpAddress
        -----------                 ------------------
        Domain Employee             domain-employee@domain.com
        Desk Egg Squad              desk-egg-squad-dl@domain.com
        General Employee Population general-employee-population@domain.com
        Prey Alerts                 prey-alerts@domain.com
        tech-org                    tech-org@domain.com
        #>

        $DistroResults = Get-DistributionGroup | Where-Object {(Get-DistributionGroupMember $_.Name | ForEach-Object {$_.PrimarySmtpAddress}) -contains "$Username"} | Format-Table DisplayName, PrimarySmtpAddress
        Clear-Host
        Write-Host "Username: " -NoNewline
        Write-Host $Username -ForegroundColor Yellow
        $DistroResults
}

function Get-AssignedShared(
    [Parameter(Mandatory)] $Username) {
        <#
        .SYNOPSIS
        Display all shared mailboxes that a user is assigned to.

        .DESCRIPTION
        This will list all the shared mailboxes that the specified user is a member of.
        You must first be connected to the server by using the Connect-ExchangeOnline command.

        .PARAMETER Username
        Specifies the user's email address.

        .INPUTS
        PS> Get-AssignedShared -Username "james.cadd@domain.com"

        .OUTPUTS
        Username: james.cadd@domain.com

        DisplayName         PrimarySmtpAddress
        -----------         ------------------
        Desk Egg Squad Team desk-egg-squad@domain.com
        #>

        $SharedResults = Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission -User $Username | Where-Object {$_.AccessRights -match "FullAccess" -and $_.IsInherited -eq $False} | ForEach-Object {Get-Mailbox $_.Identity | Select-Object DisplayName, PrimarySmtpAddress}
        Clear-Host
        Write-Host "Username: " -NoNewline
        Write-Host $Username -ForegroundColor Yellow
        $SharedResults
}

function Get-DistroMembers(
    [Parameter(Mandatory)] $DistributionList) {
        <#
        .SYNOPSIS
        Display all members of a specified distribution list.

        .DESCRIPTION
        This will list all the email addresses of every member within a specified distribution list.
        You must first be connected to the server by using the Connect-ExchangeOnline command.

        .PARAMETER DistributionList
        Specifies the distribution list's email address

        .INPUTS
        PS> Get-DistroMembers -DistributionList "desk-egg-squad-dl@domain.com"

        .OUTPUTS
        Distribution List Members:

        james.cadd@domain.com
        colin.turner@domain.com
        tommy.nguyen@domain.com
        #>

        Write-Host "Distribution List Members:`n"
        (Get-DistributionGroupMember -Identity $DistributionList).PrimarySmtpAddress
}

function Get-SharedMailbox(
    [Parameter(Mandatory)] $SharedMailbox) {
        <#
        .SYNOPSIS
        Display all members of a specified shared mailbox.

        .DESCRIPTION
        This will list every user and their access rights to a specified shared mailbox.
        You must first be connected to the server by using the Connect-ExchangeOnline command.

        .PARAMETER SharedMailbox
        Specifies the shared mailbox's email address

        .INPUTS
        PS> Get-SharedMailbox -SharedMailbox "desk-egg-squad@domain.com"

        .OUTPUTS
        Identity            User                     AccessRights
        --------            ----                     ------------
        Desk Egg Squad Team james.cadd@domain.com   {FullAccess}
        Desk Egg Squad Team colin.turner@domain.com {FullAccess}
        #>

        Get-MailboxPermission -Identity $SharedMailbox | Sort-Object User | Format-Table Identity, User, AccessRights
}

function Get-EmailTrace(
    [Parameter(Mandatory)] $User,
    [Parameter(Mandatory)] $StartTime,
    [Parameter(Mandatory)] $EndTime) {
        <#
        .SYNOPSIS
        Display an email trace for a specific email address, between a specific time frame

        .DESCRIPTION
        This function will allow you to run an email trace for the specified email address on the Exchange Server.
        You must first be connected to the server by using the Connect-ExchangeOnline command.

        .PARAMETER User
        Specifies the User's email address

        .PARAMETER StartTime
        Specifies the start date of the trace in MM/DD/YYYY format

        .PARAMETER EndTime
        Specifies the end date of the trace in MM/DD/YYYY format

        .INPUTS
        PS> Get-EmailTrace -User james.cadd@domain.com -StartTime 07/01/2024 -EndTime 07/02/2024

        .OUTPUTS
        Recipient Address: james.cadd@domain.com

        Received           SenderAddress              Subject                               Status
        --------           -------------              -------                               ------
        7/8/2024 7:57:29PM jira@domain.atlassian.net [JIRA] (ITS-123456) My Computer Broke Delivered
        #>

        Clear-Host

        Write-Host "Recipient Address: " -NoNewline 
        Write-Host $User -ForegroundColor Yellow
        Get-MessageTrace -RecipientAddress $User -StartDate $StartTime -EndDate $EndTime | Format-Table Received, SenderAddress, Subject, Status
}

function Get-DynamicDistroMembers(
    [Parameter(Mandatory)] $DirectManager) {
        <#
        .SYNOPSIS
        List all members that are included in a dynamic distribution list.

        .DESCRIPTION
        This will look up every user in the Exchange server that has the specified user listed as their manager.
        Everyone showing in this list will be added to the dynamic distribution list.
        If a user isn't showing up in this list when they should be, run Get-ADUser and check thier Manager parameter.
        You must first be connected to the server by using the Connect-ExchangeOnline command.

        .PARAMETER DirectManager
        Specifies the manager's AD username of the dynamic distribution list

        .INPUTS
        PS> Get-DynamicDistroMembers -DirectManager "james.cadd"

        .OUTPUTS
        Dynamic Distribution Name: James Cadd Direct Reports
        Dynamic Distribution Email: jamescadd-directs@domain.com

        Name         PrimarySmtpAddress
        ----         ------------------
        James Cadd   james.cadd@domain.com
        Colin Turner colin.turner@domain.com
        #>

        $ManagerName = ((Get-Culture).TextInfo.ToTitleCase($DirectManager)).replace("."," ")
        $EmailManagerName = $DirectManager.replace(".","")
        $DistributionName = $ManagerName + " Direct Reports"
        $DistributionEmail = $EmailManagerName + "-directs@domain.com"
        $DistinguishedName = (Get-Recipient -Identity $DirectManager).DistinguishedName
        $Filter = "Manager -eq '$DistinguishedName'"

        Clear-Host
        Write-Host "Dynamic Distrubition Name: " -NoNewline
        Write-Host $DistributionName -ForegroundColor Yellow
        Write-Host "Dynamic Distribution Email: "-NoNewline
        Write-Host $DistributionEmail -ForegroundColor Yellow
        Get-Recipient -Filter $Filter | Format-Table Name, PrimarySmtpAddress
}

function New-DynamicDistro(
    [Parameter(Mandatory)] $DirectManager) {
        <#
        .SYNOPSIS
        Creates a new dynamic distribution list, based on the specified user.

        .DESCRIPTION
        This will create a new dynamic distribution list for the specified user.
        Once a day, the server will check every user's manager parameter and add/remove them accordingly.
        You must first be connected to the server by using the Connect-ExchangeOnline command.

        .PARAMETER DirectManager
        Specifies the manager's AD username of the dynamic distribution list

        .INPUTS
        PS> New-DynamicDistro -DirectManager "james.cadd"

        .OUTPUTS
        As of now, nothing is outputed.
        The dynamic distribution list gets created.
        #>

        $ManagerName = ((Get-Culture).TextInfo.ToTitleCase($DirectManager)).replace("."," ")
        $EmailManagerName = $DirectManager.replace(".","")
        $DistributionName = $ManagerName + " Direct Reports"
        $DistributionEmail = $EmailManagerName + "-directs@domain.com"
        $DistinguishedName = (Get-Recipient -Identity $DirectManager).DistinguishedName
        $Filter = "Manager -eq '$DistinguishedName'"

        New-DynamicDistributionGroup -Name $DistributionName -PrimarySmtpAddress $DistributionEmail -RecipientFilter $Filter
}

function Get-UserInfo(
    [Parameter(Mandatory)] $User) {
        <#
        .SYNOPSIS
        List common use information about a specified user.

        .DESCRIPTION
        This will output basic account information about a given user.
        Locked out status, when their password was last changed, when their current password expires, the OU folder directory they are hosted under, and their employee number.

        .PARAMETER User
        Specifies the user you are looking up.

        .INPUTS
        PS> Get-UserInfo -User "james.cadd"

        .OUTPUTS
        *******************************************************
        Username: james.cadd
        Current Time: 07/08/2024 10:00:00 AM
        User Locked Out: False
        Password Last Set: 6/4/2024 08:00:00 AM
        Password Expires: 9/2/2024 08:00:00 AM
        Location: Technology
        Employee Number: 1694
        *******************************************************
        #>

        $currentTime = Get-Date -UFormat "%m/%d/%Y %r"
        $lockedOut = (Get-ADUser $User -Properties "LockedOut")."LockedOut"
        $passwordChanged = (Get-ADUser $User -Properties passwordlastset).PasswordLastSet
        $tempNewPassword = (Get-ADUser $User -Properties "msDS-UserPasswordExpiryTimeComputed")."msDS-UserPasswordExpiryTimeComputed"
        $newPassword = ([datetime]::FromFileTime($tempNewPassword))
        $adName = (Get-ADUser $User -Properties Name).Name
        $treeLocation = (Get-ADUser $User -Properties CanonicalName).CanonicalName
        $ouIndex = $treeLocation.IndexOf('Active-Users/')
        $ouLocation = (($treeLocation.substring($ouIndex)).Replace('Active-Users/', '')).Replace('/' + $adName, '')
        $employeeID = (Get-ADUser $User -Properties employeeNumber).employeeNumber
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
        Write-Host "Employee Number:" $employeeID
        Write-Host "*******************************************************"
}

function Get-ADLockouts(
    [Parameter(Mandatory=$False)] $DayCount = -91) {
        <#
        .SYNOPSIS
        List every issue user in the AD server.

        .DESCRIPTION
        This function will do three things in total:
        1. List every user with an expired password between today and however many days are specified.
        2. List every user that is currently locked out.
        3. List any duplicate user accounts.

        If any of the values show up empty, the "Nothing was found" message will be listed instead
        Example:

        Locked Out Accounts:
        
        *****************************
        *                           *
        *     Nothing was found     *
        *                           *
        *****************************

        .PARAMETER DayCount
        Specifies how many days back the script should filter out.
        This is not a required parameter.
        By default, this will be set to 91 days ago (the previous rolling quarter).
        If you are changing this parameter, make sure to make it a negative number.

        .INPUTS
        PS> Get-ADLockouts
        PS> Get-ADLockouts -DayCount -14

        .OUTPUTS
        Accounts with Expired Passwords:
        
        DisplayName  SamAccountName Expiration Date
        -----------  -------------- ---------------
        James Cadd   james.cadd     6/4/2024 08:00:00 AM
        Colin Turner colin.turner   6/7/2024 08:00:00 AM
        

        Locked Out Accounts:
        
        DisplayName  SamAccountName LockoutTime
        -----------  -------------- -----------
        James Cadd   james.cadd     6/4/2024 08:00:00 AM
        Colin Turner colin.turner   6/5/2024 08:00:00 AM
        

        Duplicate Accounts:
        
        DisplayName  SamAccountName WhenCreated          LastLogon
        -----------  -------------- -----------          ---------
        James Cadd   james.cadd1    6/3/2024 08:00:00 AM         0
        Colin Turner colin.turner1  6/5/2024 08:00:00 AM         0
        #>
        
        Write-Host "Checking accounts..."

        $CurrentTime = Get-Date
        $PreviousDays = (Get-Date).AddDays($DayCount)
        $AccountPasswords = Get-ADUser -Filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} -Properties "DisplayName", "SamAccountName", "msDS-UserPasswordExpiryTimeComputed" | Select-Object -Property "DisplayName", "SamAccountName", @{Name="Expiration Date";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}} | Where-Object {$_."Expiration Date" -lt $CurrentTime -and $_."Expiration Date" -gt $PreviousDays} | Sort-Object "DisplayName" | Format-Table
        $AccountLockouts = Get-ADUser -Filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} -Properties "DisplayName", "SamAccountName", "lockedOut", "lockoutTime" | Select-Object -Property "DisplayName", "SamAccountName", "lockedOut", @{Name="Lockout Time";Expression={[datetime]::FromFileTime($_."lockoutTime")}} | Where-Object {$_."lockedOut" -eq $True} | Sort-Object "DisplayName" | Format-Table -Property "DisplayName", "SamAccountName", "Lockout Time"
        $AccountDuplicates = Get-ADUser -Filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} -Properties "DisplayName", "SamAccountName", "WhenCreated", "LastLogon" | Where-Object "SamAccountName" -like "*[0-9]" | Sort-Object "SamAccountName" | Format-Table -Property "DisplayName", "SamAccountName", "WhenCreated", "LastLogon"

        Write-Host "Compiling lists..."
        Start-Sleep -Seconds 3
        Clear-Host
        Start-Sleep -Seconds 2
        Write-Host "Accounts with Expired Passwords:" -ForegroundColor Yellow
        CheckString $AccountPasswords
        Write-Host "Locked Out Accounts:" -ForegroundColor Yellow
        CheckString $AccountLockouts
        Write-Host "Duplicate Accounts:" -ForegroundColor Yellow
        CheckString $AccountDuplicates

        Write-Host "Current Time:" $CurrentTime -ForegroundColor Yellow
}

function UpdateDistros(
    [Parameter(Mandatory)] $DistributionList,
    [Parameter(Mandatory)] $CSV) {
        <#
        .SYNOPSIS
        Replace a current list of users in a distribution list with an updated list of members.

        .DESCRIPTION
        This will take a supplied CSV and compare every user listed to every user in the specified distribution list.
        Every user that is in the CSV but not in the distribution list will be added.
        Every user that is in the distribution list but not in the CSV will be removed.
        If a user shows up in both the CSV and distribution list, nothing will happen to them.
        The function will output all the changes to be made and ask for confirmation before performing the task.
        If the distribution list doesn't exist, the function will exit.
        You must first be connected to the server by using the Connect-ExchangeOnline command.

        .PARAMETER DistributionList
        Specifies the distribution list to edit.

        .PARAMETER CSV
        Specifies the location of the CSV.
        The CSV file should only contain 1 column of user's email addresses, with no column header.

        .INPUTS
        PS> UpdateDistros -DistributionList "desk-egg-squad-dl@domain.com" -CSV "users.csv"

        .OUTPUTS
        Identity: Desk Egg Squad Team
        Email Address: desk-egg-squad-dl@domain.com

        Users to be Added:
        james.cadd@domain.com
        colin.turner@domain.com
        tommy.nguyen@domain.com

        Users to be Removed:
        alexis.heritage@domain.com
        melissa.canfield@domain.com

        Do you wish to continue with this process?? (Y/N)
        #>

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

function Get-OutlookRetentionReport(
    [Parameter(Mandatory)] $ExportLocation) {
        <#
        .SYNOPSIS
        List specific Outlook mailbox statistics for our retention reports.

        .DESCRIPTION
        This will look over EVERY mailbox in our Exchange server and list specific information about the number of items and folders that were archived, due to our retention policies.
        All the information will be stored in a CSV file.
        As this is looking over every mailbox, this will take quite a bit of time.
        You must first be connected to the server by using the Connect-ExchangeOnline command.

        .PARAMETER ExportLocation
        Specifies the location of the CSV export.

        .INPUTS
        PS> Get-OutlookRetentionReport -ExportLocation "report.csv"

        .OUTPUTS
        The stats will be output into the CSV file specified.
        #>

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

# Utility function used in Get-ADLockouts
# If $Accounts is empty or null, output $nothing_message
# If not empty, output $Accounts
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
