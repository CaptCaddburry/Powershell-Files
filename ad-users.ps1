Clear-Host

#If parameter is empty, then outputs the nothing was found message, else it will output the specific accounts
function CheckAccounts($accounts) {
    if([string]::IsNullOrEmpty($accounts)) {
        Write-Host "`n*****************************"
        Write-Host "*                           *"
        Write-Host "*     Nothing was found     *"
        Write-Host "*                           *"
        Write-Host "*****************************`n"
    } else {
        $accounts
    }
}

$CurrentTime = Get-Date -UFormat "%m/%d/%Y %r"
$PriorDays = (Get-Date).AddDays(-91)
#Checks to see if any account has a password that expired within the past 91 days (within the last quarter of a year)
$AccountPasswords = Get-ADUser -Filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} -Properties "DisplayName", "SamAccountName", "msDS-UserPasswordExpiryTimeComputed" | Select-Object -Property "DisplayName", "SamAccountName", @{Name="Expiration Date";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}} | Where-Object {$_."Expiration Date" -lt $CurrentTime -and $_."Expiration Date" -gt $PriorDays} | Sort-Object "DisplayName" | Format-Table

#Checks to see if any account is locked out
$AccountLockouts = Search-ADAccount -LockedOut | Select-Object Name, SamAccountName, lockedOut | Sort-Object Name | Format-Table
#Outputs the results through custom function
Write-Host "Accounts with Expired Passwords:" -ForegroundColor DarkYellow
CheckAccounts($AccountPasswords)
Write-Host "Locked Out Accounts:" -ForegroundColor DarkYellow
CheckAccounts($AccountLockouts)
