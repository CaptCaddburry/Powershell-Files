Clear-Host

function CheckAccounts($accounts) {
    $nothing_message = @("`n*****************************", "*                           *", "*     Nothing was found     *", "*                           *", "*****************************`n")
    if([string]::IsNullOrEmpty($accounts)) {
        foreach ($member in $nothing_message) {
            Write-Host $member
        }
    } else {
        $accounts
    }
}

$CurrentTime = Get-Date
$PriorDays = (Get-Date).AddDays(-91)
$AccountPasswords = Get-ADUser -Filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} -Properties "DisplayName", "SamAccountName", "msDS-UserPasswordExpiryTimeComputed" | Select-Object -Property "DisplayName", "SamAccountName", @{Name="Expiration Date";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}} | Where-Object {$_."Expiration Date" -lt $CurrentTime -and $_."Expiration Date" -gt $PriorDays} | Sort-Object "DisplayName" | Format-Table
$AccountLockouts = Get-ADUser -Filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} -Properties "DisplayName", "SamAccountName", "lockedOut", "lockoutTime" | Select-Object -Property "DisplayName", "SamAccountName", "lockedOut", @{Name="Lockout Time";Expression={[datetime]::FromFileTime($_."lockoutTime")}} | Where-Object {$_."lockedOut" -eq $True} | Sort-Object "DisplayName" | Format-Table -Property "DisplayName", "SamAccountName", "Lockout Time"

Write-Host "Accounts with Expired Passwords:" -ForegroundColor Yellow
CheckAccounts($AccountPasswords)
Write-Host "Locked Out Accounts:" -ForegroundColor Yellow
CheckAccounts($AccountLockouts)
