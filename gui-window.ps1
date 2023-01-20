Add-Type -AssemblyName System.Windows.Forms

#---------------------Variables---------------------

$labelWidth = 25
$labelHeight = 10
$displayFont = 'Microsoft Sans Serif, 10'
$buttonWidth = 60
$buttonHeight = 30

#---------------------Functions---------------------

function Check_Password {
    $username = $TextBox.Text
    $lockedOut = (get-aduser $username -Properties "LockedOut")."LockedOut"
    $passwordChanged = (get-aduser $username -Properties passwordlastset).PasswordLastSet
    $tempNewPassword = (get-aduser $username -Properties "msDS-UserPasswordExpiryTimeComputed")."msDS-UserPasswordExpiryTimeComputed"
    $newPassword = ([datetime]::FromFileTime($tempNewPassword))
    $DisplayUsername.Text = "Username: " + $username
    $UserLock.Text = "Locked Out: " + $lockedOut
    $ChangedPassword.Text = "Password Last Set: " + $passwordChanged
    $CurrentPassword.Text = "Password Expires: " + $newPassword
    $TextBox.Text = ""
}

function Clear_Window {
    $TextBox.Text = ""
    $DisplayUsername.Text = ""
    $UserLock.Text = ""
    $ChangedPassword.Text = ""
    $CurrentPassword.Text = ""
}

function Close_Window {
    $PasswordGUI.Close()
}

#---------------------GUI Window---------------------

$PasswordGUI = New-Object System.Windows.Forms.Form
$PasswordGUI.ClientSize = '300,300'
$PasswordGUI.text = "Test Project - GUI Example"
$PasswordGUI.BackColor = "#ffffff"

$Title = New-Object system.Windows.Forms.Label
$Title.text = "Check Password Expiration"
$Title.AutoSize = $true
$Title.width = $labelWidth
$Title.height = $labelHeight
$Title.location = New-Object System.Drawing.Point(20,20)
$Title.Font = 'Microsoft Sans Serif, 13'

$TextBox = New-Object system.Windows.Forms.TextBox
$TextBox.multiline = $false
$TextBox.width = 250
$TextBox.height = 15
$TextBox.location = New-Object System.Drawing.Point(25,50)
$TextBox.Font = $displayFont
$TextBox.Add_KeyDown({if($_.KeyCode -eq "Enter"){Check_Password}})
$TextBox.Add_KeyDown({if($_.KeyCode -eq "Escape"){Close_Window}})

$DisplayUsername = New-Object System.Windows.Forms.Label
$DisplayUsername.AutoSize = $true
$DisplayUsername.width = $labelWidth
$DisplayUsername.height = $labelHeight
$DisplayUsername.location = New-Object System.Drawing.Point(20,90)
$DisplayUsername.Font = $displayFont

$UserLock = New-Object System.Windows.Forms.Label
$UserLock.AutoSize = $true
$UserLock.width = $labelWidth
$UserLock.height = $labelHeight
$UserLock.location = New-Object System.Drawing.Point(20,120)
$UserLock.Font = $displayFont

$ChangedPassword = New-Object System.Windows.Forms.Label
$ChangedPassword.AutoSize = $true
$ChangedPassword.width = $labelWidth
$ChangedPassword.height = $labelHeight
$ChangedPassword.location = New-Object System.Drawing.Point(20,150)
$ChangedPassword.Font = $displayFont

$CurrentPassword = New-Object System.Windows.Forms.Label
$CurrentPassword.AutoSize = $true
$CurrentPassword.width = $labelWidth
$CurrentPassword.height = $labelHeight
$CurrentPassword.location = New-Object System.Drawing.Point(20,180)
$CurrentPassword.Font = $displayFont

$SubmitButton = New-Object system.Windows.Forms.Button
$SubmitButton.BackColor = "#012456"
$SubmitButton.text = "Submit"
$SubmitButton.width = $buttonWidth
$SubmitButton.height = $buttonHeight
$SubmitButton.location = New-Object System.Drawing.Point(35,250)
$SubmitButton.Font = $displayFont
$SubmitButton.ForeColor = "#ffffff"
$SubmitButton.Add_Click({ Check_Password })

$ClearButton = New-Object system.Windows.Forms.Button
$ClearButton.BackColor = "#b0b0b0"
$ClearButton.text = "Clear"
$ClearButton.width = $buttonWidth
$ClearButton.height = $buttonHeight
$ClearButton.location = New-Object System.Drawing.Point(115,250)
$ClearButton.Font = $displayFont
$ClearButton.ForeColor = "#ffffff"
$ClearButton.Add_Click({ Clear_Window })

$ExitButton = New-Object System.Windows.Forms.Button
$ExitButton.BackColor = "#b0b0b0"
$ExitButton.text = "Close"
$ExitButton.width = $buttonWidth
$ExitButton.height = $buttonHeight
$ExitButton.location = New-Object System.Drawing.Point(195,250)
$ExitButton.Font = $displayFont
$ExitButton.ForeColor = "#ffffff"
$ExitButton.Add_Click({ Close_Window })

#This line adds each item to the GUI Window
$PasswordGUI.controls.AddRange(@($Title,$TextBox,$DisplayUsername,$UserLock,$ChangedPassword,$CurrentPassword,$SubmitButton,$ClearButton,$ExitButton))

#This line displays the GUI Window
[void]$PasswordGUI.ShowDialog()