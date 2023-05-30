Add-Type -AssemblyName System.Windows.Forms

#---------------------Variables---------------------

$displayFont = 'Microsoft Sans Serif, 10'
$buttonWidth = 60
$buttonHeight = 30

#---------------------Functions---------------------

#This function runs the given username through the active directory and sets the values for UserLock, ChangedPassword, and CurrentPassword
function Check_Password {
    $UserLock.ForeColor = "#000000"
    $CurrentPassword.ForeColor = "#000000"
    $username = $TextBox.Text
    $currentTime = Get-Date -UFormat "%m/%d/%Y %r"
    $currentTimeCheck = Get-Date -UFormat "%r"
    $currentTimeLogFolder = Get-Date -UFormat "%m-%Y"
    $currentTimeLogFile = Get-Date -UFormat "%m-%d-%Y"
    $lockedOut = (get-aduser $username -Properties "LockedOut")."LockedOut"
    $passwordChanged = (get-aduser $username -Properties passwordlastset).PasswordLastSet
    $tempNewPassword = (get-aduser $username -Properties "msDS-UserPasswordExpiryTimeComputed")."msDS-UserPasswordExpiryTimeComputed"
    $newPassword = ([datetime]::FromFileTime($tempNewPassword))
    $newPasswordCheck = Get-Date -Date $newPassword -UFormat "%r"
    $DisplayUsername.Text = "Username: " + $username
    if($lockedOut -eq $true) {
        $UserLock.ForeColor = "#ff0000"
    }
    $TimeDisplay.Text = "Current Time: " + $currentTime
    $UserLock.Text = "Locked Out: " + $lockedOut
    $ChangedPassword.Text = "Password Last Set: " + $passwordChanged
    if(($currentTime -ge $newPassword) -and ($currentTimeCheck -ge $newPasswordCheck)) {
        $CurrentPassword.ForeColor = "#ff0000"
    }
    $CurrentPassword.Text = "Password Expires: " + $newPassword
    $TextBox.Text = ""
    #This section will append the requested information into a csv file at the specific location
    $logPath = "FILEPATH"
    New-Item -ItemType Directory -Force -Path $logPath -Name $currentTimeLogFolder
    $logFile = $logPath + $currentTimeLogFolder + "\Password_Log_" + $currentTimeLogFile + ".csv"
    $reportLog = @($username, $currentTimeCheck, $lockedOut, $passwordChanged, $newPassword)
    $reportName = @("Username", "Current Time", "Locked Out", "Password Last Set", "Password Expires")
    $userObj = New-Object PSObject
    for($i = 0; $i -lt $reportLog.Length; $i++) {
        $userObj | Add-Member -MemberType NoteProperty -Name $reportName[$i] -Value $reportLog[$i]
    }
    $userObj | Export-Csv -Path $logFile -NoTypeInformation -Append
}

#This function is to clear out the previous search
function Clear_Window {
    $TextBox.Text = ""
    $DisplayUsername.Text = ""
    $TimeDisplay.Text = ""
    $UserLock.ForeColor = "#000000"
    $UserLock.Text = ""
    $ChangedPassword.Text = ""
    $CurrentPassword.Text = ""
    $CurrentPassword.ForeColor = "#000000"
}

#This function closes the GUI
function Close_Window {
    $PasswordGUI.Close()
}

#---------------------GUI Window---------------------

#This is the frame of the GUI
$PasswordGUI = New-Object System.Windows.Forms.Form
$PasswordGUI.ClientSize = '300,300'
$PasswordGUI.text = "Password Expiration"
$PasswordGUI.BackColor = "#ffffff"

#This designs the Title Label
$Title = New-Object system.Windows.Forms.Label
$Title.text = "Check Password Expiration"
$Title.AutoSize = $true
$Title.location = New-Object System.Drawing.Point(20,20)
$Title.Font = 'Microsoft Sans Serif, 13'

#This designs the Textbox for searching AD; Can also use Enter, CTRL, and Esc to Check, Clear, and Close respectively
$TextBox = New-Object system.Windows.Forms.TextBox
$TextBox.multiline = $false
$TextBox.width = 250
$TextBox.height = 15
$TextBox.location = New-Object System.Drawing.Point(25,50)
$TextBox.Font = 'Consolas, 10'
$TextBox.Add_KeyDown({if($_.KeyCode -eq "Enter"){Check_Password}})
$TextBox.Add_KeyDown({if($_.KeyCode -eq "ControlKey"){Clear_Window}})
$TextBox.Add_KeyDown({if($_.KeyCode -eq "Escape"){Close_Window}})

#Designs the username after the textbox is cleared
$DisplayUsername = New-Object System.Windows.Forms.Label
$DisplayUsername.AutoSize = $true
$DisplayUsername.location = New-Object System.Drawing.Point(20,90)
$DisplayUsername.Font = $displayFont

#Designs the current time label to appear when you search for a username
$TimeDisplay = New-Object System.Windows.Forms.Label
$TimeDisplay.AutoSize = $true
$TimeDisplay.location = New-Object System.Drawing.Point(20,120)
$TimeDisplay.Font = $displayFont

#Designs whether the user is locked or not in AD
$UserLock = New-Object System.Windows.Forms.Label
$UserLock.AutoSize = $true
$UserLock.location = New-Object System.Drawing.Point(20,150)
$UserLock.Font = $displayFont

#Designs last date when user changed their password
$ChangedPassword = New-Object System.Windows.Forms.Label
$ChangedPassword.AutoSize = $true
$ChangedPassword.location = New-Object System.Drawing.Point(20,180)
$ChangedPassword.Font = $displayFont

#Designs when user's password will expire
$CurrentPassword = New-Object System.Windows.Forms.Label
$CurrentPassword.AutoSize = $true
$CurrentPassword.location = New-Object System.Drawing.Point(20,210)
$CurrentPassword.Font = $displayFont

#Designs the Submit Button
$SubmitButton = New-Object system.Windows.Forms.Button
$SubmitButton.BackColor = "#012456"
$SubmitButton.text = "Submit"
$SubmitButton.width = $buttonWidth
$SubmitButton.height = $buttonHeight
$SubmitButton.location = New-Object System.Drawing.Point(35,250)
$SubmitButton.Font = $displayFont
$SubmitButton.ForeColor = "#ffffff"
$SubmitButton.Add_Click({ Check_Password })

#Designs the Clear Button
$ClearButton = New-Object system.Windows.Forms.Button
$ClearButton.BackColor = "#d0d0d0"
$ClearButton.text = "Clear"
$ClearButton.width = $buttonWidth
$ClearButton.height = $buttonHeight
$ClearButton.location = New-Object System.Drawing.Point(123,250)
$ClearButton.Font = $displayFont
$ClearButton.ForeColor = "#000000"
$ClearButton.Add_Click({ Clear_Window })

#Designs the Close Button
$CloseButton = New-Object System.Windows.Forms.Button
$CloseButton.BackColor = "#b0b0b0"
$CloseButton.text = "Close"
$CloseButton.width = $buttonWidth
$CloseButton.height = $buttonHeight
$CloseButton.location = New-Object System.Drawing.Point(210,250)
$CloseButton.Font = $displayFont
$CloseButton.ForeColor = "#000000"
$CloseButton.Add_Click({ Close_Window })

#This line adds each item to the GUI Window
$PasswordGUI.controls.AddRange(@($Title,$TextBox,$DisplayUsername,$TimeDisplay,$UserLock,$ChangedPassword,$CurrentPassword,$SubmitButton,$ClearButton,$CloseButton))

#This line displays the GUI Window
[void]$PasswordGUI.ShowDialog()
