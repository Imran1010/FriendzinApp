Add-Type -AssemblyName System.Windows.Forms
 
# GitHub Configuration
$usersFileUrl = "https://raw.githubusercontent.com/Imran1010/FriendzinApp/refs/heads/FriendzinQB/Users.json"
$repoUrl = "https://github.com/Imran1010/FriendzinApp.git"
$localPath = "$env:TEMP\UserAuthRepo"
$usersFilePath = "$localPath\Users.json"
 
# Function to Fetch User Data
function Get-Users {
    try {
        $jsonData = (Invoke-WebRequest -Uri $usersFileUrl -UseBasicParsing).Content
        return $jsonData | ConvertFrom-Json
    } catch {
        return @{}
    }
}
 
# Function to Hash Password
function Hash-Password {
    param ($password)
    return ConvertTo-SecureString $password -AsPlainText -Force | ConvertFrom-SecureString
}
 
# Function to Register a User
function Register-User {
    $username = $txtUser.Text
    $password = $txtPass.Text
 
    if ($username -eq "" -or $password -eq "") {
[System.Windows.Forms.MessageBox]::Show("Username and Password cannot be empty!", "Error", "OK", "Error")
        return
    }
 
    $users = Get-Users
    if ($users.$username) {
[System.Windows.Forms.MessageBox]::Show("User already exists!", "Error", "OK", "Error")
        return
    }
 
    # Hash password and store
    $hash = Hash-Password -password $password
    $users | Add-Member -MemberType NoteProperty -Name $username -Value $hash
 
    # Save to GitHub
    $users | ConvertTo-Json | Set-Content $usersFilePath
    git add users.json
    git commit -m "Added user $username"
    git push origin main
 
[System.Windows.Forms.MessageBox]::Show("User registered successfully!", "Success", "OK", "Information")
}
 
# Function to Login
function Login-User {
    $username = $txtUser.Text
    $password = $txtPass.Text
 
    if ($username -eq "" -or $password -eq "") {
[System.Windows.Forms.MessageBox]::Show("Username and Password cannot be empty!", "Error", "OK", "Error")
        return
    }
 
    $users = Get-Users
    $hash = Hash-Password -password $password
 
    if ($users.$username -eq $hash) {
[System.Windows.Forms.MessageBox]::Show("Login successful!", "Success", "OK", "Information")
        Show-AdminPanel
    } else {
[System.Windows.Forms.MessageBox]::Show("Invalid credentials!", "Error", "OK", "Error")
    }
}
 
# Function to Show Admin Panel
function Show-AdminPanel {
$adminForm = New-Object System.Windows.Forms.Form
    $adminForm.Text = "Admin Panel"
    $adminForm.Size = New-Object System.Drawing.Size(250, 150)
 
$btnReset = New-Object System.Windows.Forms.Button
    $btnReset.Location = New-Object System.Drawing.Point(50, 30)
    $btnReset.Size = New-Object System.Drawing.Size(150, 30)
    $btnReset.Text = "Reset AD Password"
    $btnReset.Add_Click({ Reset-ADPassword })
    $adminForm.Controls.Add($btnReset)
 
    $adminForm.ShowDialog()
}
 
# GUI Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "User Authentication"
$form.Size = New-Object System.Drawing.Size(300, 200)
 
# Username Label
$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Location = New-Object System.Drawing.Point(10, 20)
$lblUser.Size = New-Object System.Drawing.Size(80, 20)
$lblUser.Text = "Username:"
$form.Controls.Add($lblUser)
 
# Username Textbox
$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(100, 20)
$txtUser.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($txtUser)
 
# Password Label
$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Location = New-Object System.Drawing.Point(10, 50)
$lblPass.Size = New-Object System.Drawing.Size(80, 20)
$lblPass.Text = "Password:"
$form.Controls.Add($lblPass)
 
# Password Textbox
$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location = New-Object System.Drawing.Point(100, 50)
$txtPass.Size = New-Object System.Drawing.Size(150, 20)
$txtPass.PasswordChar = '*'
$form.Controls.Add($txtPass)
 
# Register Button
$btnRegister = New-Object System.Windows.Forms.Button
$btnRegister.Location = New-Object System.Drawing.Point(30, 90)
$btnRegister.Size = New-Object System.Drawing.Size(80, 30)
$btnRegister.Text = "Register"
$btnRegister.Add_Click({ Register-User })
$form.Controls.Add($btnRegister)
 
# Login Button
$btnLogin = New-Object System.Windows.Forms.Button
$btnLogin.Location = New-Object System.Drawing.Point(140, 90)
$btnLogin.Size = New-Object System.Drawing.Size(80, 30)
$btnLogin.Text = "Login"
$btnLogin.Add_Click({ Login-User })
$form.Controls.Add($btnLogin)
 
# Run Form
$form.ShowDialog()
