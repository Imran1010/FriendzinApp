Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$MKDIR = "C:\IHS-Application"
$ALL = "C:\IHS-Application\IHS-Template.csv"
if (-not (Test-Path $MKDIR)) {
New-Item -Path $MKDIR -ItemType Directory 
}
$IHSCSV = "UserName,AssignLicenses,RevokeLicenses,AddGroups,RemoveGroups,DLName,UserPrincipalName,Action.ToLower"
Set-Content -Path $ALL -Value $IHSCSV
Start-Transcript -Path C:\IHS-Application\IHS-LOGS-PS-MENU.log -Append
Get-Date -Format "dddd MM/dd/yyyy HH:mm K"
$maximumfunctioncount = '32768'

function Get-ImageFromUrl {
    param (
        [string]$Url
    )
    try {
        $webClient = New-Object System.Net.WebClient
        $imageStream = $webClient.OpenRead($Url)
        $image = [System.Drawing.Image]::FromStream($imageStream)
        $imageStream.Close()
        return $image
    } catch {
        Write-Error "Failed to load image from $Url. $_"
        return $null
    }
}

function Fetch-GistContent {
    param (
        [string]$url
    )
    try {
        $content = Invoke-RestMethod -Uri $url -Method Get
        return $content
    } catch {
        Write-Host "Error fetching content from IHS DataBase: $_" -ForegroundColor Red
        return $null
    }
}

function New-StyledButton {
    param (
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$Width = 180,
        [int]$Height = 40
    )
    $button = New-Object System.Windows.Forms.Button
    $button.Location = New-Object System.Drawing.Point($X, $Y)
    $button.Size = New-Object System.Drawing.Size($Width, $Height)
    $button.Text = $Text
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.FlatAppearance.BorderSize = 0
    $button.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $button.ForeColor = [System.Drawing.Color]::White
    $button.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $button.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(100, 160, 210) })
    $button.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180) })
    return $button
}

function New-StyledForm {
    param (
        [string]$Title,
        [int]$Width,
        [int]$Height
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size($Width, $Height)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::White
    return $form
}

function Show-LoginForm {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public class User32 {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, string lParam);
}
"@
    $EM_SETCUEBANNER = 0x1501

    $formLogin = New-Object System.Windows.Forms.Form
    $formLogin.Text = "Login Page"
    $formLogin.Size = New-Object System.Drawing.Size(525, 450)
    $formLogin.StartPosition = "CenterScreen"

    $backgroundPanel = New-Object System.Windows.Forms.Panel
    $backgroundPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $formLogin.Controls.Add($backgroundPanel)
    $backgroundImageUrl = "https://github.com/Imran1010/Applogin/blob/main/Untitled.jpg?raw=true"
    $backgroundImage = Get-ImageFromUrl -Url $backgroundImageUrl
    if ($backgroundImage) {
        $backgroundPanel.BackgroundImage = $backgroundImage
        $backgroundPanel.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Stretch
    }

    $controlsPanel = New-Object System.Windows.Forms.Panel
    $controlsPanel.Size = New-Object System.Drawing.Size(280, 280)
    $controlsPanel.Location = New-Object System.Drawing.Point(100, 112)  # Center the panel
    $controlsPanel.BackColor = [System.Drawing.Color]::FromArgb(0, [System.Drawing.Color]::White) 
    $backgroundPanel.Controls.Add($controlsPanel)

    $profilePicture = New-Object System.Windows.Forms.PictureBox
    $profilePictureUrl = "https://github.com/Imran1010/Applogin/blob/main/Logo.png?raw=true" 
    $profileImage = Get-ImageFromUrl -Url $profilePictureUrl
    if ($profileImage) {
        $profilePicture.Image = $profileImage
        $profilePicture.Size = New-Object System.Drawing.Size(160, 160)
        $profilePicture.Location = New-Object System.Drawing.Point(70, 0)
        $profilePicture.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
        $controlsPanel.Controls.Add($profilePicture)
    }

    $usernameBox = New-Object System.Windows.Forms.TextBox
    $usernameBox.Size = New-Object System.Drawing.Size(200, 20)
    $usernameBox.Location = New-Object System.Drawing.Point(50, 160)
    [User32]::SendMessage($usernameBox.Handle, $EM_SETCUEBANNER, 0, "Username")
    $controlsPanel.Controls.Add($usernameBox)

    $passwordBox = New-Object System.Windows.Forms.TextBox
    $passwordBox.Size = New-Object System.Drawing.Size(200, 20)
    $passwordBox.Location = New-Object System.Drawing.Point(50, 190)
    $passwordBox.UseSystemPasswordChar = $true
    [User32]::SendMessage($passwordBox.Handle, $EM_SETCUEBANNER, 0, "Password")
    $controlsPanel.Controls.Add($passwordBox)

    $loginButton = New-Object System.Windows.Forms.Button
    $loginButton.Text = "Login"
    $loginButton.Size = New-Object System.Drawing.Size(200, 30)
    $loginButton.Location = New-Object System.Drawing.Point(50, 220)
    $loginButton.BackColor = [System.Drawing.Color]::Gray
    $loginButton.ForeColor = [System.Drawing.Color]::White
    $loginButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $loginButton.FlatAppearance.BorderSize = 0
    $controlsPanel.Controls.Add($loginButton)

    $loginButton.Add_Click({
        $url = "https://raw.githubusercontent.com/Imran1010/Applogin/refs/heads/main/README.md"
        $gistContent = Fetch-GistContent -url $url
        if ($null -ne $gistContent) {
            $lines = $gistContent -split "`r`n" | ForEach-Object { $_.Trim() }
            $storedUsername = ""
            $storedPassword = ""
            foreach ($line in $lines) {
                if ($line -match "Username\s*:\s*(.*)") {
                    $storedUsername = $matches[1].Trim()
                }
                if ($line -match "Password\s*:\s*(.*)") {
                    $storedPassword = $matches[1].Trim()
                }
            }
            $username = $usernameBox.Text.Trim()
            $password = $passwordBox.Text.Trim()
            if ($username -eq $storedUsername -and $password -eq $storedPassword) {
                $formLogin.Hide()  
                Show-ModuleStatusForm
                $formLogin.Close()
            } else {
                [System.Windows.Forms.MessageBox]::Show("Invalid username or password.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Failed to fetch credentials from IHS Database Contact to Friendzin.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $formLogin.Add_Shown({ $formLogin.Activate() })
    [void]$formLogin.ShowDialog()
}

function Show-ModuleStatusForm {
    $IHSFMMD = New-StyledForm -Title "Module Check and Connect" -Width 600 -Height 200
    $IHSFMMD.BackColor = [System.Drawing.Color]::White
    $IHSFMMD.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

    $mainPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $mainPanel.Dock = "Fill"
    $mainPanel.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 10)
    $mainPanel.ColumnCount = 4
    $mainPanel.RowCount = 5
    $mainPanel.BackColor = [System.Drawing.Color]::White
    $mainPanel.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::Single

    $mainPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 22)))
    $mainPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
    $mainPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 30)))
    $mainPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 20)))

    0..4 | ForEach-Object {
        $mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 20)))
    }

    $headerStyle = @{
        Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
        ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    }

    $headers = @("Service", "Module Status", "Connection Status", "Action")
    0..3 | ForEach-Object {
        $header = New-Object System.Windows.Forms.Label
        $header.Text = $headers[$_]
        $header.Font = $headerStyle.Font
        $header.BackColor = $headerStyle.BackColor
        $header.ForeColor = $headerStyle.ForeColor
        $header.Dock = "Fill"
        $header.TextAlign = "MiddleCenter"
        $mainPanel.Controls.Add($header, $_, 0)
    }

    $services = @{
        "Azure AD" = @{
            ModuleName = "AzureAD"
            ConnectCmd = { Connect-AzureAD -ShowBanner:$false }
        }
        "Exchange Online" = @{
            ModuleName = "ExchangeOnlineManagement"
            ConnectCmd = { Connect-ExchangeOnline -ShowBanner:$false }
        }
        "Microsoft Graph" = @{
            ModuleName = "Microsoft.Graph"
            ConnectCmd = { Connect-MgGraph -ShowBanner:$false }
        }
    }

    $labelStyle = @{
        Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
        BackColor = [System.Drawing.Color]::White
        TextAlign = "MiddleCenter"
        Dock = "Fill"
        Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 3)
    }

    $row = 1
    foreach ($service in $services.Keys) {
        $serviceLabel = New-Object System.Windows.Forms.Label
        $serviceLabel.Text = $service
        $serviceLabel.Font = $labelStyle.Font
        $serviceLabel.BackColor = $labelStyle.BackColor
        $serviceLabel.TextAlign = $labelStyle.TextAlign
        $serviceLabel.Dock = $labelStyle.Dock
        $serviceLabel.Margin = $labelStyle.Margin
        $mainPanel.Controls.Add($serviceLabel, 0, $row)

        $moduleStatus = New-Object System.Windows.Forms.Label
        $moduleStatus.Font = $labelStyle.Font
        $moduleStatus.BackColor = $labelStyle.BackColor
        $moduleStatus.TextAlign = $labelStyle.TextAlign
        $moduleStatus.Dock = $labelStyle.Dock
        $moduleStatus.Margin = $labelStyle.Margin
        if (Get-Module -ListAvailable $services[$service].ModuleName) {
            $moduleStatus.Text = "✓ Installed"
            $moduleStatus.ForeColor = [System.Drawing.Color]::Green
        } else {
            $moduleStatus.Text = "✗ Not Installed"
            $moduleStatus.ForeColor = [System.Drawing.Color]::Red
        }
        $mainPanel.Controls.Add($moduleStatus, 1, $row)

        $connectionStatus = New-Object System.Windows.Forms.Label
        $connectionStatus.Text = "Not Connected"
        $connectionStatus.Font = $labelStyle.Font
        $connectionStatus.BackColor = $labelStyle.BackColor
        $connectionStatus.TextAlign = $labelStyle.TextAlign
        $connectionStatus.Dock = $labelStyle.Dock
        $connectionStatus.Margin = $labelStyle.Margin
        $connectionStatus.ForeColor = [System.Drawing.Color]::Red
        $mainPanel.Controls.Add($connectionStatus, 2, $row)

        $connectButton = New-Object System.Windows.Forms.Button
        $connectButton.Text = "Connect"
        $connectButton.Font = $labelStyle.Font
        $connectButton.Dock = "Fill"
        $connectButton.Margin = New-Object System.Windows.Forms.Padding(10, 5, 10, 5)
        $connectButton.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
        $connectButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $connectButton.Tag = @{
            Service = $service
            StatusLabel = $connectionStatus
        }

        $connectButton.Add_Click({
            $serviceInfo = $this.Tag
            $service = $serviceInfo.Service
            $status = $serviceInfo.StatusLabel

            $this.Enabled = $false
            $status.Text = "⟳ Connecting..."
            $status.ForeColor = [System.Drawing.Color]::Blue

            Start-Job -Name "Connect_$service" -ScriptBlock {
                param($serviceName, $moduleData)
                Import-Module $moduleData.ModuleName -Force
                & $moduleData.ConnectCmd
            } -ArgumentList $service, $services[$service]
        })

        $mainPanel.Controls.Add($connectButton, 3, $row)
        $row++
    }

    $nextButton = New-StyledButton -Text "Next" -X 400 -Y 150 -Width 180 -Height 40
    $nextButton.Add_Click({
        $IHSFMMD.Hide()
        Show-MainForm
        $IHSFMMD.Close()
    })
    $mainPanel.Controls.Add($nextButton, 2, 4)

    $loginButton = New-StyledButton -Text "Login" -X 200 -Y 150 -Width 180 -Height 40
    $loginButton.Add_Click({
        $IHSFMMD.Hide()
        Show-LoginForm
        $IHSFMMD.Close()
    })
    $mainPanel.Controls.Add($loginButton, 0, 4)

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 500
    $timer.Add_Tick({
        Get-Job | Where-Object { $_.Name -like "Connect_*" } | ForEach-Object {
            $serviceName = $_.Name -replace "Connect_"
            $row = [array]::IndexOf(($services.Keys), $serviceName) + 1
            $button = $mainPanel.GetControlFromPosition(3, $row)
            $status = $mainPanel.GetControlFromPosition(2, $row)

            if ($_.State -eq "Completed") {
                $status.Text = "✓ Connected"
                $status.ForeColor = [System.Drawing.Color]::Green
                $button.Text = "Connected"
                $button.BackColor = [System.Drawing.Color]::FromArgb(225, 240, 225)
                Remove-Job $_
            }
            elseif ($_.State -eq "Failed") {
                $error = Receive-Job $_
                $status.Text = "✗ Failed"
                $status.ForeColor = [System.Drawing.Color]::Red
                $button.Text = "Connect"
                $button.Enabled = $true
                Remove-Job $_
                [System.Windows.Forms.MessageBox]::Show($error, "Connection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })
    $timer.Start()
    $IHSFMMD.Controls.Add($mainPanel)
    [void]$IHSFMMD.ShowDialog()
}
function IHS-USPWRST-MLUP {
Add-Type -AssemblyName System.Windows.Forms
$UGDNInputForm = New-Object System.Windows.Forms.Form
$UGDNInputForm.Text = "Password Reset"
$UGDNInputForm.Size = New-Object System.Drawing.Size(300, 150)
$UGDNInputForm.StartPosition = "CenterScreen"

$UGDNInputLabel = New-Object System.Windows.Forms.Label
$UGDNInputLabel.Text = " ⬇ Enter UGDN ⬇ "
$UGDNInputLabel.AutoSize = $true
$UGDNInputLabel.Location = New-Object System.Drawing.Point(10, 20)
$UGDNInputForm.Controls.Add($UGDNInputLabel)

$UGDNInputTextBox = New-Object System.Windows.Forms.TextBox
$UGDNInputTextBox.Location = New-Object System.Drawing.Point(10, 40)
$UGDNInputTextBox.Size = New-Object System.Drawing.Size(260, 20)
$UGDNInputForm.Controls.Add($UGDNInputTextBox)

$UGDNInputOKButton = New-Object System.Windows.Forms.Button
$UGDNInputOKButton.Location = New-Object System.Drawing.Point(50, 80)
$UGDNInputOKButton.Size = New-Object System.Drawing.Size(75, 23)
$UGDNInputOKButton.Text = "OK"
$UGDNInputOKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$UGDNInputForm.Controls.Add($UGDNInputOKButton)

$UGDNInputCancelButton = New-Object System.Windows.Forms.Button
$UGDNInputCancelButton.Location = New-Object System.Drawing.Point(150, 80)
$UGDNInputCancelButton.Size = New-Object System.Drawing.Size(75, 23)
$UGDNInputCancelButton.Text = "Cancel"
$UGDNInputCancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$UGDNInputForm.Controls.Add($UGDNInputCancelButton)
$UGDNInputResult = $UGDNInputForm.ShowDialog()
if ($UGDNInputResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $IHSUS = $UGDNInputTextBox.Text

    $IHSUSIF = Get-ADUser -Filter { SamAccountName -like $IHSUS } -Properties SamAccountName, DistinguishedName, LockedOut, Enabled | Select-Object SamAccountName, DistinguishedName, LockedOut, Enabled

    if ($IHSUSIF) {
        $lockstatus = $IHSUSIF.LockedOut
        $enabledStatus = $IHSUSIF.Enabled

        if ($enabledStatus -eq $false) {
            Write-Host -ForegroundColor Red "UGDN Account is disabled"
        }
        elseif ($lockstatus -eq $true) {
            Write-Host -ForegroundColor Red "Account locked"
            Write-Host ""
            Write-Host "Unlocking Account"
            
            Unlock-ADAccount $IHSUSIF.SamAccountName
        }
        else {
            Write-Host -ForegroundColor Green "We checked Account is Unlocked"
           
            $IHSNWPW = Read-Host "Enter New Password" -AsSecureString
            
            Set-ADAccountPassword -Identity $IHSUSIF.SamAccountName -NewPassword $IHSNWPW -Reset
            Write-Host -ForegroundColor Green "Password reset successfully"
            [System.Windows.Forms.MessageBox]::Show("UGDN Password reset", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

} }
        
     else {
     
        Write-Host "User not found" -ForegroundColor Red
     [System.Windows.Forms.MessageBox]::Show("User not found `n Check AD Connectivity or VPN", "Failled", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    }
} else {
    Write-Host "Operation canceled" -ForegroundColor Red
   [System.Windows.Forms.MessageBox]::Show("Operation canceled", "Retry", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

}
}

function IHS-USEMUP {
Add-Type -AssemblyName System.Windows.Forms

$UGDNInputForm = New-Object System.Windows.Forms.Form
$UGDNInputForm.Text = "Enter UGDN"
$UGDNInputForm.Size = New-Object System.Drawing.Size(300, 150)
$UGDNInputForm.StartPosition = "CenterScreen"

$UGDNInputLabel = New-Object System.Windows.Forms.Label
$UGDNInputLabel.Text = "Enter the UGDN for Email Update:"
$UGDNInputLabel.AutoSize = $true
$UGDNInputLabel.Location = New-Object System.Drawing.Point(10, 20)
$UGDNInputForm.Controls.Add($UGDNInputLabel)

$UGDNInputTextBox = New-Object System.Windows.Forms.TextBox
$UGDNInputTextBox.Location = New-Object System.Drawing.Point(10, 40)
$UGDNInputTextBox.Size = New-Object System.Drawing.Size(260, 20)
$UGDNInputForm.Controls.Add($UGDNInputTextBox)

$UGDNInputOKButton = New-Object System.Windows.Forms.Button
$UGDNInputOKButton.Location = New-Object System.Drawing.Point(50, 80)
$UGDNInputOKButton.Size = New-Object System.Drawing.Size(75, 23)
$UGDNInputOKButton.Text = "OK"
$UGDNInputOKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$UGDNInputForm.Controls.Add($UGDNInputOKButton)

$UGDNInputCancelButton = New-Object System.Windows.Forms.Button
$UGDNInputCancelButton.Location = New-Object System.Drawing.Point(150, 80)
$UGDNInputCancelButton.Size = New-Object System.Drawing.Size(75, 23)
$UGDNInputCancelButton.Text = "Cancel"
$UGDNInputCancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$UGDNInputForm.Controls.Add($UGDNInputCancelButton)

$UGDNInputResult = $UGDNInputForm.ShowDialog()


$MAILInputForm = New-Object System.Windows.Forms.Form
$MAILInputForm.Text = "Enter Mail"
$MAILInputForm.Size = New-Object System.Drawing.Size(300, 150)
$MAILInputForm.StartPosition = "CenterScreen"

$MAILInputLabel = New-Object System.Windows.Forms.Label
$MAILInputLabel.Text = "Enter the MAIL:"
$MAILInputLabel.AutoSize = $true
$MAILInputLabel.Location = New-Object System.Drawing.Point(10, 20)
$MAILInputForm.Controls.Add($MAILInputLabel)

$MAILInputTextBox = New-Object System.Windows.Forms.TextBox
$MAILInputTextBox.Location = New-Object System.Drawing.Point(10, 40)
$MAILInputTextBox.Size = New-Object System.Drawing.Size(260, 20)
$MAILInputForm.Controls.Add($MAILInputTextBox)

$MAILInputOKButton = New-Object System.Windows.Forms.Button
$MAILInputOKButton.Location = New-Object System.Drawing.Point(50, 80)
$MAILInputOKButton.Size = New-Object System.Drawing.Size(75, 23)
$MAILInputOKButton.Text = "OK"
$MAILInputOKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$MAILInputForm.Controls.Add($MAILInputOKButton)

$MAILInputCancelButton = New-Object System.Windows.Forms.Button
$MAILInputCancelButton.Location = New-Object System.Drawing.Point(150, 80)
$MAILInputCancelButton.Size = New-Object System.Drawing.Size(75, 23)
$MAILInputCancelButton.Text = "Cancel"
$MAILInputCancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$MAILInputForm.Controls.Add($MAILInputCancelButton)

if ($UGDNInputResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $IHSUS = $UGDNInputTextBox.Text

    $IHSUSIF = Get-ADUser -Filter { SamAccountName -like $IHSUS } -Properties SamAccountName, DistinguishedName, LockedOut, Enabled | Select-Object SamAccountName, DistinguishedName, LockedOut, Enabled, Mail

    if ($IHSUSIF) {
        $lockstatus = $IHSUSIF.LockedOut
        $enabledStatus = $IHSUSIF.Enabled

        if ($enabledStatus -eq $false) {
            Write-Host -ForegroundColor Red "UGDN Account is disabled"
        }
        elseif ($lockstatus -eq $true) {
            Write-Host -ForegroundColor Red "Account locked"
            Write-Host ""
            Write-Host "Unlocking Account"
            
            Unlock-ADAccount $IHSUSIF.SamAccountName
        }
        else {

            $EmailInputResult = [System.Windows.Forms.MessageBox]::Show("Do you want to Add an Email Address & Proxy Address?", "Add Email Address and Proxy Address", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

            if ($EmailInputResult -eq [System.Windows.Forms.DialogResult]::Yes)
{

$MAILInputResult = $MAILInputForm.ShowDialog()

            $userEmailAddress = $MAILInputTextBox.Text

                    $user = Get-ADUser -Identity $IHSUS -Properties mail

                    if ($user) {
                        Set-ADUser -Identity $IHSUS -EmailAddress $userEmailAddress

                        $proxyAddress = "SMTP:" + $userEmailAddress

                        Set-ADUser -Identity $IHSUS -Add @{proxyAddresses=$proxyAddress}

                        Write-Host "`nEmail address and SMTP proxy address added successfully for user $IHSUS" -ForegroundColor Green
      [System.Windows.Forms.MessageBox]::Show("Email address and SMTP proxy address added for $IHSUS", "Successfully", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

                    } else {
                        Write-Host "`nUser $IHSUS not found in Active Directory" -ForegroundColor Red  
      [System.Windows.Forms.MessageBox]::Show("User $IHSUS not found in Active Directory", "Failled", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)


 } } } }
        
     else {
     
        Write-Host "User not found." -ForegroundColor Red
      [System.Windows.Forms.MessageBox]::Show("User not found `n Check AD Connectivity Or VPN", "Failled", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    }
} else {
    Write-Host "Operation canceled." -ForegroundColor Red
          [System.Windows.Forms.MessageBox]::Show("Operation canceled", "Retry", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

}


}

function IHS-USCMTRST {

$UGDNInputForm = New-Object System.Windows.Forms.Form
$UGDNInputForm.Text = "Enter UGDN"
$UGDNInputForm.Size = New-Object System.Drawing.Size(300, 150)
$UGDNInputForm.StartPosition = "CenterScreen"

$UGDNInputLabel = New-Object System.Windows.Forms.Label
$UGDNInputLabel.Text = "Enter the UGDN:"
$UGDNInputLabel.AutoSize = $true
$UGDNInputLabel.Location = New-Object System.Drawing.Point(10, 20)
$UGDNInputForm.Controls.Add($UGDNInputLabel)

$UGDNInputTextBox = New-Object System.Windows.Forms.TextBox
$UGDNInputTextBox.Location = New-Object System.Drawing.Point(10, 40)
$UGDNInputTextBox.Size = New-Object System.Drawing.Size(260, 20)
$UGDNInputForm.Controls.Add($UGDNInputTextBox)

$UGDNInputOKButton = New-Object System.Windows.Forms.Button
$UGDNInputOKButton.Location = New-Object System.Drawing.Point(50, 80)
$UGDNInputOKButton.Size = New-Object System.Drawing.Size(75, 23)
$UGDNInputOKButton.Text = "OK"
$UGDNInputOKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$UGDNInputForm.Controls.Add($UGDNInputOKButton)

$UGDNInputCancelButton = New-Object System.Windows.Forms.Button
$UGDNInputCancelButton.Location = New-Object System.Drawing.Point(150, 80)
$UGDNInputCancelButton.Size = New-Object System.Drawing.Size(75, 23)
$UGDNInputCancelButton.Text = "Cancel"
$UGDNInputCancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$UGDNInputForm.Controls.Add($UGDNInputCancelButton)

$UGDNInputResult = $UGDNInputForm.ShowDialog()



if ($UGDNInputResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $IHSUS = $UGDNInputTextBox.Text

    $IHSUSIF = Get-ADUser -Filter { SamAccountName -like $IHSUS } -Properties SamAccountName, DistinguishedName, LockedOut, Enabled | Select-Object SamAccountName, DistinguishedName, LockedOut, Enabled

    if ($IHSUSIF) {
        $lockstatus = $IHSUSIF.LockedOut
        $enabledStatus = $IHSUSIF.Enabled

        if ($enabledStatus -eq $false) {
            Write-Host -ForegroundColor Red "UGDN Account is disabled"
        }
        elseif ($lockstatus -eq $true) {
            Write-Host -ForegroundColor Red "Account locked"
            Write-Host ""
            Write-Host "Unlocking Account"
            
            Unlock-ADAccount $IHSUSIF.SamAccountName
        }
        else {
            Write-Host -ForegroundColor Green "We checked Account is Unlocked"

            Set-ADUser -Identity $IHSUSIF.SamAccountName -Clear Comment
            $user = Get-ADUser -Identity $IHSUSIF.SamAccountName -Properties Comment
            
            Write-Host -ForegroundColor Green "Comment reset successfully"
[System.Windows.Forms.MessageBox]::Show("Comment reset", "Successfully", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)



} }
        
     else {
     
         Write-Host "User not found." -ForegroundColor Red
          [System.Windows.Forms.MessageBox]::Show("User not found `n Check AD Connectivity Or VPN", "Failled", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    }
} else {
    Write-Host "Operation canceled." -ForegroundColor Red
 [System.Windows.Forms.MessageBox]::Show("Operation canceled", "Retry", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

}


}

function IHS-GRPLFT {

Add-Type -AssemblyName System.Windows.Forms

$form = New-Object System.Windows.Forms.Form
$form.Text = "User Group Assignment"
$form.Size = New-Object System.Drawing.Size(300, 150)
$form.StartPosition = "CenterScreen"

$userALabel = New-Object System.Windows.Forms.Label
$userALabel.Text = "User A:"
$userALabel.AutoSize = $true
$userALabel.Location = New-Object System.Drawing.Point(10, 20)
$form.Controls.Add($userALabel)

$userATextBox = New-Object System.Windows.Forms.TextBox
$userATextBox.Location = New-Object System.Drawing.Point(120, 20)
$userATextBox.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($userATextBox)

$userBLabel = New-Object System.Windows.Forms.Label
$userBLabel.Text = "User B:"
$userBLabel.AutoSize = $true
$userBLabel.Location = New-Object System.Drawing.Point(10, 50)
$form.Controls.Add($userBLabel)

$userBTextBox = New-Object System.Windows.Forms.TextBox
$userBTextBox.Location = New-Object System.Drawing.Point(120, 50)
$userBTextBox.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($userBTextBox)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(50, 90)
$okButton.Size = New-Object System.Drawing.Size(75, 23)
$okButton.Text = "OK"
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150, 90)
$cancelButton.Size = New-Object System.Drawing.Size(75, 23)
$cancelButton.Text = "Cancel"
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.Controls.Add($cancelButton)

$form.AcceptButton = $okButton
$form.CancelButton = $cancelButton
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $userA = $userATextBox.Text
    $userB = $userBTextBox.Text

    $userAGroups = Get-ADUser -Identity $userA -Properties MemberOf | Select-Object -ExpandProperty MemberOf

    foreach ($group in $userAGroups) {
        Add-ADGroupMember -Identity $group -Members $userB
    }

 [System.Windows.Forms.MessageBox]::Show("Thanks for Using IHS Script `nAdded $userB to groups where $userA is a member.", "Group Lifting Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)


    Write-Host "Added $userB to groups where $userA is a member." 
} else {
    Write-Host "Operation canceled." -ForegroundColor Red
[System.Windows.Forms.MessageBox]::Show("Operation canceled", "Retry", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

}


}

function IHS-ADEXP-PT {

$currentDate = (Get-Date).ToString("dd-MM-yyyy")

$IHSEXPATH = "C:\IHS-Application\ADUserexport-$currentDate.csv"

$IHSPRPTS = 'SamAccountName','UserPrincipalName','Mail','Name','CanonicalName','DisplayName','Enabled','distinguishedName','manager','employeeid','extensionAttribute1','extensionAttribute2','extensionAttribute3','extensionAttribute4','extensionAttribute5','extensionAttribute6','extensionAttribute7','extensionAttribute8','extensionAttribute9','extensionAttribute10','extensionAttribute11','extensionAttribute12','extensionAttribute13','extensionAttribute14','extensionAttribute15','CannotChangePassword','LastLogonDate','street','PasswordLastSet','PasswordExpired','PasswordNeverExpires','AccountExpirationDate','Office','City','whenCreated','Description'

Get-ADUser -Filter * -Properties $IHSPRPTS | Select-Object $IHSPRPTS | Export-csv -path $IHSEXPATH -NoTypeInformation -Encoding UTF8

Write-Host "AD Users has been successfully Exported on C Drive" 

[System.Windows.Forms.MessageBox]::Show("Thanks for Using IHS Script `n AD Users has been successfully Exported on C Drive .", "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

function IHS-ADOUEXP-PT {

$IHSFMOU = New-Object System.Windows.Forms.Form
$IHSFMOU.Text = "OU Path Input"
$IHSFMOU.Size = New-Object System.Drawing.Size(300, 150)
$IHSFMOU.StartPosition = "CenterScreen"

$IHSLBOU = New-Object System.Windows.Forms.Label
$IHSLBOU.Text = "Enter the OU path:"
$IHSLBOU.AutoSize = $true
$IHSLBOU.Location = New-Object System.Drawing.Point(10, 20)
$IHSFMOU.Controls.Add($IHSLBOU)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10, 40)
$textBox.Size = New-Object System.Drawing.Size(260, 20)
$IHSFMOU.Controls.Add($textBox)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(50, 80)
$okButton.Size = New-Object System.Drawing.Size(75, 23)
$okButton.Text = "OK"
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$IHSFMOU.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150, 80)
$cancelButton.Size = New-Object System.Drawing.Size(75, 23)
$cancelButton.Text = "Cancel"
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$IHSFMOU.Controls.Add($cancelButton)

$result = $IHSFMOU.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $IHS = $textBox.Text
    $currentDate = (Get-Date).ToString("dd-MM-yyyy")
    $IHSEXPATH = "C:\IHS-Application\${IHS}-$currentDate.csv"


   ## $IHSPRPTS = 'Name','SamAccountName','DisplayName','UserPrincipalName','Mail','LastLogonDate','street','PasswordLastSet','CannotChangePassword','PasswordExpired','PasswordNeverExpires','AccountExpirationDate','CanonicalName','Enabled','employeeid','Office','City','whenCreated','distinguishedName','Description'
    # UPL >> 
    $IHSPRPTS = 'Name','SamAccountName','DisplayName','UserPrincipalName','Mail','LastLogonDate','street','PasswordLastSet','CannotChangePassword','PasswordExpired','PasswordNeverExpires','AccountExpirationDate','CanonicalName','Enabled','employeeid','Office','City','whenCreated','distinguishedName','extensionAttribute1','extensionAttribute2','extensionAttribute3','extensionAttribute4','extensionAttribute5','extensionAttribute6','extensionAttribute7','extensionAttribute8','extensionAttribute9','extensionAttribute10','extensionAttribute11','extensionAttribute12','extensionAttribute13','extensionAttribute14','extensionAttribute15','Description'

    Get-ADUser -Filter * -Properties $IHSPRPTS -SearchBase $IHS | Select-Object $IHSPRPTS | Export-csv -path $IHSEXPATH -NoTypeInformation -Encoding UTF8



    [System.Windows.Forms.MessageBox]::Show("Thanks for Using IHS Script `nOU Users has been successfully Exported .", "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

} else {
    Write-Host "Operation canceled." -ForegroundColor Red
 [System.Windows.Forms.MessageBox]::Show("Operation canceled", "Retry", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

}

}

function IHSADLISTTY{
$IHSINPT = 'C:\IHS-Application\IHS-Template.csv'

$IHSOPPT = 'C:\IHS-Application\ADExportUserBased.csv'

$propertiesToRetrieve = 'SamAccountName','UserPrincipalName','Mail','Name','CanonicalName','DisplayName','Enabled','distinguishedName','manager','employeeid','extensionAttribute1','extensionAttribute2','extensionAttribute3','extensionAttribute4','extensionAttribute5','extensionAttribute6','extensionAttribute7','extensionAttribute8','extensionAttribute9','extensionAttribute10','extensionAttribute11','extensionAttribute12','extensionAttribute13','extensionAttribute14','extensionAttribute15','CannotChangePassword','LastLogonDate','street','PasswordLastSet','PasswordExpired','PasswordNeverExpires','AccountExpirationDate','Office','City','whenCreated','Description'

$usernames = Import-Csv -Path $IHSINPT | Select-Object -ExpandProperty UserName

$userInfoArray = @()

foreach ($username in $usernames) {
    $user = Get-ADUser -Filter { SamAccountName -eq $username } -Properties $propertiesToRetrieve
    if ($user) {
        $userInfoArray += $user | Select-Object $propertiesToRetrieve
    } else {
        Write-Host "User $username not found in Active Directory."
    }
}
$userInfoArray | Export-Csv -Path $IHSOPPT -NoTypeInformation
[System.Windows.Forms.MessageBox]::Show("Thanks for Using IHS Script `n User information has been exported to $IHSOPPT .", "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
Write-Host "User information has been exported to $IHSOPPT"
}

function IHD-USSHIMUT {

Add-Type -AssemblyName PresentationFramework

$window = New-Object Windows.Window
$window.Title = "Enter User's UPN"
$window.SizeToContent = "WidthAndHeight"
$window.WindowStartupLocation = "CenterScreen"

$label = New-Object Windows.Controls.Label
$label.Content = "Enter the User's UPN:"
$textBox = New-Object Windows.Controls.TextBox
$textBox.Margin = "5"
$textBox.Width = "200"
$textBox.Add_KeyDown({
    if ($_.Key -eq "Enter") { 
        $window.Close() 
    }
})

$stackPanel = New-Object Windows.Controls.StackPanel
$stackPanel.Orientation = "Vertical"
$stackPanel.Children.Add($label)
$stackPanel.Children.Add($textBox)

$window.Content = $stackPanel

$window.ShowDialog() | Out-Null

$userUPN = $textBox.Text.Trim()

if (-not [string]::IsNullOrWhiteSpace($userUPN)) {
    try {
        $user = Get-ADUser -Filter "sAMAccountName -eq '$userUPN'" -Properties SamAccountName, UserPrincipalName, ObjectGUID     
        if ($null -eq $user) {
            throw "User not found in Active Directory."
        }
        $userUPN = $user.UserPrincipalName
        $immutableId = [System.Convert]::ToBase64String($user.ObjectGUID.ToByteArray())
        $azureUser = Get-MgUser -Filter "UserPrincipalName eq '$userUPN'"
        
        if ($null -eq $azureUser) {
            throw "User not found in Microsoft 365."
        }

        $messageBoxText = "User's Immutable ID: $immutableId"
        $messageBoxCaption = "Immutable ID"
        $messageBoxButtons = [System.Windows.MessageBoxButton]::OKCancel
        $messageBoxIcon = [System.Windows.MessageBoxImage]::Information
        
        $result = [System.Windows.MessageBox]::Show($messageBoxText, $messageBoxCaption, $messageBoxButtons, $messageBoxIcon)

        if ($result -eq [System.Windows.MessageBoxResult]::OK) {
            $immutableId | Set-Clipboard
            Write-Host "Immutable ID copied to clipboard."
        }

    } catch {
        [System.Windows.MessageBox]::Show("Error: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
} else {
    Write-Host "UPN not provided."
}

}

function IHS-MLTGRP {

$csvPath = "C:\IHS-Application\IHS-Template.csv"

$logPath = "C:\IHS-Application\IHS-GRPTASK.csv"

if (-not (Test-Path $csvPath)) {
    Write-Host "Error: The CSV file at '$csvPath' does not exist. Please check the file path." -ForegroundColor Red
    return
}

try {
    $entries = Import-Csv -Path $csvPath
} catch {
    Write-Host "Error: Unable to read the CSV file. $_" -ForegroundColor Red
    return
}

if (-not $entries) {
    Write-Host "Error: The CSV file is empty or improperly formatted." -ForegroundColor Red
    return
}

$logData = @()

foreach ($entry in $entries) {
    $samAccountName = $entry.Username
    $addGroups = $entry.AddGroups -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    $removeGroups = $entry.RemoveGroups -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

    if (-not $samAccountName) {
        Write-Host "Skipping entry with missing Username." -ForegroundColor Yellow
        continue
    }

    try {
        $user = Get-ADUser -Filter { SamAccountName -eq $samAccountName } -ErrorAction Stop
        if (-not $user) {
            throw "User $samAccountName not found"
        }

        foreach ($groupName in $addGroups) {
            if ($groupName) {
                try {
                    $group = Get-ADGroup -Filter { Name -eq $groupName } -ErrorAction Stop
                    if (-not $group) {
                        throw "Group $groupName not found"
                    }

                    Add-ADGroupMember -Identity $group -Members $user -ErrorAction Stop
                    Write-Host "Successfully added $samAccountName to $groupName."

                    $logData += [pscustomobject]@{
                        SamAccountName = $samAccountName
                        GroupName      = $groupName
                        Action         = "Add"
                        Status         = "Success"
                        Message        = "User added to group"
                    }
                } catch {
                    Write-Host "Failed to add $samAccountName to $groupName : $_" -ForegroundColor Yellow
                    $logData += [pscustomobject]@{
                        SamAccountName = $samAccountName
                        GroupName      = $groupName
                        Action         = "Add"
                        Status         = "Failed"
                        Message        = $_.Exception.Message
                    }
                }
            }
        }

        foreach ($groupName in $removeGroups) {
            if ($groupName) {
                try {
                    $group = Get-ADGroup -Filter { Name -eq $groupName } -ErrorAction Stop
                    if (-not $group) {
                        throw "Group $groupName not found"
                    }

                    Remove-ADGroupMember -Identity $group -Members $user -Confirm:$false -ErrorAction Stop
                    Write-Host "Successfully removed $samAccountName from $groupName."

                    $logData += [pscustomobject]@{
                        SamAccountName = $samAccountName
                        GroupName      = $groupName
                        Action         = "Remove"
                        Status         = "Success"
                        Message        = "User removed from group"
                    }
                } catch {
                    Write-Host "Failed to remove $samAccountName from $groupName : $_" -ForegroundColor Yellow
                    $logData += [pscustomobject]@{
                        SamAccountName = $samAccountName
                        GroupName      = $groupName
                        Action         = "Remove"
                        Status         = "Failed"
                        Message        = $_.Exception.Message
                    }
                }
            }
        }
    } catch {
        Write-Host "Error: User $samAccountName not found or other issue occurred." -ForegroundColor Red
        $logData += [pscustomobject]@{
            SamAccountName = $samAccountName
            GroupName      = "N/A"
            Action         = "N/A"
            Status         = "Failed"
            Message        = $_.Exception.Message
        }
    }
}

if ($logData) {
    try {
        $logData | Export-Csv -Path $logPath -NoTypeInformation -Force

   [System.Windows.Forms.MessageBox]::Show("Operation completed successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

        Write-Host "Operation complete. Log file saved to $logPath" -ForegroundColor Green
    } catch {
        Write-Host "Error: Failed to save the log file. $_" -ForegroundColor Red
    }
} else {
    Write-Host "No actions were performed. Log file not created." -ForegroundColor Yellow
}

}

function IHS-MDINT {


function IHSRSAT {

    $rsatComponents = @(
        "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0",
        "Rsat.DHCP.Tools~~~~0.0.1.0",
        "Rsat.Dns.Tools~~~~0.0.1.0",
        "Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0",
        "Rsat.ServerManager.Tools~~~~0.0.1.0"
    )

    $osVersion = [System.Environment]::OSVersion.Version
    if ($osVersion.Major -lt 10 -or ($osVersion.Major -eq 10 -and $osVersion.Build -lt 17763)) {
        Write-Host "RSAT is only supported on Windows 10 version 1809 (Build 17763) and later, Any other query connect with India.core (Ibrahim)"
        return
    }

    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "Please run this script as an administrator."
        return
    }

    try {
        foreach ($component in $rsatComponents) {
            Write-Host "Attempting to install $component..."
            Add-WindowsCapability -Name $component -Online -ErrorAction Stop
            Write-Host "$component installed successfully."
        }

        [System.Windows.Forms.MessageBox]::Show("RSAT modules installed successfully: LDAP, GPMT, DHCP, DNS, Server Manager.", "Installation Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error installing RSAT modules: $_", "Installation Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        Write-Host "Detailed error: $_"
    }



}

function IHSEXOL {

function Check-ExchangeOnlineModule {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "Exchange Online Management module is not installed. Installing it now..."
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
    } else {
        Write-Host "Exchange Online Management module is already installed."
    }
}

function Check-PowerShellVersion {
    $psVersion = $PSVersionTable.PSVersion
    if ($psVersion.Major -lt 5) {
        Write-Host "PowerShell version is less than 5. Please update PowerShell."
        exit 1
    }
}

function Check-NetworkConnectivity {
    $url = "https://outlook.office365.com"
    try {
        $request = Invoke-WebRequest -Uri $url -UseBasicP -TimeoutSec 5
        if ($request.StatusCode -ne 200) {
            Write-Host "Network issue: Cannot reach Exchange Online."
            exit 1
        }
    } catch {
        Write-Host "Network connectivity issue. Please check your internet connection."
        exit 1
    }
}

function Connect-ToExchangeOnline {
    
        Connect-ExchangeOnline -UserPrincipalName "30035113@upl-ltd.com"
[System.Windows.Forms.MessageBox]::Show("Exchange Online Connected", "Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    
}

Check-ExchangeOnlineModule
Check-PowerShellVersion
Check-NetworkConnectivity
Connect-ToExchangeOnline
}

function IHSAZAD {
function Check-AzureADModule {
    if (-not (Get-Module -ListAvailable -Name AzureAD)) {
        Write-Host "AzureAD module is not installed. Installing it now..."
        Install-Module -Name AzureAD -Force -AllowClobber -Scope CurrentUser
    } else {
        Write-Host "AzureAD module is already installed."
    }
}

function Check-PowerShellVersion {
    $psVersion = $PSVersionTable.PSVersion
    if ($psVersion.Major -lt 5) {
        Write-Host "PowerShell version is less than 5. Please update PowerShell."
        exit 1
    }
}

function Check-NetworkConnectivity {
    $url = "https://login.microsoftonline.com"
    try {
        $request = Invoke-WebRequest -Uri $url -UseBasicP -TimeoutSec 5
        if ($request.StatusCode -ne 200) {
            Write-Host "Network issue: Cannot reach Azure login."
            exit 1
        }
    } catch {
        Write-Host "Network connectivity issue. Please check your internet connection."
        exit 1
    }
}

function Connect-ToAzureAD {
        Connect-AzureAD
        [System.Windows.Forms.MessageBox]::Show("Azure AD Connected", "Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
   
}

Check-AzureADModule
Check-PowerShellVersion
Check-NetworkConnectivity
Connect-ToAzureAD

}


function IHSMSGP {
function Check-MSGPModule {
    if (-not (Get-InstalledModule Microsoft.Graph)) {
        Write-Host "MIcrosoft Graph module is not installed. Installing it now..."
     Install-Module Microsoft.Graph -AllowClobber -Force 

     [System.Windows.Forms.MessageBox]::Show("Microsoft Graph Installed", "Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    } else {
        Write-Host "Microsoft Graph module is already installed."
      [System.Windows.Forms.MessageBox]::Show("Microsoft Graph Already Installed", "Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    }
}

function Check-PowerShellVersion {
    $psVersion = $PSVersionTable.PSVersion
    if ($psVersion.Major -lt 5) {
        Write-Host "PowerShell version is less than 5. Please update PowerShell."
        exit 1
    }
}

function Check-NetworkConnectivity {
    $url = "https://login.microsoftonline.com"
    try {
        $request = Invoke-WebRequest -Uri $url -UseBasicP -TimeoutSec 5
        if ($request.StatusCode -ne 200) {
            Write-Host "Network issue: Cannot reach MG Graph login."
            exit 1
        }
    } catch {
        Write-Host "Network connectivity issue. Please check your internet connection."
        exit 1
    }
}

function Connect-ToMSGRP {
        Connect-MgGraph
        [System.Windows.Forms.MessageBox]::Show("MIcrosoft Graph Connected", "Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
   
}

Check-MSGPModule
Check-PowerShellVersion
Check-NetworkConnectivity
Connect-ToMSGRP

}




$form = New-Object System.Windows.Forms.Form
$form.Text = "UPL Powershell Module Menu"
$form.AutoSize = $True
$form.Size = New-Object System.Drawing.Size(350, 30) # Set the desired width and height
$labelIHS.Location = New-Object System.Drawing.Point(6, 10)
$form.BackColor = [System.Drawing.Color]::LightBlue

$labelIHS = New-Object System.Windows.Forms.Label
$labelIHS.Text = "Module Menu For Execute:"
$labelIHS.AutoSize = $False
$labelIHS.Size = New-Object System.Drawing.Size(350, 30) # Set the desired width and height
$labelIHS.Location = New-Object System.Drawing.Point(6, 10)
$labelIHS.Font = New-Object System.Drawing.Font("Cambria",14,[System.Drawing.FontStyle]::Bold)

$labelIHS.ForeColor = "Black"
$form.Controls.Add($labelIHS)



$IHSRSAT = New-Object System.Windows.Forms.Button
$IHSRSAT.Text = "Rsat Module"
$IHSRSAT.AutoSize = $True
$IHSRSAT.Location = New-Object System.Drawing.Point(10, 40)
$IHSRSAT.ForeColor = "White"
$IHSRSAT.BackColor = "Green"
$IHSRSAT.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$IHSRSAT.Add_Click({
    IHSRSAT
})
$IHSRSAT.Add_MouseEnter({
    $IHSRSAT.BackColor = "DarkRed"
})
$IHSRSAT.Add_MouseLeave({
    $IHSRSAT.BackColor = "Green"
})
$form.Controls.Add($IHSRSAT)

$IHSEXOL = New-Object System.Windows.Forms.Button
$IHSEXOL.Text = "Exchange Online"
$IHSEXOL.AutoSize = $True
$IHSEXOL.Location = New-Object System.Drawing.Point(10, 70)
$IHSEXOL.ForeColor = "White"
$IHSEXOL.BackColor = "Green"
$IHSEXOL.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$IHSEXOL.Add_Click({
    IHSEXOL
})
$IHSEXOL.Add_MouseEnter({
    $IHSEXOL.BackColor = "DarkRed"
})
$IHSEXOL.Add_MouseLeave({
    $IHSEXOL.BackColor = "Green"
})
$form.Controls.Add($IHSEXOL)

$IHSAZAD = New-Object System.Windows.Forms.Button
$IHSAZAD.Text = "Azure AD"
$IHSAZAD.AutoSize = $True
$IHSAZAD.Location = New-Object System.Drawing.Point(10, 100)
$IHSAZAD.ForeColor = "White"
$IHSAZAD.BackColor = "Green"
$IHSAZAD.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$IHSAZAD.Add_Click({
    IHSAZAD
})
$IHSAZAD.Add_MouseEnter({
    $IHSAZAD.BackColor = "DarkRed"
})
$IHSAZAD.Add_MouseLeave({
    $IHSAZAD.BackColor = "Green"
})
$form.Controls.Add($IHSAZAD)

$IHSMSGP = New-Object System.Windows.Forms.Button
$IHSMSGP.Text = "Microsoft Graph"
$IHSMSGP.AutoSize = $True
$IHSMSGP.Location = New-Object System.Drawing.Point(10, 130)
$IHSMSGP.ForeColor = "White"
$IHSMSGP.BackColor = "Green"
$IHSMSGP.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$IHSMSGP.Add_Click({
    IHSMSGP
})
$IHSMSGP.Add_MouseEnter({
    $IHSMSGP.BackColor = "DarkRed"
})
$IHSMSGP.Add_MouseLeave({
    $IHSMSGP.BackColor = "Green"
})
$form.Controls.Add($IHSMSGP)

$form.ShowDialog() | Out-Null

}

function IHS-ADGRPEXP-PT { 
    Add-Type -AssemblyName System.Windows.Forms

    # Create Form
    $IHSFMOU = New-Object System.Windows.Forms.Form
    $IHSFMOU.Text = "Input Group Names and File Name"
    $IHSFMOU.Size = New-Object System.Drawing.Size(300, 200)
    $IHSFMOU.StartPosition = "CenterScreen"

    # Label for Group Names
    $IHSLBOU = New-Object System.Windows.Forms.Label
    $IHSLBOU.Text = "Enter Group Names :"
    $IHSLBOU.AutoSize = $true
    $IHSLBOU.Location = New-Object System.Drawing.Point(10, 20)
    $IHSFMOU.Controls.Add($IHSLBOU)

    # Text Box for Group Names
    $groupTextBox = New-Object System.Windows.Forms.TextBox
    $groupTextBox.Location = New-Object System.Drawing.Point(10, 40)
    $groupTextBox.Size = New-Object System.Drawing.Size(260, 20)
    $IHSFMOU.Controls.Add($groupTextBox)

    # Label for File Name
    $fileLabel = New-Object System.Windows.Forms.Label
    $fileLabel.Text = "Enter Export File Name:"
    $fileLabel.AutoSize = $true
    $fileLabel.Location = New-Object System.Drawing.Point(10, 70)
    $IHSFMOU.Controls.Add($fileLabel)

    # Text Box for File Name
    $fileTextBox = New-Object System.Windows.Forms.TextBox
    $fileTextBox.Location = New-Object System.Drawing.Point(10, 90)
    $fileTextBox.Size = New-Object System.Drawing.Size(260, 20)
    $IHSFMOU.Controls.Add($fileTextBox)

    # OK Button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(50, 130)
    $okButton.Size = New-Object System.Drawing.Size(75, 23)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $IHSFMOU.Controls.Add($okButton)

    # Cancel Button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(150, 130)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $IHSFMOU.Controls.Add($cancelButton)

    # Show Dialog
    $result = $IHSFMOU.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # Read input values
        $IHSGPS = $groupTextBox.Text
        $fileName = $fileTextBox.Text.Trim()
        $currentDate = (Get-Date).ToString("dd-MM-yyyy")

        # Validate file name input
        if (-not $fileName) {
            [System.Windows.Forms.MessageBox]::Show("File name cannot be empty.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # Split input by comma, trim whitespace, and remove empty entries
        $groupNames = $IHSGPS -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

        # Array to store all group members
        $allGroupMembers = @()

        # Loop through each group name and gather members
        foreach ($groupName in $groupNames) {
            try {
                # Get group members and append to array
                $members = Get-ADGroupMember -Identity $groupName | 
                           Select-Object @{Name="GroupName";Expression={$groupName}}, SamAccountName
                $allGroupMembers += $members

                Write-Host "Collected members of group '$groupName'"
            } catch {
                Write-Host "Error collecting members for group '$groupName': $_" -ForegroundColor Red
            }
        }

        # Export all members to the specified CSV file
        $IHSEXPATH = "C:\IHS-Application\$fileName-$currentDate.csv"
        $allGroupMembers | Export-Csv -Path $IHSEXPATH -NoTypeInformation -Encoding UTF8

        # Show success message
        [System.Windows.Forms.MessageBox]::Show("Thanks for Using IHS Script. All group members have been successfully exported to $fileName.csv.", 
                                                "Export Successful", 
                                                [System.Windows.Forms.MessageBoxButtons]::OK, 
                                                [System.Windows.Forms.MessageBoxIcon]::Information)
    } else {
        Write-Host "Operation canceled." -ForegroundColor Red
    }
}

function IHSM365LCMU {
    
    $form = New-Object System.Windows.Forms.Form  
    $form.Text = "IHS-M365-Licenses-Menu"  
    $form.Size = New-Object System.Drawing.Size(400, 500)  
    $form.StartPosition = "CenterScreen"  
  
    $labelIHS = New-Object System.Windows.Forms.Label  
    $labelIHS.Text = "Welcome IHS M365 License Menu"  
    $labelIHS.AutoSize = $False  
    $labelIHS.Size = New-Object System.Drawing.Size(350, 30)  
    $labelIHS.Location = New-Object System.Drawing.Point(25, 10)  
    $labelIHS.Font = New-Object System.Drawing.Font("Cambria", 14, [System.Drawing.FontStyle]::Bold)  
    $labelIHS.ForeColor = "Black"  
  
    $labelSamAccount = New-Object System.Windows.Forms.Label  
    $labelSamAccount.Text = "Enter UGDN:"  
    $labelSamAccount.Location = New-Object System.Drawing.Point(20, 50)  
    $labelSamAccount.Size = New-Object System.Drawing.Size(80, 20)  
  
    $textBoxSamAccount = New-Object System.Windows.Forms.TextBox  
    $textBoxSamAccount.Location = New-Object System.Drawing.Point(100, 48)  
    $textBoxSamAccount.Size = New-Object System.Drawing.Size(200, 20)  

    $buttonGetNameTXT = New-Object System.Windows.Forms.TextBox  
    $buttonGetNameTXT.Location = New-Object System.Drawing.Point(100, 48)  
    $buttonGetNameTXT.Size = New-Object System.Drawing.Size(200, 20) 
  
    $labelDisplayName = New-Object System.Windows.Forms.Label  
    $labelDisplayName.Text = "Display Name:" 
    $labelDisplayName.Location = New-Object System.Drawing.Point(20, 80)  
    $labelDisplayName.Size = New-Object System.Drawing.Size(380, 30)  
  
    $labelCurrentLicenses = New-Object System.Windows.Forms.Label  
    $labelCurrentLicenses.Text = "Current Licenses:" 
    $labelCurrentLicenses.Location = New-Object System.Drawing.Point(20, 110)  
    $labelCurrentLicenses.Size = New-Object System.Drawing.Size(380, 60)  

     $licenses = @(
        @{ Name = "ENTERPRISEPACK"; Comment = "Microsoft 365 E3"},
        @{ Name = "STANDARDPACK"; Comment = "Microsoft 365 E1"},
        @{ Name = "SPE_F1"; Comment = "Microsoft 365 F3"},           
        @{ Name = "EMS"; Comment = "Enterprise Mobility + Security"},
        @{ Name = "EXCHANGESTANDARD"; Comment = "Exchange Online (Plan 1)" },
        @{ Name = "EXCHANGEARCHIVE_ADDON"; Comment = "Exchange Online Archiving"}
    )

    $checkboxes = @()
    $yPos = 140
    foreach ($license in $licenses) {
        $checkbox = New-Object System.Windows.Forms.CheckBox
        $checkbox.Text = "$($license.Name) - $($license.Comment)"
        $checkbox.Location = New-Object System.Drawing.Point(10, $yPos)
        $checkbox.Size = New-Object System.Drawing.Size(330, 30)
        $checkboxes += $checkbox
        $form.Controls.Add($checkbox)
        $yPos += 30
    }

    $buttonAssign = New-Object System.Windows.Forms.Button  
    $buttonAssign.Text = "Assign Licenses"  
    $buttonAssign.Location = New-Object System.Drawing.Point(25, $yPos)  
    $buttonAssign.Size = New-Object System.Drawing.Size(150, 30)  
    $buttonAssign.ForeColor = "White"  
    $buttonAssign.BackColor = "Green"  
    $form.Controls.Add($buttonAssign)  
  
    $buttonRevoke = New-Object System.Windows.Forms.Button  
    $buttonRevoke.Text = "Revoke Licenses"  
    $buttonRevoke.Location = New-Object System.Drawing.Point(200, $yPos)  
    $buttonRevoke.Size = New-Object System.Drawing.Size(150, 30)  
    $buttonRevoke.ForeColor = "White"  
    $buttonRevoke.BackColor = "Green"  
    $form.Controls.Add($buttonRevoke)  
  
    $statusButton = New-Object System.Windows.Forms.Button  
    $statusButton.Text = "Check Name And Licenses"  
    $statusButton.Location = New-Object System.Drawing.Point(110,350)  
    $statusButton.Size = New-Object System.Drawing.Size(150, 30)  
    $statusButton.ForeColor = "White"  
    $statusButton.BackColor = "Green"  
    $form.Controls.Add($statusButton)
      
$statusButton.Add_Click({
        $samAccountName = $textBoxSamAccount.Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($samAccountName)) {
            $user = Get-ADUser -Identity $samAccountName -Properties UserPrincipalName , DisplayName
            if ($user) {
                $userUPN = $user.UserPrincipalName
                $userUPN1 = $user.DisplayName
                $currentLicenses = Get-MgUserLicenseDetail -UserId $userUPN
                $labelDisplayName.Text = "Display Name: " + $userUPN1
                
                if ($currentLicenses) {
                    $licenseNames = $currentLicenses | ForEach-Object { $_.SkuPartNumber }
                    $labelCurrentLicenses.Text = "Current Licenses: " + ($licenseNames -join ", ")
                } else {
                    $labelCurrentLicenses.Text = "Current Licenses: None" 
                }
            } else {
                $labelCurrentLicenses.Text = "Current Licenses: User not found" 
            }
        } else {
            $labelCurrentLicenses.Text = "Current Licenses: "
        }
    })

    $buttonAssign.Add_Click({
        $samAccountName = $textBoxSamAccount.Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($samAccountName)) {
            $user = Get-ADUser -Identity $samAccountName -Properties UserPrincipalName, UserPrincipalName
            if ($user) {
                $userUPN = $user.UserPrincipalName
                $licensesToAssign = @()

                foreach ($checkbox in $checkboxes) {
                    if ($checkbox.Checked) {
                        $sku = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq $checkbox.Text.Split(' - ')[0] }
                        if ($sku) {
                            $licensesToAssign += @{ SkuId = $sku.SkuId }
                        } else {
                            Write-Host "Warning: SKU '$($checkbox.Text.Split(' - ')[0])' not found."
                        }
                    }
                }

                if ($licensesToAssign.Count -gt 0) {
                    try {
                        Set-MgUserLicense -UserId $userUPN -AddLicenses $licensesToAssign -RemoveLicenses @()
                        [System.Windows.Forms.MessageBox]::Show("Licenses assigned successfully to $userUPN", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    } catch {
                        [System.Windows.Forms.MessageBox]::Show("Failed to assign licenses to $userUPN : $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    }
                } else {
                    [System.Windows.Forms.MessageBox]::Show("No licenses selected for assignment.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show("User not found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please enter a UGDN.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    $buttonRevoke.Add_Click({
        $samAccountName = $textBoxSamAccount.Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($samAccountName)) {
            $user = Get-ADUser -Identity $samAccountName -Properties UserPrincipalName
            if ($user) {
                $userUPN = $user.UserPrincipalName
                $licensesToRevoke = @()

                foreach ($checkbox in $checkboxes) {
                    if ($checkbox.Checked) {
                        $sku = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq $checkbox.Text.Split(' - ')[0] }
                        if ($sku) {
                            $licensesToRevoke += @($sku.SkuId)
                        } else {
                            Write-Host "Warning: SKU '$($checkbox.Text.Split(' - ')[0])' not found."
                        }
                    }
                }

                if ($licensesToRevoke.Count -gt 0) {
                    try {
                        Set-MgUserLicense -UserId $userUPN -RemoveLicenses $licensesToRevoke -AddLicenses @{}
                        [System.Windows.Forms.MessageBox]::Show("Licenses revoked successfully from $userUPN", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    } catch {
                        [System.Windows.Forms.MessageBox]::Show("Failed to revoke licenses from $userUPN : $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    }
                } else {
                    [System.Windows.Forms.MessageBox]::Show("No licenses selected for revocation.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show("User not found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please enter a UGDN.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    # Add controls to the form  
    $form.Controls.Add($labelIHS)  
    $form.Controls.Add($labelSamAccount)  
    $form.Controls.Add($textBoxSamAccount)  
    $form.Controls.Add($buttonGetName)  
    $form.Controls.Add($labelDisplayName)  
    $form.Controls.Add($labelCurrentLicenses)

   # Show the form  
    $form.ShowDialog() | Out-Null  
}

function IHS-BLKLICTASK {
    $csvPath = "C:\IHS-Application\IHS-Template.csv"
    $outputCsvPath = "C:\IHS-Application\IHS-License-TaskStatus.csv"

    Connect-MgGraph -Scopes User.ReadWrite.All, Organization.Read.All

    $users = Import-Csv -Path $csvPath

    $availableLicenses = @(
        @{ Name = "ENTERPRISEPACK"; Comment = "Microsoft 365 E3" },
        @{ Name = "STANDARDPACK"; Comment = "Microsoft 365 E1" },
        @{ Name = "SPE_F1"; Comment = "Microsoft 365 F3" },           
        @{ Name = "EMS"; Comment = "Enterprise Mobility + Security" },
        @{ Name = "EXCHANGESTANDARD"; Comment = "Exchange Online (Plan 1)" },
        @{ Name = "EXCHANGEARCHIVE_ADDON"; Comment = "Exchange Online Archiving" }
    )

    $outputData = @()

    foreach ($user in $users) {
        $samAccountName = $user.UserName
        $assignLicenses = $user.AssignLicenses.Split(",") -replace " ", ""
        $revokeLicenses = $user.RevokeLicenses.Split(",") -replace " ", ""

        $adUser = Get-ADUser -Filter {SamAccountName -eq $samAccountName} -Properties UserPrincipalName
        if ($adUser) {
            $userUPN = $adUser.UserPrincipalName
            $licensesToAssign = @()
            $licensesToRevoke = @()
            $status = "Success"

            foreach ($license in $assignLicenses) {
                if ($license) {
                    $sku = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq $license }
                    if ($sku) {
                        $licensesToAssign += @{ SkuId = $sku.SkuId }
                    } else {
                        Write-Host "Warning: License '$license' not found for $samAccountName."
                    }
                }
            }

            foreach ($license in $revokeLicenses) {
                if ($license) {
                    $sku = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq $license }
                    if ($sku) {
                        $licensesToRevoke += $sku.SkuId
                    } else {
                        Write-Host "Warning: License '$license' not found for $samAccountName."
                    }
                }
            }

            if ($licensesToAssign.Count -gt 0 -or $licensesToRevoke.Count -gt 0) {
                try {
                    Set-MgUserLicense -UserId $userUPN -AddLicenses $licensesToAssign -RemoveLicenses $licensesToRevoke
                    Write-Host "Licenses processed for $samAccountName Assigned = $($assignLicenses -join ', '); Revoked = $($revokeLicenses -join ', ')"
                } catch {
                    Write-Host "Error processing licenses for $samAccountName $_"
                    $status = "Error: $_"
                }
            } else {
                Write-Host "No licenses to assign or revoke for $samAccountName."
                $status = "No changes"
            }

            $outputData += [pscustomobject]@{
                UserName        = $samAccountName
                UserPrincipalName = $userUPN
                AssignedLicenses = $assignLicenses -join ', '
                RevokedLicenses  = $revokeLicenses -join ', '
                Status           = $status
            }

        } else {
            Write-Host "User $samAccountName not found in Active Directory."
            $outputData += [pscustomobject]@{
                UserName          = $samAccountName
                UserPrincipalName = "N/A"
                AssignedLicenses  = $assignLicenses -join ', '
                RevokedLicenses   = $revokeLicenses -join ', '
                Status            = "User not found in AD"
            }
        }
    }

    $outputData | Export-Csv -Path $outputCsvPath -NoTypeInformation -Force
    Write-Host "License processing status exported to $outputCsvPath"
}

function IHS-USBDLICSTS { $IHSUSINLICSTS = "C:\IHS-Application\IHS-Template.csv"
$IHSUSOULICSTS = "C:\IHS-Application\USER-Based-Licenses-Status.csv"
$samAccountNames = Import-Csv -Path $IHSUSINLICSTS
Connect-MgGraph
$results = @()

foreach ($user in $samAccountNames) {
    $samAccountName = $user.Username

    if ($samAccountName) {
        try {
            $adUser = Get-ADUser -Identity $samAccountName -Properties UserPrincipalName -ErrorAction Stop
            
            if ($adUser) {
                $userUPN = $adUser.UserPrincipalName
                
                $currentLicenses = Get-MgUserLicenseDetail -UserId $userUPN
                
                if ($currentLicenses) {
                    $licenseNames = $currentLicenses | ForEach-Object { $_.SkuPartNumber }
                    $licenseDetails = ($licenseNames -join ", ")
                } else {
                    $licenseDetails = "None"
                }

                $result = [PSCustomObject]@{
                    SamAccountName = $samAccountName
                    UserPrincipalName = $userUPN
                    Licenses = $licenseDetails
                }

                $results += $result
            }
        } catch {
            $result = [PSCustomObject]@{
                SamAccountName = $samAccountName
                UserPrincipalName = "Not Found"
                Licenses = "Not Found"
            }
            $results += $result
            Write-Host "User not found: $samAccountName"
        }
    } else {
        Write-Host "SamAccountName is empty or invalid."
    }
}

$results | Export-Csv -Path $IHSUSOULICSTS -NoTypeInformation

Write-Host "License details exported to $IHSUSOULICSTS"
} 

function IHS-USBDFWDSTS {

$userListPath = "C:\IHS-Application\IHS-Template.csv"
$outputPath = "C:\IHS-Application\IHS-Forwarding-Report-$(Get-Date -Format dd-MM-yy_hh-mm).csv"

$userList = Import-Csv -Path $userListPath | Select-Object -ExpandProperty UserPrincipalName

$forwardingReport = @()

foreach ($user in $userList) {
    try {
        $mailboxSettings = Get-MgUserMailboxSetting -UserId $user -ErrorAction SilentlyContinue

        if ($mailboxSettings.ForwardingSmtpAddress) {
            $forwardingReport += [PSCustomObject]@{
                UserPrincipalName          = $user
                ForwardingSMTPAddress      = $mailboxSettings.ForwardingSmtpAddress
                KeepCopyInMailbox          = $mailboxSettings.KeepForwardedMessages
                ForwardingEnabled          = $true
            }
            Write-Host "Forwarding enabled for user: $user" -ForegroundColor Green
        } else {
            $forwardingReport += [PSCustomObject]@{
                UserPrincipalName          = $user
                ForwardingSMTPAddress      = "No forwarding"
                KeepCopyInMailbox          = $null
                ForwardingEnabled          = $false
            }
        }
    } catch {
        Write-Host "Error retrieving forwarding settings for user: $user - $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

if ($forwardingReport.Count -gt 0) {
    $forwardingReport | Export-Csv -Path $outputPath -NoTypeInformation -Force
    Write-Host "Forwarding report generated at: $outputPath" -ForegroundColor Green
} else {
    Write-Host "No forwarding settings found for any users." -ForegroundColor Yellow
}

}

function IHS-USBDDLGSTS {
function Check-ExchangeOnlineModule { if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {  Write-Host "Exchange Online Management module is not installed. Installing it now..."  Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser  } else {  Write-Host "Exchange Online Management module is already installed."  }}
function Connect-ToExchangeOnlineFWD {  try {  Connect-ExchangeOnline  } catch { Write-Host "Error connecting to Exchange Online: $_"  exit 1 }  Write-Host "Successfully connected to Exchange Online."  }
Check-ExchangeOnlineModule   
Connect-ToExchangeOnlineFWD

$userListPath = "C:\IHS-Application\IHS-Template.csv"
$outputPath = "C:\IHS-Application\IHS-CombinedDelegationReport-$(Get-Date -Format yyyy-MM-dd_hh-mm-ss).csv"

$userList = Import-Csv -Path $userListPath | Select-Object -ExpandProperty UserPrincipalName

$combinedReport = @()

foreach ($user in $userList) {
    $fullAccessPermissions = Get-MailboxPermission -Identity $user -ErrorAction SilentlyContinue | Where-Object {
        $_.AccessRights -contains "FullAccess" -and $_.User -notlike "NT AUTHORITY\SELF"
    } | ForEach-Object {
        [PSCustomObject]@{
            Mailbox        = $user
            DelegatedUser  = $_.User
            PermissionType = "FullAccess"
            AccessRights   = $_.AccessRights -join ", "
        }
    }
    $combinedReport += $fullAccessPermissions

    $sendAsPermissions = Get-RecipientPermission -Identity $user -ErrorAction SilentlyContinue | Where-Object {
        $_.AccessRights -contains "SendAs" -and $_.Trustee -notlike "NT AUTHORITY\SELF"
    } | ForEach-Object {
        [PSCustomObject]@{
            Mailbox        = $user
            DelegatedUser  = $_.Trustee
            PermissionType = "SendAs"
            AccessRights   = $_.AccessRights -join ", "
        }
    }
    $combinedReport += $sendAsPermissions

    $mailbox = Get-Mailbox -Identity $user -ErrorAction SilentlyContinue
    if ($mailbox -and $mailbox.GrantSendOnBehalfTo) {
        $sendOnBehalfPermissions = $mailbox.GrantSendOnBehalfTo | ForEach-Object {
            [PSCustomObject]@{
                Mailbox        = $user
                DelegatedUser  = $_.Name
                PermissionType = "SendOnBehalf"
                AccessRights   = "SendOnBehalf"
            }
        }
        $combinedReport += $sendOnBehalfPermissions
    }
}

$combinedReport | Export-Csv -Path $outputPath -NoTypeInformation -Force

Write-Host "Combined delegation report generated at: $outputPath" -ForegroundColor Green


}

function IHSMDIHS{

function New-StyledButton {
    param (
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$Width = 180,
        [int]$Height = 40
    )
    $button = New-Object System.Windows.Forms.Button
    $button.Location = New-Object System.Drawing.Point($X, $Y)
    $button.Size = New-Object System.Drawing.Size($Width, $Height)
    $button.Text = $Text
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.FlatAppearance.BorderSize = 0
    $button.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
    $button.ForeColor = [System.Drawing.Color]::White
    $button.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $button.Add_MouseEnter({ $this.BackColor = [System.Drawing.Color]::FromArgb(100, 160, 210) })
    $button.Add_MouseLeave({ $this.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180) })
    return $button
}

function New-StyledForm {
    param (
        [string]$Title,
        [int]$Width,
        [int]$Height
    )
    $IHSFMMD = New-Object System.Windows.Forms.Form
    $IHSFMMD.Text = $Title
    $IHSFMMD.Size = New-Object System.Drawing.Size($Width, $Height)
    $IHSFMMD.StartPosition = "CenterScreen"
    $IHSFMMD.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $IHSFMMD.MaximizeBox = $false
    $IHSFMMD.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    return $IHSFMMD
}

$IHSFMMD = New-StyledForm -Title "Module Check and Connect" -Width 600 -Height 220
$IHSFMMD.MaximizeBox = $false
$IHSFMMD.MinimizeBox = $false
$IHSFMMD.BackColor = [System.Drawing.Color]::White
$IHSFMMD.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

$mainPanel = New-Object System.Windows.Forms.TableLayoutPanel
$mainPanel.Dock = "Fill"
$mainPanel.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 10)
$mainPanel.ColumnCount = 4
$mainPanel.RowCount = 4
$mainPanel.BackColor = [System.Drawing.Color]::White
$mainPanel.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::Single

$mainPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
$mainPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
$mainPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
$mainPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))

0..3 | ForEach-Object {
    $mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 25)))
}

$headerStyle = @{
    Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
}

$headers = @("Service", "Module Status", "Connection Status", "Action")
0..3 | ForEach-Object {
    $header = New-Object System.Windows.Forms.Label
    $header.Text = $headers[$_]
    $header.Font = $headerStyle.Font
    $header.BackColor = $headerStyle.BackColor
    $header.ForeColor = $headerStyle.ForeColor
    $header.Dock = "Fill"
    $header.TextAlign = "MiddleCenter"
    $mainPanel.Controls.Add($header, $_, 0)
}

$services = @{
    "Azure AD" = @{
        ModuleName = "AzureAD"
        ConnectCmd = { Connect-AzureAD -ShowBanner:$false }
    }
    "Exchange Online" = @{
        ModuleName = "ExchangeOnlineManagement"
        ConnectCmd = { Connect-ExchangeOnline -ShowBanner:$false }
    }
    "Microsoft Graph" = @{
        ModuleName = "Microsoft.Graph"
        ConnectCmd = { Connect-MgGraph -ShowBanner:$false  } #-Scopes "User.Read.All" 
    }
}

$labelStyle = @{
    Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    BackColor = [System.Drawing.Color]::White
    TextAlign = "MiddleCenter"
    Dock = "Fill"
    Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 3)
}

# Add Service Rows
$row = 1
foreach ($service in $services.Keys) {
    $serviceLabel = New-Object System.Windows.Forms.Label
    $serviceLabel.Text = $service
    $serviceLabel.Font = $labelStyle.Font
    $serviceLabel.BackColor = $labelStyle.BackColor
    $serviceLabel.TextAlign = $labelStyle.TextAlign
    $serviceLabel.Dock = $labelStyle.Dock
    $serviceLabel.Margin = $labelStyle.Margin
    $mainPanel.Controls.Add($serviceLabel, 0, $row)

    $moduleStatus = New-Object System.Windows.Forms.Label
    $moduleStatus.Font = $labelStyle.Font
    $moduleStatus.BackColor = $labelStyle.BackColor
    $moduleStatus.TextAlign = $labelStyle.TextAlign
    $moduleStatus.Dock = $labelStyle.Dock
    $moduleStatus.Margin = $labelStyle.Margin
    if (Get-Module -ListAvailable $services[$service].ModuleName) {
        $moduleStatus.Text = "✓ Installed"
        $moduleStatus.ForeColor = [System.Drawing.Color]::Green
    } else {
        $moduleStatus.Text = "✗ Not Installed"
        $moduleStatus.ForeColor = [System.Drawing.Color]::Red
    }
    $mainPanel.Controls.Add($moduleStatus, 1, $row)

    $connectionStatus = New-Object System.Windows.Forms.Label
    $connectionStatus.Text = "Not Connected"
    $connectionStatus.Font = $labelStyle.Font
    $connectionStatus.BackColor = $labelStyle.BackColor
    $connectionStatus.TextAlign = $labelStyle.TextAlign
    $connectionStatus.Dock = $labelStyle.Dock
    $connectionStatus.Margin = $labelStyle.Margin
    $connectionStatus.ForeColor = [System.Drawing.Color]::Red
    $mainPanel.Controls.Add($connectionStatus, 2, $row)

    $connectButton = New-Object System.Windows.Forms.Button
    $connectButton.Text = "Connect"
    $connectButton.Font = $labelStyle.Font
    $connectButton.Dock = "Fill"
    $connectButton.Margin = New-Object System.Windows.Forms.Padding(10, 5, 10, 5)
    $connectButton.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $connectButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $connectButton.Tag = @{
        Service = $service
        StatusLabel = $connectionStatus
    }

    $connectButton.Add_Click({
        $serviceInfo = $this.Tag
        $service = $serviceInfo.Service
        $status = $serviceInfo.StatusLabel
        
        $this.Enabled = $false
        $status.Text = "⟳ Connecting..."
        $status.ForeColor = [System.Drawing.Color]::Blue
        
        Start-Job -Name "Connect_$service" -ScriptBlock {
            param($serviceName, $moduleData)
            Import-Module $moduleData.ModuleName -Force
            & $moduleData.ConnectCmd
        } -ArgumentList $service, $services[$service]
    })
    
    $mainPanel.Controls.Add($connectButton, 3, $row)
    $row++
}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 500
$timer.Add_Tick({
    Get-Job | Where-Object { $_.Name -like "Connect_*" } | ForEach-Object {
        $serviceName = $_.Name -replace "Connect_"
        $row = [array]::IndexOf(($services.Keys), $serviceName) + 1
        $button = $mainPanel.GetControlFromPosition(3, $row)
        $status = $mainPanel.GetControlFromPosition(2, $row)
        
        if ($_.State -eq "Completed") {
            $status.Text = "✓ Connected"
            $status.ForeColor = [System.Drawing.Color]::Green
            $button.Text = "Connected"
            $button.BackColor = [System.Drawing.Color]::FromArgb(225, 240, 225)
            Remove-Job $_
        }
        elseif ($_.State -eq "Failed") {
            $error = Receive-Job $_
            $status.Text = "✗ Failed"
            $status.ForeColor = [System.Drawing.Color]::Red
            $button.Text = "Connect"
            $button.Enabled = $true
            Remove-Job $_
            [System.Windows.Forms.MessageBox]::Show($error, "Connection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})
$timer.Start()

$IHSFMMD.Controls.Add($mainPanel)
$IHSFMMD.Add_FormClosing({ 
    $timer.Stop()
    Get-Job | Remove-Job -Force
})
[void]$IHSFMMD.ShowDialog()

}


function Show-MainForm {

function Get-ImageFromUrl1 {
    param (
        [string]$Url
    )
    try {
        $webClient = New-Object System.Net.WebClient
        $imageStream = $webClient.OpenRead($Url)
        $image = [System.Drawing.Image]::FromStream($imageStream)
        $imageStream.Close()
        return $image
    } catch {
        Write-Error "Failed to load image from $Url. $_"
        return $null
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Ibrahim UPL Scripts = Advanced IHS Menu"
$form.AutoSize = $True
#$form.BackColor = [System.Drawing.Color]::LightBlue
$form.StartPosition = "CenterScreen"

$backgroundPanel1 = New-Object System.Windows.Forms.Panel
$backgroundPanel1.Dock = [System.Windows.Forms.DockStyle]::Fill
$formLogin.Controls.Add($backgroundPanel1)

$backgroundImageUrl1 = "https://github.com/Imran1010/Applogin/blob/main/Untitled.jpg?raw=true"
$backgroundImage1 = Get-ImageFromUrl1 -Url $backgroundImageUrl1
if ($backgroundImage1) {
    $backgroundPanel1.BackgroundImage = $backgroundImage1
    $backgroundPanel1.BackgroundImageLayout = [System.Windows.Forms.ImageLayout]::Stretch
}

$controlsPanel = New-Object System.Windows.Forms.Panel
$controlsPanel.Size = New-Object System.Drawing.Size(300, 210)
$controlsPanel.Location = New-Object System.Drawing.Point(112, 75)  # Center the panel
$controlsPanel.BackColor = [System.Drawing.Color]::FromArgb(80, [System.Drawing.Color]::White)  
$backgroundPanel1.Controls.Add($controlsPanel)

$labelIHS = New-Object System.Windows.Forms.Label
$labelIHS.Text = "IHS UPL Script : AD Menu"
$labelIHS.AutoSize = $True
$labelIHS.Size = New-Object System.Drawing.Size(250, 30) 
$labelIHS.Location = New-Object System.Drawing.Point(6, 10)
$labelIHS.Font = New-Object System.Drawing.Font("Cambria",14,[System.Drawing.FontStyle]::Bold)
$labelIHS.ForeColor = "Black"

$helpPanel = New-Object System.Windows.Forms.Panel
$helpPanel.Size = New-Object System.Drawing.Size(200, 330)
$helpPanel.Location = New-Object System.Drawing.Point(480, 10)
$helpPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$helpPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)

$helpLabel = New-Object System.Windows.Forms.Label
$helpLabel.Size = New-Object System.Drawing.Size(200, 300)
$helpLabel.Location = New-Object System.Drawing.Point(10, 10)
$helpLabel.Font = New-Object System.Drawing.Font("Calibri", 10)
$helpLabel.Text = "Hover over any button to see its description here..."
$helpPanel.Controls.Add($helpLabel)

    $logoutButton = New-Object System.Windows.Forms.Button
    $logoutButton.Text = "Logout"
    $logoutButton.Location = New-Object System.Drawing.Point(380, 10)
    $logoutButton.ForeColor = "White"
    $logoutButton.BackColor = "Gray"
    $logoutButton.Add_Click({          
                $form.Hide()  
                Show-LoginForm  
                $form.Close()
    })
    $logoutButton.Add_MouseEnter({ $logoutButton.BackColor = "DarkRed" })
    $logoutButton.Add_MouseLeave({ $logoutButton.BackColor = "Gray" })

     $IHDNDLBT = New-Object System.Windows.Forms.Button
    $IHDNDLBT.Text = "Module Check"
    $IHDNDLBT.Location = New-Object System.Drawing.Point(280, 10)
    $IHDNDLBT.AutoSize = $True
    $IHDNDLBT.ForeColor = "White"
    $IHDNDLBT.BackColor = "Gray"
    $IHDNDLBT.Add_Click({
             Show-ModuleStatusForm    
    })
    $IHDNDLBT.Add_MouseEnter({ $IHDNDLBT.BackColor = "DarkRed" })
    $IHDNDLBT.Add_MouseLeave({ $IHDNDLBT.BackColor = "Gray" })

$script1Button = New-Object System.Windows.Forms.Button
$script1Button.Text = "User Password Reset"
$script1Button.AutoSize = $True
$script1Button.Location = New-Object System.Drawing.Point(10, 40)
$script1Button.ForeColor = "White"
$script1Button.BackColor = "Green"
$script1Button.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$script1Button.Add_Click({
    IHS-USPWRST-MLUP
})
$script1Button.Add_MouseEnter({
    $script1Button.BackColor = "DarkRed"
    $helpLabel.Text = "User Password Reset`n`nThis tool allows administrators to reset user passwords in Active Directory.`n`nFeatures:`n- Secure password generation`n- Immediate password reset`n- Forces password change at next logon`n- Logs password reset actions"
})
$script1Button.Add_MouseLeave({
    $script1Button.BackColor = "Green"
    $helpLabel.Text = "Hover over any button to see its description here..."
})

$EmailUPIHS = New-Object System.Windows.Forms.Button
$EmailUPIHS.Text = "User Email Update"
$EmailUPIHS.AutoSize = $True
$EmailUPIHS.Location = New-Object System.Drawing.Point(270, 40)
$EmailUPIHS.ForeColor = "White"
$EmailUPIHS.BackColor = "Green"
$EmailUPIHS.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$EmailUPIHS.Add_Click({
    IHS-USEMUP
})
$EmailUPIHS.Add_MouseEnter({
    $EmailUPIHS.BackColor = "DarkRed"
     $helpLabel.Text = "User Email Update`n`nUpdates user email addresses in Active Directory.`n`nFeatures:`n- Batch email updates`n- Validation of email format`n- Updates primary and proxy addresses`n- Syncs with Exchange"
})
$EmailUPIHS.Add_MouseLeave({
    $EmailUPIHS.BackColor = "Green"
    $helpLabel.Text = "Hover over any button to see its description here..."
})


$CMTIHS = New-Object System.Windows.Forms.Button
$CMTIHS.Text = "Comment Reset"
$CMTIHS.AutoSize = $True
$CMTIHS.Location = New-Object System.Drawing.Point(270, 70)
$CMTIHS.ForeColor = "White"
$CMTIHS.BackColor = "Green"
$CMTIHS.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$CMTIHS.Add_Click({
    IHS-USCMTRST
})
$CMTIHS.Add_MouseEnter({
    $CMTIHS.BackColor = "DarkRed"
    $helpLabel.Text = "Comment Reset`n`nManages and resets user comment fields in Active Directory.`n`nFeatures:`n- Clear existing comments`n- Add new standardized comments`n- Bulk comment updates`n- Comment history tracking"
})
$CMTIHS.Add_MouseLeave({
    $CMTIHS.BackColor = "Green"
    $helpLabel.Text = "Hover over any button to see its description here..."
})

$GPLFTIHS = New-Object System.Windows.Forms.Button
$GPLFTIHS.Text = "Group Lifting After Mapping"
$GPLFTIHS.AutoSize = $True
$GPLFTIHS.Location = New-Object System.Drawing.Point(10, 70)
$GPLFTIHS.ForeColor = "White"
$GPLFTIHS.BackColor = "Green"
$GPLFTIHS.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$GPLFTIHS.Add_Click({
    IHS-GRPLFT
})
$GPLFTIHS.Add_MouseEnter({
    $GPLFTIHS.BackColor = "DarkRed"
$helpLabel.Text = "Group Lifting`n`nPurpose: Manages automatic group migration from OLD UGDN to NEW UGDN in Active Directory.`nFeatures :- Automatically detects all groups from OLD UGDN `n- Creates corresponding groups in NEW UGDN `n- Validates successful migration `n3. Transfers all group memberships `nNote: For assistance, contact Ibrahim (India.coreitsupport@support.com)"
})
$GPLFTIHS.Add_MouseLeave({
    $GPLFTIHS.BackColor = "Green"
    $helpLabel.Text = "Hover over any button to see its description here..."
})

$script2Button = New-Object System.Windows.Forms.Button
$script2Button.Text = "AD User Dump IHS Report "
$script2Button.AutoSize = $True
$script2Button.Location = New-Object System.Drawing.Point(10, 100)
$script2Button.ForeColor = "White"
$script2Button.BackColor = "Green"
$script2Button.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$script2Button.Add_Click({
$date = Get-Date -Format "yyyyMMdd"
$IHSEXPATH = 'c:\ADUserexport -$date .csv’
    IHS-ADEXP-PT -ExportPath $IHSEXPATH
})
$script2Button.Add_MouseEnter({
    $script2Button.BackColor = "DarkRed"
    $helpLabel.Text = "AD User Dump IHS Report`n`nPurpose: Export All user in Active Directory With UPL Attributes.`nFeatures :- Automatically detects all Members from AD`n- Export with in 15 min in one click`n`nNote: For assistance, contact Ibrahim (India.coreitsupport@support.com)"
})
$script2Button.Add_MouseLeave({
    $script2Button.BackColor = "Green"
    $helpLabel.Text = "Hover over any button to see its description here..."
})

$script3Button = New-Object System.Windows.Forms.Button
$script3Button.Text = "AD Export OU BASED"
$script3Button.AutoSize = $True
$script3Button.Location = New-Object System.Drawing.Point(270, 100)
$script3Button.ForeColor = "White"
$script3Button.BackColor = "Green"
$script3Button.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$script3Button.Add_Click({
    IHS-ADOUEXP-PT
})
$script3Button.Add_MouseEnter({
    $script3Button.BackColor = "DarkRed"
    $helpLabel.Text = "AD OU Users detail export`n`nPurpose: Export user based on OU ,Its will export users from in Active Directory With UPL Attributes.`nFeatures :- Automatically detects OU & export Users `n- Export Member detail based on OU`n`nNote: For assistance, contact Ibrahim (India.coreitsupport@support.com)"
})
$script3Button.Add_MouseLeave({
    $script3Button.BackColor = "Green"
    $helpLabel.Text = "Hover over any button to see its description here..."
})

$script4Button = New-Object System.Windows.Forms.Button
$script4Button.Text = "AD Export as List based"
$script4Button.AutoSize = $True
$script4Button.Location = New-Object System.Drawing.Point(10, 130)
$script4Button.ForeColor = "White"
$script4Button.BackColor = "Green"
$script4Button.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$script4Button.Add_Click({
    IHSADLISTTY
})
$script4Button.Add_MouseEnter({
    $script4Button.BackColor = "DarkRed"
    $helpLabel.Text = "AD user detail Based on user list`n`nPurpose: Export user based on template fill the UGDN in template its will export users from in Active Directory With UPL Attributes.`nFeatures :- Automatically detects Users from Template Members `n- Export Member detail based on users List`n`nNote: For assistance, contact Ibrahim (India.coreitsupport@support.com)"
})
$script4Button.Add_MouseLeave({
    $script4Button.BackColor = "Green"
    $helpLabel.Text = "Hover over any button to see its description here..."
})


$IHSMDINT = New-Object System.Windows.Forms.Button
$IHSMDINT.Text = "Modules Installer"
$IHSMDINT.AutoSize = $True
$IHSMDINT.Location = New-Object System.Drawing.Point(10, 160)
$IHSMDINT.ForeColor = "White"
$IHSMDINT.BackColor = "Green"
$IHSMDINT.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$IHSMDINT.Add_Click({
    IHS-MDINT
    
})
$IHSMDINT.Add_MouseEnter({
    $IHSMDINT.BackColor = "DarkRed"
        $helpLabel.Text = "One Click Module Install`n`nPurpose: We can Install Module on One Click`nFeatures :- RSAT Windows 10 & above`n- `n Azure AD Module `nExchange Online Module `nMS Graph Module  `nNote: For assistance, contact Ibrahim (India.coreitsupport@support.com)"
})
$IHSMDINT.Add_MouseLeave({
    $IHSMDINT.BackColor = "Green"
    $helpLabel.Text = "Hover over any button to see its description here..."
})

$IHSGRPEXP = New-Object System.Windows.Forms.Button
$IHSGRPEXP.Text = "Group Users Export"
$IHSGRPEXP.AutoSize = $True
$IHSGRPEXP.Location = New-Object System.Drawing.Point(270, 130)
$IHSGRPEXP.ForeColor = "White"
$IHSGRPEXP.BackColor = "Green"
$IHSGRPEXP.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$IHSGRPEXP.Add_Click({

    IHS-ADGRPEXP-PT
    
})
$IHSGRPEXP.Add_MouseEnter({
    $IHSGRPEXP.BackColor = "DarkRed"
            $helpLabel.Text = "Groups User Export`n`nPurpose: Export user based on Group Name from in Active Directory .`nFeatures :- Automatically detects Users from Groups Members`n`nNote: For assistance, contact Ibrahim (India.coreitsupport@support.com)"
})
$IHSGRPEXP.Add_MouseLeave({
    $IHSGRPEXP.BackColor = "Green"
    $helpLabel.Text = "Hover over any button to see its description here..."
})

$IHSGRPTSK = New-Object System.Windows.Forms.Button
$IHSGRPTSK.Text = "Multiple Group Task"
$IHSGRPTSK.AutoSize = $True
$IHSGRPTSK.Location = New-Object System.Drawing.Point(270, 160)
$IHSGRPTSK.ForeColor = "White"
$IHSGRPTSK.BackColor = "Green"
$IHSGRPTSK.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$IHSGRPTSK.Add_Click({

    IHS-MLTGRP
    
})
$IHSGRPTSK.Add_MouseEnter({
    $IHSGRPTSK.BackColor = "DarkRed"
    $helpLabel.Text = "Groups User Export`n`nPurpose: We can Add & remove group in bulk format using template.`nFeatures :- Fill the template based on Groups Members`n`nNote: For assistance, contact Ibrahim (India.coreitsupport@support.com)"
})
$IHSGRPTSK.Add_MouseLeave({
    $IHSGRPTSK.BackColor = "Green"
    $helpLabel.Text = "Hover over any button to see its description here..."
})

$M365IHS = New-Object System.Windows.Forms.Label
$M365IHS.Text = "IHS UPL Script : M365 Menu"
$M365IHS.AutoSize = $True
$M365IHS.Size = New-Object System.Drawing.Size(250, 30) # Set the desired width and height
$M365IHS.Location = New-Object System.Drawing.Point(6, 200)
$M365IHS.Font = New-Object System.Drawing.Font("Cambria",14,[System.Drawing.FontStyle]::Bold)
$M365IHS.ForeColor = "Black"


$M365IHSIMID = New-Object System.Windows.Forms.Button
$M365IHSIMID.Text = "Show Immutable ID"
$M365IHSIMID.AutoSize = $True
$M365IHSIMID.Location = New-Object System.Drawing.Point(10, 230)
$M365IHSIMID.ForeColor = "White"
$M365IHSIMID.BackColor = "Green"
$M365IHSIMID.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$M365IHSIMID.Add_Click({
    IHD-USSHIMUT
})
$M365IHSIMID.Add_MouseEnter({
    $M365IHSIMID.BackColor = "DarkRed"
        $helpLabel.Text = "One Click Immutable ID`n`nPurpose: Just Need UGDN one Click Immutable ID Copy`nFeatures :- Auto Copy just put UGDN & Enter`n`nNote: For assistance, contact Ibrahim (India.coreitsupport@support.com)"

})
$M365IHSIMID.Add_MouseLeave({
    $M365IHSIMID.BackColor = "Green"
    $helpLabel.Text = "Hover over any button to see its description here..."
})

$IHSUSBDLICSTATUS = New-Object System.Windows.Forms.Button
$IHSUSBDLICSTATUS.Text = "User based License status"
$IHSUSBDLICSTATUS.AutoSize = $True
$IHSUSBDLICSTATUS.Location = New-Object System.Drawing.Point(10, 260)
$IHSUSBDLICSTATUS.ForeColor = "White"
$IHSUSBDLICSTATUS.BackColor = "Green"
$IHSUSBDLICSTATUS.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$IHSUSBDLICSTATUS.Add_Click({
    IHS-USBDLICSTS
 [System.Windows.Forms.MessageBox]::Show("License details exported to $IHSUSOULICSTS", "Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)


})
$IHSUSBDLICSTATUS.Add_MouseEnter({
    $IHSUSBDLICSTATUS.BackColor = "DarkRed"
        $helpLabel.Text = "User Licenses detail based on Users`n`nPurpose: Fill the template based on User we get the licenses Detail`nFeatures :- One click licenses detail Exported to App Folder`n`nNote: For assistance, contact Ibrahim (India.coreitsupport@support.com)"

})
$IHSUSBDLICSTATUS.Add_MouseLeave({
    $IHSUSBDLICSTATUS.BackColor = "Green"
    $helpLabel.Text = "Hover over any button to see its description here..."
})


$IHSM365LC = New-Object System.Windows.Forms.Button
$IHSM365LC.Text = "IHS License Menu"
$IHSM365LC.AutoSize = $True
$IHSM365LC.Location = New-Object System.Drawing.Point(270, 230)
$IHSM365LC.ForeColor = "White"
$IHSM365LC.BackColor = "Green"
$IHSM365LC.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$IHSM365LC.Add_Click({

IHSM365LCMU


})
$IHSM365LC.Add_MouseEnter({
    $IHSM365LC.BackColor = "DarkRed"
        $helpLabel.Text = "License assigned based on UGDN`n`nPurpose: We check UGDN licenses assigned & revoked easly but possible in Entra Network`nFeatures :- Chcek & assigned & Revoke licenses based on UGDN`n`nNote: For assistance, contact Ibrahim (India.coreitsupport@support.com)"

})
$IHSM365LC.Add_MouseLeave({
    $IHSM365LC.BackColor = "Green"
    $helpLabel.Text = "Hover over any button to see its description here..."
})


$IHSBLKLIC = New-Object System.Windows.Forms.Button
$IHSBLKLIC.Text = "IHS Bulk License Task"
$IHSBLKLIC.AutoSize = $True
$IHSBLKLIC.Location = New-Object System.Drawing.Point(270, 260)
$IHSBLKLIC.ForeColor = "White"
$IHSBLKLIC.BackColor = "Green"
$IHSBLKLIC.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$IHSBLKLIC.Add_Click({

IHS-BLKLICTASK
[System.Windows.Forms.MessageBox]::Show("We have Completed the licenses Task", "Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)


})
$IHSBLKLIC.Add_MouseEnter({
    $IHSBLKLIC.BackColor = "DarkRed"
        $helpLabel.Text = "Licenses assign Task in Bulk`n`nPurpose: We can assigned licenes in bulk using template on app folder.`nFeatures :- We can assign & Revoke license in Bulk Format `n`nNote: For assistance, contact Ibrahim (India.coreitsupport@support.com)"

})
$IHSBLKLIC.Add_MouseLeave({
    $IHSBLKLIC.BackColor = "Green"
    $helpLabel.Text = "Hover over any button to see its description here..."
})


$IHSUSBDFWDSTATUS = New-Object System.Windows.Forms.Button
$IHSUSBDFWDSTATUS.Text = "Mail Forwarding Report"
$IHSUSBDFWDSTATUS.AutoSize = $True
$IHSUSBDFWDSTATUS.Location = New-Object System.Drawing.Point(10, 290)
$IHSUSBDFWDSTATUS.ForeColor = "White"
$IHSUSBDFWDSTATUS.BackColor = "Green"
$IHSUSBDFWDSTATUS.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$IHSUSBDFWDSTATUS.Add_Click({
    IHS-USBDFWDSTS
    
 [System.Windows.Forms.MessageBox]::Show("Forwarding report generated at: $outputPath", "Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)


})
$IHSUSBDFWDSTATUS.Add_MouseEnter({
    $IHSUSBDFWDSTATUS.BackColor = "DarkRed"
        $helpLabel.Text = "Mail forwarding report`n`nPurpose: Fill the UPN in template get the report in App Folder based on users.`nFeatures :- Fill the template based On Users`n`nNote: For assistance, contact Ibrahim (India.coreitsupport@support.com)"

})
$IHSUSBDFWDSTATUS.Add_MouseLeave({
    $IHSUSBDFWDSTATUS.BackColor = "Green"
    $helpLabel.Text = "Hover over any button to see its description here..."
})


$IHSUSBDDLGSTATUS = New-Object System.Windows.Forms.Button
$IHSUSBDDLGSTATUS.Text = "Users Delegation Report"
$IHSUSBDDLGSTATUS.AutoSize = $True
$IHSUSBDDLGSTATUS.Location = New-Object System.Drawing.Point(270, 290)
$IHSUSBDDLGSTATUS.ForeColor = "White"
$IHSUSBDDLGSTATUS.BackColor = "Green"
$IHSUSBDDLGSTATUS.Font = New-Object System.Drawing.Font("Cambria",12,[System.Drawing.FontStyle]::Bold)
$IHSUSBDDLGSTATUS.Add_Click({
    IHS-USBDDLGSTS
    
 [System.Windows.Forms.MessageBox]::Show("Forwarding report generated at: $outputPath", "Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)


})
$IHSUSBDDLGSTATUS.Add_MouseEnter({
    $IHSUSBDDLGSTATUS.BackColor = "DarkRed"
            $helpLabel.Text = "Mail Delegation report`n`nPurpose: Fill the UPN in template get the report in App Folder based on users.`nFeatures :- Fill the template based On Users`n`nNote: For assistance, contact Ibrahim (India.coreitsupport@support.com)"

})
$IHSUSBDDLGSTATUS.Add_MouseLeave({
    $IHSUSBDDLGSTATUS.BackColor = "Green"
    $helpLabel.Text = "Hover over any button to see its description here..."
})


$form.Controls.Add($labelIHS)
$form.Controls.Add($IHSBLKLIC)
$form.Controls.Add($IHSM365LC)
$form.Controls.Add($IHSUSBDLICSTATUS)
$form.Controls.Add($M365IHS)
$form.Controls.Add($IHSGRPTSK)
$form.Controls.Add($IHSGRPEXP)
$form.Controls.Add($IHSMDINT)
$form.Controls.Add($script4Button)
$form.Controls.Add($script3Button)
$form.Controls.Add($script2Button)
$form.Controls.Add($GPLFTIHS)
$form.Controls.Add($CMTIHS)
$form.Controls.Add($EmailUPIHS)
$form.Controls.Add($script1Button)
$form.Controls.Add($M365IHSIMID)
$form.Controls.Add($IHSUSBDFWDSTATUS)
$form.Controls.Add($IHSUSBDDLGSTATUS)
$form.Controls.Add($logoutButton)
$form.Controls.Add($helpPanel)
$form.Controls.Add($IHDNDLBT)





$tooltip = New-Object System.Windows.Forms.ToolTip
$tooltip.SetToolTip($labelIHS, "Ibrahim Script this will make task easier.")
$tooltip.SetToolTip($IHSBLKLIC, "Fill the Template first for Licenses assign in Bulk Format.")
$tooltip.SetToolTip($IHSM365LC, " Check the user name & licenses task for single user.")
$tooltip.SetToolTip($IHSUSBDLICSTATUS, "Need UGDN on template for licenses Status & export status in Application Folder.")
$tooltip.SetToolTip($M365IHS, "This Menu for M365 Task")
$tooltip.SetToolTip($IHSGRPTSK, "We can Add & remove group in bulk format using template")
$tooltip.SetToolTip($IHSGRPEXP, "Just need a group name it will export user list in CSV on application folder")
$tooltip.SetToolTip($IHSMDINT, "Module Menu We can install & connect")
$tooltip.SetToolTip($script4Button, "We can export some user detail like AD dumps")
$tooltip.SetToolTip($script3Button, "We can export User based on OU")
$tooltip.SetToolTip($script2Button, "One click Export AD Dump in 20 min")
$tooltip.SetToolTip($GPLFTIHS, "After mapping OLD UGDN group member auto add to New UGDN")
$tooltip.SetToolTip($CMTIHS, "We can reset Comment for ServicesNow access")
$tooltip.SetToolTip($EmailUPIHS, "We can add primary Email ID from UGDN ")
$tooltip.SetToolTip($script1Button, "One Click Password reset")
$tooltip.SetToolTip($M365IHSIMID, "If Need Immutable ID from UGDN After click it will Copy as well")
$tooltip.SetToolTip($IHSUSBDFWDSTATUS, "User based Forwording Report using UPN from Template")
$tooltip.SetToolTip($IHSUSBDDLGSTATUS, "User based Delegation Report using UPN from Template")

$form.ShowDialog() | Out-Null

}

Show-LoginForm
