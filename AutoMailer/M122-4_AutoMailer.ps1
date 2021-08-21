[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned

$LoginGUI = New-Object System.Windows.Forms.Form
$MainGUI = New-Object System.Windows.Forms.Form

Clear-host
$programLocation = "$env:userprofile\Documents\AutoMailer"
Set-Location $programLocation

#  Login GUI
# ----------

# Cosmetics

$LoginGUI.Text = "Login - Auto E-Mailer"
# $LoginGUI.Icon = "AutoMailer.ico" # ERROR: Icon doesnt work
$LoginGUI.Width = 500
$LoginGUI.Height = 300

# Function

$func_Login = {

    if ($txt_CredEMail.text -match "^[a-z0-9](\.?[a-z0-9]){5,}@g(oogle)?mail\.com$") {
        $cred_email = $txt_CredEMail.text
        $cred_password = $txt_CredPassword.text
    
        $LoginGUI.Close()
        $MainGUI.ShowDialog()
    }
    else {
        $LoginGUI.Controls.Add($lbl_LoginError)
    }
}

# Label for Login Error

$lbl_LoginError = New-Object System.Windows.Forms.Label
$lbl_LoginError.top = 55
$lbl_LoginError.left = 60
$lbl_LoginError.Width = 400
$lbl_LoginError.Height = 30
$lbl_LoginError.Text = "Please enter a valid Gmail address"
$lbl_LoginError.Font = New-Object System.Drawing.Font("Arial", 10)
$lbl_LoginError.ForeColor = "Red"

# Title for Login

$lbl_Login = New-Object System.Windows.Forms.Label
$LoginGUI.Controls.Add($lbl_Login)
$lbl_Login.top = 25
$lbl_Login.left = 220
$lbl_Login.Width = 100
$lbl_Login.Height = 30
$lbl_Login.Text = "Login"
$lbl_Login.Font = New-Object System.Drawing.Font("Arial", 12)

# Textbox for E-Mail

$txt_CredEMail = New-Object System.Windows.Forms.TextBox
$LoginGUI.Controls.Add($txt_CredEMail)
$txt_CredEMail.Left = 40
$txt_CredEMail.top = 100
$txt_CredEMail.Width = 400
$txt_CredEMail.Height = 50
$txt_CredEMail.text = "E-Mail (example@gmail.com)"
$txt_CredEMail.Font = New-Object System.Drawing.Font("Arial", 12)

# Textbox for password

$txt_CredPassword = New-Object System.Windows.Forms.TextBox
$LoginGUI.Controls.Add($txt_CredPassword)
$txt_CredPassword.Left = 40
$txt_CredPassword.top = 140
$txt_CredPassword.Width = 400
$txt_CredPassword.Height = 50
$txt_CredPassword.text = "Password"
$txt_CredPassword.Font = New-Object System.Drawing.Font("Arial", 12)

# Button to log in

$btn_Login = New-Object System.Windows.Forms.Button
$LoginGUI.Controls.Add($btn_Login)
$btn_Login.Text = "Login"
$btn_Login.Left = 200
$btn_Login.Top = 180
$btn_Login.Width = 100
$btn_Login.Height = 30
$btn_Login.Enabled = 1
$btn_Login.Add_Click($func_Login)

#  Main GUI
# ----------

# Cosmetics

$MainGUI.Text = "Auto E-Mailer"
# $MainGUI.Icon = "AutoMailer.ico" # ERROR: Icon doesnt work
$MainGUI.Width = 1000
$MainGUI.Height = 650

# Button Functions

$func_ShowSavedMails = {
    # Assembly für Forms laden            
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null            
                
    # Ordner auswählen            
    $OrdnerWahl = New-Object System.Windows.Forms.FolderBrowserDialog            
                
    # Neuer Ordner anlegen ausschalten            
    $OrdnerWahl.ShowNewFolderButton = $false            
                
    # Dialog anzeigen            
    $OrdnerWahl.ShowDialog()            
                
    # Falls ein Ordner gewählt wurde (nicht Abbrechen)            
    if ($OrdnerWahl.SelectedPath -ne "")            
        {            
        
        # Neue Datei erstellen            
        New-Item -Path $OrdnerWahl.SelectedPath -Name "TestDatei.txt" -ItemType File -Force | Out-Null            
                
        # Objekt für Datei Auswahl erstellen            
        $DateiWahl = New-Object System.Windows.Forms.OpenFileDialog            
                    
        # Start Ordner festlegen            
        $DateiWahl.InitialDirectory = $OrdnerWahl.SelectedPath            
                    
        # Filter mit Dateiendungen erstellen            
        $DateiWahl.Filter = "Textdateien (*.txt) | *.txt"            
                    
        # Auswahl von mehreren Dateien ausschalten            
        $DateiWahl.Multiselect = $false            
                    
        # Ordner Auswahl Dialog anzeigen            
        $DateiWahl.ShowDialog() | Out-Null            
                
        # Prüfen ob eine Datei ausgewaehlt wurde            
        if ($DateiWahl.FileName -ne "")            
            {            
                            
                # Ausgabe der gewaehlten Datei            
                Write-Host "Gewaehlte Datei ist $($DateiWahl.FileName)"
                
                $txt_MailText.Text = Get-Content $($DateiWahl.FileName)
                $MainGUI.Refresh()
            }
        }
}

$func_SaveMailTxt = {

    $saveMailPath = "$programLocation\SavedMails\" + $txt_subject.Text + ".txt"
    $txt_MailText.Text | Out-File $saveMailPath
    
    Write-Host "Log: $txt_MailText.Text > $saveMailPath"
}

$func_OpenSettings = {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

    $SettingsGUI = New-Object System.Windows.Forms.Form 

    $SettingsGUI.Text = "Settings - Auto E-Mailer"
    $SettingsGUI.Icon = "C:\Users\Sandro Lenz\Downloads\Martz90-Circle-Windows-8.ico"

    # Disabled
}

$func_Send = {

    try {
        $From = $txt_SenderName.text + " <" + $txt_CredEmail.Text + ">"
        $To = $txt_Receiver.text
        $Subject = $txt_Subject.text
        $Body = $txt_MailText.text
        $SMTPServer = "smtp.gmail.com"
        $SMTPPort = "587"
        $Credential = New-Object PSCredential($txt_CredEmail.Text, (ConvertTo-SecureString $txt_CredPassword.Text -AsPlainText -Force))
        $i = 0

        Write-Host "-From $From -To $To -Subject $Subject -Body $Body -Credential ($Credential)"
        do{
            Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -BodyAsHTML -SMTPServer $SMTPServer -Port $SMTPPort -UseSSL -Credential ($Credential)
            $i = $i + 1
            Write-Host "Log: Amount Sent $i"
            Start-Sleep -Seconds $txt_Interval.Text
        }
        while($i -lt $txt_count.Text)
    } 
    catch 
    {
        $MainGUI.Controls.Add($lbl_SendError)
        Write-Host "Log: Error while sending"
    } 
    
}

# Label for Sending Error

$lbl_SendError = New-Object System.Windows.Forms.Label
$lbl_SendError.top = 470
$lbl_SendError.left = 60
$lbl_SendError.Width = 400
$lbl_SendError.Height = 30
$lbl_SendError.Text = "Something went wrong, please check the fields"
$lbl_SendError.Font = New-Object System.Drawing.Font("Arial", 10)
$lbl_SendError.ForeColor = "Red"

# Button to see saved mails

$btn_SavedMails = New-Object System.Windows.Forms.Button
$MainGUI.Controls.Add($btn_SavedMails)
$btn_SavedMails.Text = "Saved E-Mails"
$btn_SavedMails.Left = 510
$btn_SavedMails.Top = 20
$btn_SavedMails.Width = 150
$btn_SavedMails.Height = 30
$btn_SavedMails.Add_Click($func_ShowSavedMails)

# Button to save current mail text

$btn_SaveMailTxt = New-Object System.Windows.Forms.Button
$MainGUI.Controls.Add($btn_SaveMailTxt)
$btn_SaveMailTxt.Text = "Save Text"
$btn_SaveMailTxt.Left = 680
$btn_SaveMailTxt.Top = 20
$btn_SaveMailTxt.Width = 100
$btn_SaveMailTxt.Height = 30
$btn_SaveMailTxt.Enabled = 1
$btn_SaveMailTxt.Add_Click($func_SaveMailTxt)

# Button to open settings

$btn_Settings = New-Object System.Windows.Forms.Button
$MainGUI.Controls.Add($btn_Settings)
$btn_Settings.Text = "Settings"
$btn_Settings.Left = 800
$btn_Settings.Top = 20
$btn_Settings.Width = 100
$btn_Settings.Height = 30
$btn_Settings.Enabled = 0
$btn_Settings.Add_Click($func_OpenSettings)

# Textbox for Mail-Text

$txt_MailText = New-Object System.Windows.Forms.TextBox
$MainGUI.Controls.Add($txt_MailText)
$txt_MailText.Multiline = $True
$txt_MailText.Scrollbars = "Vertical"
$txt_MailText.Left = 510
$txt_MailText.top = 70
$txt_MailText.Width = 390
$txt_MailText.Height = 500
$txt_MailText.text = "Please enter E-Mail-Text in HTML"
$txt_MailText.Font = New-Object System.Drawing.Font("Arial", 12)

# Title for "Infos" section

$lbl_Infos = New-Object System.Windows.Forms.Label
$MainGUI.Controls.Add($lbl_Infos)
$lbl_Infos.top = 25
$lbl_Infos.left = 75
$lbl_Infos.Width = 100
$lbl_Infos.Height = 30
$lbl_Infos.Text = "Infos:"
$lbl_Infos.Font = New-Object System.Drawing.Font("Arial", 12)

# Textbox for sender name

$txt_SenderName = New-Object System.Windows.Forms.TextBox
$MainGUI.Controls.Add($txt_SenderName)
$txt_SenderName.Left = 75
$txt_SenderName.top = 70
$txt_SenderName.Width = 400
$txt_SenderName.Height = 50
$txt_SenderName.text = "Name of Sender"
$txt_SenderName.Font = New-Object System.Drawing.Font("Arial", 12)

# Textbox for receiver

$txt_Receiver = New-Object System.Windows.Forms.TextBox
$MainGUI.Controls.Add($txt_Receiver)
$txt_Receiver.Left = 75
$txt_Receiver.top = 110
$txt_Receiver.Width = 400
$txt_Receiver.Height = 50
$txt_Receiver.text = "E-Mail of Receiver"
$txt_Receiver.Font = New-Object System.Drawing.Font("Arial", 12)

# Textbox for subject

$txt_Subject = New-Object System.Windows.Forms.TextBox
$MainGUI.Controls.Add($txt_Subject)
$txt_Subject.Left = 75
$txt_Subject.top = 150
$txt_Subject.Width = 400
$txt_Subject.Height = 50
$txt_Subject.text = "Subject"
$txt_Subject.Font = New-Object System.Drawing.Font("Arial", 12)

# Textbox for attachement

$txt_Attachement = New-Object System.Windows.Forms.TextBox
$MainGUI.Controls.Add($txt_Attachement)
$txt_Attachement.Left = 75
$txt_Attachement.top = 190
$txt_Attachement.Width = 275
$txt_Attachement.Height = 50
$txt_Attachement.text = "Attachement"
$txt_Attachement.Enabled = 0
$txt_Attachement.Font = New-Object System.Drawing.Font("Arial", 12)

# Button to select attachement

$btn_Attachement = New-Object System.Windows.Forms.Button
$MainGUI.Controls.Add($btn_Attachement)
$btn_Attachement.Left = 360
$btn_Attachement.top = 190
$btn_Attachement.Width = 100
$btn_Attachement.Height = 25
$btn_Attachement.text = "Browse"
$btn_Attachement.Enabled = 0

# Line

$lbl_Line = New-Object System.Windows.Forms.Label
$MainGUI.Controls.Add($lbl_Line)
$lbl_Line.top = 230
$lbl_Line.left = 75
$lbl_Line.Width = 400
$lbl_Line.Height = 30
$lbl_Line.Text = "_______________________________________________________________"
$lbl_Line.Font = New-Object System.Drawing.Font("Arial", 12)

# Title for "Attributes" section

$lbl_Attributes = New-Object System.Windows.Forms.Label
$MainGUI.Controls.Add($lbl_Attributes)
$lbl_Attributes.top = 270
$lbl_Attributes.left = 75
$lbl_Attributes.Width = 400
$lbl_Attributes.Height = 30
$lbl_Attributes.Text = "Attributes:"
$lbl_Attributes.Font = New-Object System.Drawing.Font("Arial", 12)

# Label for count

$lbl_count = New-Object System.Windows.Forms.Label
$MainGUI.Controls.Add($lbl_count)
$lbl_count.Left = 75
$lbl_count.top = 310
$lbl_count.Width = 275
$lbl_count.Height = 30
$lbl_count.text = "Count:"
$lbl_count.Font = New-Object System.Drawing.Font("Arial", 10)

# Textbox for count

$txt_count = New-Object System.Windows.Forms.TextBox
$MainGUI.Controls.Add($txt_count)
$txt_count.Left = 360
$txt_count.top = 310
$txt_count.Width = 100
$txt_count.Height = 25
$txt_count.text = ""

# Label for interval

$lbl_interval = New-Object System.Windows.Forms.Label
$MainGUI.Controls.Add($lbl_interval)
$lbl_interval.Left = 75
$lbl_interval.top = 350
$lbl_interval.Width = 275
$lbl_interval.Height = 30
$lbl_interval.text = "Interval in seconds (e.g. 60):"
$lbl_interval.Font = New-Object System.Drawing.Font("Arial", 10)

# Textbox for interval

$txt_interval = New-Object System.Windows.Forms.TextBox
$MainGUI.Controls.Add($txt_interval)
$txt_interval.Left = 360
$txt_interval.top = 350
$txt_interval.Width = 100
$txt_interval.Height = 25
$txt_interval.text = ""

# Button to send

$btn_Send = New-Object System.Windows.Forms.Button
$MainGUI.Controls.Add($btn_Send)
$btn_Send.Text = "Send"
$btn_Send.Left = 75
$btn_Send.Top = 485
$btn_Send.Width = 400
$btn_Send.Height = 75
$btn_Send.Enabled = 1 # Disable if not all fields are filled (ex. Attachement)
$btn_Send.Add_Click($func_Send)

# Show Window

$LogInGUI.ShowDialog()
