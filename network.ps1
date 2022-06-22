#set-executionpolicy remotesigned
import-module activedirectory
Add-Type -assembly System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Description of the main form
$Form = New-Object System.Windows.Forms.Form
    $Form.Text = 'Network'
    $Form.Width = 435
    $Form.Height = 365
    $Form.StartPosition = 'CenterScreen'
    $Form.BackColor = '#272537'
    $Form.ForeColor = 'Black'
    $Form.FormBorderStyle = "Fixed3D"
    $Form.AllowTransparency = 75
    $Form.ShowIcon = $false
    $Form.ShowInTaskbar = $true
    $Form.MaximizeBox = $false
    $Form.MinimizeBox = $false

# Label "PC:"
$LabelPC = New-Object System.Windows.Forms.Label
    $LabelPC.Text = "PC:"
    $LabelPC.Location = New-Object System.Drawing.Point(18, 6)
    $LabelPC.Autosize = $true
    $LabelPC.Font = 'Calibri, 12'
    $LabelPC.ForeColor = 'White'
$Form.Controls.Add($LabelPC)

# Label "URL:"
$LabelURL = New-Object System.Windows.Forms.Label
    $LabelURL.Text = "URL:"
    $LabelURL.Location = New-Object System.Drawing.Point(8, 36)
    $LabelURL.Autosize = $true
    $LabelURL.Font = 'Calibri, 12'
	$LabelURL.ForeColor = 'White'
$Form.Controls.Add($LabelURL)

# Input field PC
$TextBoxPC = New-Object System.Windows.Forms.Textbox
    $TextBoxPC.Location = New-Object System.Drawing.Point(45, 5)
    $TextBoxPC.Size = New-Object System.Drawing.Size(100, 10)
    $TextBoxPC.BackColor = 'Black'
    $TextBoxPC.ForeColor = 'White'
    $TextBoxPC.Font = 'Calibri, 10'
$Form.Controls.Add($TextBoxPC)

# Input field URL
$TextBoxURL = New-Object System.Windows.Forms.Textbox
    $TextBoxURL.Location = New-Object System.Drawing.Point(45,35)
    $TextBoxURL.MinimumSize = New-Object System.Drawing.Size(365,25)
    $TextBoxURL.BackColor = 'Black'
    $TextBoxURL.ForeColor = 'White'
    $TextBoxURL.Font = 'Calibri, 10'
$Form.Controls.Add($TextBoxURL)

# Output field (ping)
$MessageTextBox = New-Object System.Windows.Forms.Textbox
    $MessageTextBox.Location = New-Object System.Drawing.Point(10,70)
    $MessageTextBox.MinimumSize = New-Object System.Drawing.Size(400,250)
    $MessageTextBox.Multiline = $true
    $MessageTextBox.Text = $null
    $MessageTextBox.BackColor = 'Black'
    $MessageTextBox.ForeColor = 'White'
    $MessageTextBox.Font = 'Calibri, 10'
$Form.Controls.Add($MessageTextBox)
			
# Button "Ping"
$Button_Find2 = New-Object System.Windows.Forms.Button
    $Button_Find2.Location = New-Object System.Drawing.Point(150, 5)
    $Button_Find2.Size = New-Object System.Drawing.Size(50,25)
    $Button_Find2.Text = 'Ping'
    $Button_Find2.Font = 'Calibri, 12'
    $Button_Find2.BackColor = 'Black'
    $Button_Find2.ForeColor = 'White'
$Form.Controls.Add($Button_Find2)

# Button "PAC"
$pacb = New-Object System.Windows.Forms.Button
    $pacb.Location = New-Object System.Drawing.Point(205, 5)
    $pacb.Size = New-Object System.Drawing.Size(50, 25)
    $pacb.Text = 'PAC'
    $pacb.Font = 'Calibri, 12'
    $pacb.ForeColor = 'White'
    $pacb.BackColor = 'Black'
$Form.Controls.Add($pacb)

# Button "Diag"
$Diag = New-Object System.Windows.Forms.Button
    $Diag.Location = New-Object System.Drawing.Point(260, 5)
    $Diag.Size = New-Object System.Drawing.Size(50, 25)
    $Diag.Text = 'Diag'
    $Diag.Font = 'Calibri, 12'
    $Diag.ForeColor = 'White'
    $Diag.BackColor = 'Black'
$Form.Controls.Add($Diag)

# Button "Download"
$Download = New-Object System.Windows.Forms.Button
    $Download.Location = New-Object System.Drawing.Point(315, 5)
    $Download.Size = New-Object System.Drawing.Size(95, 25)
    $Download.Text = 'Download'
    $Download.Font = 'Calibri, 12'
    $Download.ForeColor = 'White'
    $Download.BackColor = 'Black'
$Form.Controls.Add($Download)

$host1 = hostname

$null = ""

#### Actions
# Button ping
try {

$Button_Find2.Add_Click({

       $comp = $null
       $comp = $TextBoxPC.Text
       $url = $null
       $url = $TextBoxURL.Text
            
            # ping
            $pi = ping $comp | out-string


            $MessageTextBox.Text  = $null
            $MessageTextBox.Text  = $pi
			    $MessageTextBox.BackColor = 'Black'
			    $MessageTextBox.ForeColor = 'White'
			    $MessageTextBox.Font = 'Calibri, 10'
})

} catch {
  $MessageTextBox.Text = "An error occurred:"
  $MessageTextBox.Text  += $_
}

# Button download proxy settings file (pac)
try {

$pacb.Add_Click({
       $error.clear()
       $comp = $null
       $comp = $TextBoxPC.Text
           
           # Enable WinRM
           psexec \\$comp -s powershell Enable-PSRemoting -Force 
           psexec \\$comp -s powershell Start-Service WinRM

              # Logged on user name
              $user = Get-WmiObject win32_computersystem -Property UserName -ComputerName $comp
              $users = $user.username
              $users = $users.Split("\")[1]

                 # User SID
                 $SID = Get-ADUser -Identity $users | select sid
                 $SID = ($SID.sid).value

       Invoke-Command -computername $comp -ScriptBlock {

             cd D:\
             del pac.txt

                 # Get AutoConfigURL
                 $proxy = Get-ItemProperty -Path "registry::HKEY_Users\$using:SID\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings" -Name AutoConfigURL
                 $proxy_conf = $proxy.AutoConfigURL 
                 $proxy_conf | out-string

                 # The file will be downloaded to the D:\ drive of the entered computer
                 Invoke-WebRequest -Uri $proxy_conf -outfile "D:\pac.txt"
}
                     $MessageTextBox.ForeColor = "LightGreen"
                     $MessageTextBox.Text  = "PAC file have created"
                     $MessageTextBox.Font = 'Calibri, 12'
})

} catch {
  $MessageTextBox.Text = "An error occurred:"

}

# Network Diagnostics button
try {
$Diag.Add_Click({
       $error.clear()
       $comp = $null
       $comp = $TextBoxPC.Text
       $url = $null
       $url = $TextBoxURL.Text

           # Enable WinRM
           psexec \\$comp -s powershell Enable-PSRemoting -Force 
           psexec \\$comp -s powershell Start-Service WinRM

           Invoke-Command -computername $comp -ScriptBlock {

              # Logged on user name
              $pc = hostname
              $user = Get-WmiObject win32_computersystem -Property UserName -ComputerName $pc
              $users = $user.username
              $users = $users.Split("\")[1]

              $path = "D:\"
              $url = $TextBoxURL.Text

                 cd D:\
                 del Diagnosis.docx

$pc >> $path
echo *********************** >> $path
$users >> $path
echo *********************** >> $path
ipconfig /all >> $path
echo *********************** >> $path
netstat -rn >> $path 
echo *********************** >> $path
ping  -w 1 -n 5 $using:url >> $path
echo *********************** >> $path
ping  -w 1 -n 5 whatismyip.com >> $path  
echo *********************** >> $path
echo *********************** >> $path
ping  -w 1 -n 5 www.google.com >> $path
echo *********************** >> $path
tracert -w 1 $using:url >> $path
echo *********************** >> $path
tracert -w 1 google.com >> $path
echo *********************** >> $path
get-hotfix >> $path

}

         $MessageTextBox.ForeColor = "LightGreen"
         $MessageTextBox.Text  = "Diagnosis is complete"
         $MessageTextBox.Font = 'Calibri, 12'
})

} catch {
  $MessageTextBox.Text = "An error occurred:"
  $MessageTextBox.Text  += $_
}

# Button Download information from the user's computer to your computer

try {

$Download.Add_Click({
       $error.clear()
       $comp = $null
       $comp = $TextBoxPC.Text

###
### $path2 = path to diagnostic files on your computer
###
           $path2 = "D:\"

           $session = New-PSSession –ComputerName $comp

           Copy-Item –Path "D:\pac.txt" –Destination $path2 –FromSession $session
           Copy-Item –Path "D:\Diagnosis.docx" –Destination $path2 –FromSession $session

              Disable-PSRemoting -Force
              psservice \\$comp Stop WinRM 

                    $MessageTextBox.ForeColor = "LightGreen"
                    $MessageTextBox.Text  = "Downloading is complete"
                    $MessageTextBox.Font = 'Calibri, 12'
}) 

} catch {
  $MessageTextBox.Text = "An error occurred:"
  $MessageTextBox.Text  += $_
}


$Form.Controls.Add($MessageTextBox.Text)

# Form display
$Form.ShowDialog()