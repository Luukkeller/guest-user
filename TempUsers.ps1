Import-Module ActiveDirectory

Search-ADAccount -SearchBase 'OU=GuestAccounts,OU=Guests,DC=domain-name,DC=nl' -AccountExpired | 
Remove-ADUser -Confirm:$false

$UserList = Get-ADuser -Filter * -SearchBase 'OU=GuestAccounts,OU=Guests,DC=domain-name,DC=nl'
$UserList = $UserList | Sort-Object
$FirstUser = 1
$NextUser = 1
$Proceed=$true

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object Windows.Forms.Form


$form.Text = 'Select a Date'
$form.Size = New-Object Drawing.Size @(243,230)
$form.StartPosition = 'CenterScreen'

$calendar = New-Object System.Windows.Forms.MonthCalendar
$calendar.ShowTodayCircle = $false
$calendar.MaxSelectionCount = 1
$form.Controls.Add($calendar)

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(38,165)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(113,165)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$form.Topmost = $true

$result = $form.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $date = $calendar.SelectionStart
    Write-Host "Date selected: $($date.ToShortDateString())"
}

foreach ($user in $userlist) {
  $Numberstring = ($user.SAMaccountname).substring(5,2)
  $Numberrstring
  $Number = $Numberstring / 1
  $Number

  if (($Number -eq $NextUser) -and $Proceed) {
     $NextUser ++
     }
  else {
    $Proceed=$false 
    }
  
}

$NextUser
$NextString = $NextUser.ToString("00")
$NextString

$NewUser = "Guest-" + $NextString
$NewUser


function Get-RandomCharacters($length, $characters) {
$random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
$private:ofs=""
return [String]$characters[$random]
}

$password = Get-RandomCharacters -length 3 -characters 'abcdefghijklmnopqrstuvwxyz'
$password += Get-RandomCharacters -length 1 -characters 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
$password += Get-RandomCharacters -length 1 -characters '1234567890'


New-ADUser `
-Name "$NewUser" `
-AccountPassword (ConvertTo-SecureString "$password" -AsPlainText -Force) `
-Path "OU=GuestAccounts,OU=Guests,DC=domain-name,DC=nl" `
-AccountExpirationDate $($date.ToShortDateString()) `
-Enabled 1 

Add-ADGroupMember -Identity ScriptGast -Member "$NewUser"

$Outlook = New-Object -com Outlook.Application

$ProcessActive = Get-Process Outlook -ErrorAction SilentlyContinue
if($ProcessActive -eq $null)
{
 $Outlook = New-Object -com Outlook.Application
}

$mail = $Outlook.CreateItem(0)
$mail.Display()
$mail.subject = “Credentials $NewUser“

$mail.body = “ Username = $NewUser `n Password = $password `n Expirationedate =  $date“

