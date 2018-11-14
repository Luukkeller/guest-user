Import-Module ActiveDirectory


Search-ADAccount -SearchBase 'OU=Script,OU=GuestAccounts,OU=TemporaryUsers,DC=Example,DC=nl' -AccountExpired | 
Remove-ADUser -Confirm:$false
Remove-Item -Path D:\Scripts\GuestAccounts\ScriptGuest.txt -Confirm:$false


Write-Host (Get-ADUser -Filter * -Searchbase "OU=Script,OU=GuestAccounts,OU=TemporaryUsers,DC=Example,DC=nl" | Select-Object Name)


$username = Read-Host -Prompt "(Hierboven staan de actieve gebruikers) Vul gebruikersnaam in"
$ExpirationDate = Read-Host -Prompt "Voer Expiration Date in zoals dit 10/18/2018 [Day-Month-Year]"


function Get-RandomCharacters($length, $characters) {
$random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
$private:ofs=""
return [String]$characters[$random]
}


$password = Get-RandomCharacters -length 3 -characters 'abcdefghijklmnopqrstuvwxyz'
$password += Get-RandomCharacters -length 1 -characters 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
$password += Get-RandomCharacters -length 1 -characters '1234567890'


New-ADUser `
-Name "$username" `
-AccountPassword (ConvertTo-SecureString "$password" -AsPlainText -Force) `
-Path "OU=Script,OU=GuestAccounts,OU=TemporaryUsers,DC=Example,DC=nl" `
-AccountExpirationDate "$ExpirationDate" `
-Enabled 1 


Add-ADGroupMember -Identity ScriptGast -Member "$username"


if (!(Test-Path "D:\Scripts\GuestAccounts\ScriptGuest.txt"))
{
   New-Item -path D:\Scripts\GuestAccounts -name ScriptGuest.txt -type "file" -value "$username : $Password : $ExpirationDate"
   Write-Host "Created new file and text content added"
}


notepad.exe D:\Scripts\GuestAccounts\ScriptGuest.txt