## Beginning of Workstation Script ##
## Run Powershell as Admin ##

## Collect Serial Number ##
$Serial = (get-WMIObject win32_bios).serialnumber

## Create Working Directory, Changed Execution Policy, Download Script ##
md c:\\HWID
Set-Location c:\\HWID
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted
Install-Script -Name Get-WindowsAutoPilotInfo
Set-ExecutionPolicy RemoteSigned
Get-WindowsAutoPilotInfo.ps1 -OutputFile ($Serial + ' AutoPilotHWID.csv')

## Compressing the Working Directory ##
Compress-Archive C:\HWID\* -DestinationPath ('c:\HWID ' + $Serial + '.zip')

## O365 SMTP, Requires Creds to E-mail ##
$creds = get-credential
$subject = 'HWID ' + $Serial
$attachment = 'C:\HWID ' + $serial + '.zip'
Send-MailMessage -Credential $creds -SmtpServer smtp.office365.com -Port 587 -usessl -From source@customer.com -To destination@customer.com -Subject $subject -Body HWID -Attachments $attachment  -Priority High -DeliveryNotificationOption OnSuccess, OnFailure
## End of Workstation Script ##
