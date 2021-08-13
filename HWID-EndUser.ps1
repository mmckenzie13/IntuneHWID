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
## Try Using Authentication method - Option 1 - https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-microsoft-365-or-office-365 ##
try {
$creds = get-credential
$subject = 'HWID ' + $Serial
$attachment = 'C:\HWID ' + $serial + '.zip'
Send-MailMessage -Credential $creds -SmtpServer smtp.office365.com -Port 587 -usessl -From source@customer.com -To destination@customer.com -Subject $subject -Body HWID -Attachments $attachment  -Priority High -DeliveryNotificationOption OnSuccess, OnFailure
}
## Else, Try unauthenticated - Option 2 - https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-microsoft-365-or-office-365 ##
catch [System.Net.WebException],[System.IO.IOException] {
$subject = 'HWID ' + $Serial
$attachment = 'C:\HWID ' + $serial + '.zip'
Send-MailMessage -SmtpServer customer-com.mail.protection.outlook.com -Port 25 -usessl -From source@customer.com -To destination@customer.com -Subject $subject -Body HWID -Attachments $attachment  -Priority High -DeliveryNotificationOption OnSuccess, OnFailure
} ## Else, Open Up New Outlook E-mail Message with attachment with existing profile ##
catch {
$Outlook = New-Object -ComObject Outlook.Application
$Mail = $Outlook.CreateItem(0)
$Mail.To = "destination@customer.com"
$Mail.Subject = "HWID Hash"
$Mail.Body ="HWID Hash"
$Mail.Attachments.Add($attachment)
##$Mail.Display() - To Test if needed, uncomment ##
$Mail.Send()
}
## End of Workstation Script ##




