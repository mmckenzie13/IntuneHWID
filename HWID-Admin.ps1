## For Your Workstation to Build Master CSV ##

<# Create C:\HWIDZIP directory. Drag and Drop E-Mail Attachments from Outlook messages into the folder (fastest method) #> 

## After Downloading All Zips to a Directory, Extract ##
Get-ChildItem 'C:\HWIDZIP' -Filter *.zip | Expand-Archive -DestinationPath 'C:\HWID-Extracted' -Force

## Combine CSVs into One for Intune Import ##
Set-Location C:\HWID-Extracted
Get-ChildItem -Filter *.csv | Select-Object -ExpandProperty FullName | Import-Csv | Export-Csv .\FullAutoPilotHWIDList.csv -NoTypeInformation -Append

<# Upload FullAutoPilotHWIDList.csv to https://devicemanagement.portal.azure.com/#blade/Microsoft_Intune_Enrollment/AutoPilotDevicesBlade #>

## End of Your Workstation Upload ##
