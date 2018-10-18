# Author      : Marcus Dempsey, TeraByte
# Description : This script is used for sending emails via Office 365 to a company who is undergoing the Cyber Essentials Plus assessment.
# Version     : 1.0
# Last Updated: 17/10/2018

$smtpCredentials = (Get-Credential)
$ToAddress       = "cyberessentialsplus@terabyteit.co.uk" # This is a distribution group which has rules in place to bypass checking of attachments and contains the person being emailed.
$FromAddress     = "marcus@terabyteit.co.uk" # Fom email address
$SMTPServer      = "smtp.office365.com"
$SMTPPort        = '587'
$CompanyName     = "TeraByte"
$EmailFooter     = "email: info@terabyteit.co.uk`ntel  : 01325 628587"

$Files = Get-Childitem 'C:\git\CyberEssentialsPlus\files' -File -Recurse -ErrorAction SilentlyContinue | Select-Object Name, Directory # Change this location if needed

$count = 1
foreach ($File in $Files) {        
    $FileAttachment = (Resolve-Path ($File.Directory)).Path + "\" + $File.Name
    $Subject = $CompanyName + ": Cyber Essentials PLUS email attachment test " + $count + " of " + $Files.Count
    $Body = "Hello,`r
This is a test email for the Cyber Essentials Plus audit which is being provided by TeraByte.`r
The attachment (and any others received) are used to validate that the defences of your system are in place and are working in accordance with the scheme.`r
Do not open the attachment, as this needs to be recorded.`n
Regards,`n`n
$CompanyName`n
$EmailFooter"

    Write-Host "Attempting to send attachment: " + $FileAttachment -ForegroundColor Green
    Send-MailMessage -To $ToAddress -From $FromAddress -Subject $Subject -Body $Body -SmtpServer $SMTPServer -Port $SMTPPort -Credential $smtpCredentials -Attachments $FileAttachment -UseSsl
    $count += 1
}
