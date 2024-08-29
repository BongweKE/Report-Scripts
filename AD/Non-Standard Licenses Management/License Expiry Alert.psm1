#Sending Alert

function Log-Message([String]$Message)
{
    $logFilePath = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Account Expiry Alert\LicenseAlertLogs.txt"
    Add-Content -Path $logFilePath $Message
}

function Send-LicenseExpiryAlert
{
    param($accounts_to_alert)

    #define O365 credentials for sending the email alerts.
    $smtpServer = 'SMTP.Office365.com'
    $alertMailUserName = 'servicedesk@cifor-icraf.org'  #change to 'servicedesk@cifor-icraf.org' after auth issue is resolved
    $alertMailPassword = ConvertTo-SecureString -String 'Green32!' -AsPlainText -Force #Change to secure mode credential after testing. Service desk password is Green@32! 'Winter2023'
    $alertCredentials = New-Object System.Management.Automation.PSCredential($alertMailUserName,$alertMailPassword)
    $log_message = "Expiry Alert on Date: "+(Get-Date).ToString("dd-MM-yyyy")+"`n"
    #Send Alert to Users whose accounts are about to expire
    foreach ($users in $accounts_to_alert)
    {
        $firstName= $users."First Name"
        $lastName= $users."Last Name"
        $userEmail = $users."Email Address"
        $license = $users.License
        $Start= $users.Start
        $End = $users.End
        $log_message += "User:"+$firstName+"_"+$lastName+" User_Email_Address:"+$userEmail+"  License:"+$license+" Start Data:"+$Start+" End Date:"+$End+"`n"
        $emailSubject = $license + "License Expiration Notice"
    $alert = @"
Dear $firstName</br></br>

<p>Please note your $license license that was assigned to you on $Start has expired. Incase of any enquiry please contact servicedesk@cifor-icraf.org.</p>
<p>ICT Service Desk</p>
<p>Ext. 4500 / +254 711 0 4500</p>
<p>servicedesk@cifor-icraf.org</p>
"@
    #Changed the 'to' parameter to the recipient variable when ready.
    Send-MailMessage -to $userEmail -From $alertMailUserName -Subject $emailSubject -Body $alert -BodyAsHtml -SmtpServer $smtpServer -Port 587 -UseSsl -Credential $alertCredentials
    Start-Sleep -s 60
    }
    Log-Message $log_message
}

function Send-LicenseExpiryWarning
{
    param($accounts_to_alert)

    #define O365 credentials for sending the email alerts.
    $smtpServer = 'SMTP.Office365.com'
    $alertMailUserName = 'servicedesk@cifor-icraf.org'
    $alertMailPassword = ConvertTo-SecureString -String 'Green32!' -AsPlainText -Force
    $alertCredentials = New-Object System.Management.Automation.PSCredential($alertMailUserName,$alertMailPassword)
    $log_message = "Expiry Warning on Date: "+(Get-Date).ToString("dd-MM-yyyy")+"`n"
    #Send Alert to Users whose accounts are about to expire
    foreach ($users in $accounts_to_alert)
    {
        $firstName= $users."First Name"
        $lastName= $users."Last Name"
        $userEmail = $users."Email Address"
        $license = $users.License
        $Start= $users.Start
        $End = $users.End
        $Expiry = $users.Expiry
        $log_message += "User:"+$firstName+"_"+$lastName+" User_Email_Address:"+$userEmail+"  License:"+$license+" Start Data:"+$Start+" End Date:"+$End+"`n"
        $emailSubject = $license + "License Expiration Warning"
    $alert = @"
Dear $firstName</br></br>

<p>Please note your $license license that was assigned to you on $Start will expiry in $Expiry. If you need an extension, kindly fill in the form linked below.</p>
<a href='https://ecv.microsoft.com/A3qIJOX8bA'>License Application Form</a>
<br>
<p>Incase of any enquiry please contact servicedesk@cifor-icraf.org.</p>
<p>ICT Service Desk</p>
<p>Ext. 4500 / +254 711 0 4500</p>
<p>servicedesk@cifor-icraf.org</p>
"@
    #Changed the 'to' parameter to the recipient variable when ready.
    Send-MailMessage -to $userEmail -From $alertMailUserName -Subject $emailSubject -Body $alert -BodyAsHtml -SmtpServer $smtpServer -Port 587 -UseSsl -Credential $alertCredentials
    Start-Sleep -s 60
    }
    Log-Message $log_message
}

Export-ModuleMember -Function *