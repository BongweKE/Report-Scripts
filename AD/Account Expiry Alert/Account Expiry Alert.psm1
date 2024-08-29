#Sending Alert

function Log-Message([String]$Message)
{
    $logFilePath = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Account Expiry Alert\AccountExpiryAlertLogs.txt"
    Add-Content -Path $logFilePath $Message
}

function Send-AccountExpiryAlert
{
    param($accounts_to_alert)

    #define O365 credentials for sending the email alerts.
    $smtpServer = 'SMTP.Office365.com'
    $alertMailUserName = 'servicedesk@cifor-icraf.org'
    $alertMailPassword = ConvertTo-SecureString -String 'Green32!' -AsPlainText -Force #Change to secure mode credential after testing. Service desk password is Green@32! 'Winter2023'
    $alertCredentials = New-Object System.Management.Automation.PSCredential($alertMailUserName,$alertMailPassword)
    $log_message = "Date: "+(Get-Date).ToString("dd-MM-yyyy")+"`n"
    #Send Alert to Users whose accounts are about to expire
    $log_message += "Beginning exeuction of the script...`n"
    foreach ($users in $accounts_to_alert)
    {
        $FirstName = $users.GivenName
        $AccountExpiryDate = $users.AccountExpirationDate
        $Surname = $users.Surname
        $AccountExpiry = $users.AccountExpiry
        $userEmail = $users.EmailAddress
        $SupervisorEmail = $users.SupervisorEmailAddress
        $HRFocalPointEmail = $users.HrFocalPoint
        $emails_to_cc = @($SupervisorEmail, $HRFocalPointEmail)
        $log_message += "User:"+$FirstName+"_"+$Surname+" User_Email_Address:"+$userEmail+"  Account_Expiry_Date:"+$AccountExpiryDate+" Supervisor_Email_Address:"+$SupervisorEmail+" HR_Focal_Point_Email_address:"+$HRFocalPointEmail+"`n"
    $alert = @"
Dear $FirstName</br></br>

<p>Please note your AD/Email account, will expire $AccountExpiry from today, i.e. on $AccountExpiryDate. This is in line with your contract expiry date.</p>
<p>Please reach out to your HR focal point on cc to alert IT  by responding to this email (or sending an email to servicedesk@cifor-icraf.org) on your new contract dates. Upon expiry, you will not have access to email and other institutional  systems.</p>
<p>ICT Service Desk</p>
<p>Ext. 4500 / +254 711 0 4500</p>
<p>servicedesk@cifor-icraf.org</p>
"@
    #Changed the 'to' parameter to the recipient variable when ready.
    Send-MailMessage -to $userEmail -From $alertMailUserName -Cc $emails_to_cc -Subject '[ACTION REQUIRED] Account Expiration Notice' -Body $alert -BodyAsHtml -SmtpServer $smtpServer -Port 587 -UseSsl -Credential $alertCredentials
    Start-Sleep -s 60
    }
    $log_message += "Completed exeuction of the script`n`n"
    Log-Message $log_message
}
Export-ModuleMember -Function *