function Send-Alert
{
    param ($recipients,$reportDirectory,$inactiveUsers, $daysToCleanUp)
    $reportDate = Get-Date -Format 'dd MMMM yyyy'
    $smtpServer = 'SMTP.Office365.com'
    $alertMailUserName = 'CIFORICRAFAutoReport@cifor-icraf.org'
    $subject = 'Active Directory Alert: Inactive Accounts (ICRAF) - ' + $reportDate
    $alertMailUserName = 'CIFORICRAFAutoreport@cifor-icraf.org'
    $alertMailPassword = ConvertTo-SecureString -String 'Winter2023' -AsPlainText -Force #Change to secure mode credential after testing
    $mailCredential = New-Object System.Management.Automation.PSCredential($alertMailUserName,$alertMailPassword)
if ($daysToCleanUp -eq 0)
{
    $message = @"   
Dear Administrator,

Please note that the attached list of users have been inactive for over 180 days and are due for cleanup. `r`n

Due to the lapse of the disabling window provided as per the alerts sent earlier, the accounts will be disabled and moved to the "Disabled Due To Inactivity" OU for deletion within one-month as per policy. `r`n `n
------------------------------------------------
$inactiveUsers
------------------------------------------------

CIFOR ICRAF Auto Report
"@
}
else
{
$message = @"   
Dear Administrator,

Please note that the attached list of users have been inactive for over 180 days. You have $daysToCleanUp day(s) to review and move the accounts to an OU that is exempted if wrongly placed. `r`n

The exempted OUs are: `r`n

ICRAF BOT | ICRAF Spouses | ICRAF Kenya General Accounts | ICRAF Administrators | ICRAF MFI | ICRAF Meeting Rooms | ICRAF ICT | ICRAF OCS | EC Regreening Project | ICRAF Sharepoint Users | ICRAF Kenya Shared Mailboxes | ICRAF Kenya Servie Accounts | OCS Users | ICRAF Regions`r`n

The exempted Security Groups are: `r`n

ICRAF MAC Clients | RemoteNonADLogins

At the end of the $daysToCleanUp-day(s) notice, the accounts will be disabled and moved to the "Disabled Due To Inactivity" OU for deletion within one-month as per policy. `r`n `n
------------------------------------------------
$inactiveUsers
------------------------------------------------

CIFOR ICRAF Auto Report
"@
}
    Send-MailMessage -to $recipients -From $alertMailUserName -Subject $subject -Body $message -Attachments $reportDirectory -SmtpServer $smtpServer -Port 587 -UseSsl -Credential $mailCredential
}
Export-ModuleMember -Function 'Send-Alert'