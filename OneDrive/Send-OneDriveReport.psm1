function Send-OneDriveReport
{
    param ($recipients,$reportBody,$reportAttach)
    $smtpServer = 'SMTP.Office365.com'
    $alertMailUserName = 'CIFORICRAFAutoreport@cifor-icraf.org'
    $alertMailPassword = ConvertTo-SecureString -String 'Winter2023' -AsPlainText -Force #Change to secure mode credential after testing
    $mailCredential = New-Object System.Management.Automation.PSCredential($alertMailUserName,$alertMailPassword)
    $reportDate = Get-Date -Format 'MMMM dd yyyy'
    $subject = 'OneDrive Auto Report - ' + $reportDate
$message = @"   
Dear Administrator,

Please find the OneDrive report for the period of the past 2 weeks. `r`n
---------------------------------------------------------
$reportBody
---------------------------------------------------------

CIFOR ICRAF Auto Report
"@

    Send-MailMessage -to $recipients -From $alertMailUserName -Subject $subject -Body $message -Attachments $reportAttach -SmtpServer $smtpServer -Port 587 -UseSsl -Credential $mailCredential
}
Export-ModuleMember -Function *