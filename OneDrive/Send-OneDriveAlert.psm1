#Sending Alert
function Log-Message([String]$Message)
{
    $logFilePath = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Scripts\OneDrive\OneDriveAlertLogs.txt"
    Add-Content -Path $logFilePath $Message
}

function Send-OneDriveAlert
{
    param($UsersOldActivityDate_NotGeneral,$UsersLowStorageUtilization_NotGeneral)
    $inactiveOnly = @()
    $inactiveNLowUtilization = @()
    $lowUtilizationOnly = @()
    $log_message = "Date: "+(Get-Date).ToString("dd-MM-yyyy")+"`n"
    $log_message += "Beginning exeuction of the script...`n"
    #Separate users with low usage and old activity date
    foreach ($user in $UsersOldActivityDate_NotGeneral)
    {
        if ($user -in $UsersLowStorageUtilization_NotGeneral){$inactiveNLowUtilization += $user}
        else {$inactiveOnly += $user}
    }

    foreach ($user in $UsersLowStorageUtilization_NotGeneral)
    {
        if ($user -notin $inactiveNLowUtilization){$lowUtilizationOnly += $user}
    }

    #define O365 credentials for sending the email alerts.
    $smtpServer = 'SMTP.Office365.com'
    $alertMailUserName = 'CIFORICRAFAutoreport@cifor-icraf.org'
    $alertMailPassword = ConvertTo-SecureString -String 'Winter2023' -AsPlainText -Force #Change to secure mode credential after testing
    $alertCredentials = New-Object System.Management.Automation.PSCredential($alertMailUserName,$alertMailPassword)
    $test_recepient = @('s.mariwa@cifor-icraf.org') # comment when setting for production and use $recepient in its place
    $onedrivePresentationDirectory = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Documents\OneDrive Guide.pptx"
    #Send Alert to Users with Old Activity Date 
    foreach ($users in $inactiveOnly)
    {
    $userUtilization = [math]::Round($users.StorageUsageinGB,2)
    $recepient = $users.Email
    $user = $users.DisplayName
    $log_message += "User:"+$user+" User_Email_Address:"+$recepient+"  Alert: Old Activity Date`n"
    Write-Host $log_message
    $alert = @"
Dear $user,</br></br>

<p>We've noticed that your <b>OneDrive cloud backup has not synchronized in the last 15 days</b>. To ensure the safety of your documents and to maintain efficient collaboration within our organization, we kindly request your attention to the following steps:</p>
<ol>
<li>Check Your Document Storage: Please confirm that you are storing your documents in the designated OneDrive directory. This ensures that your files are included in the synchronization process.</li>

<li>Verify Synchronization Status: To check the synchronization status, click the OneDrive icon located in the taskbar of your computer. If you encounter any sync errors, make sure to follow the instructions provided alongside the errors to resolve them. Certain errors can impede a complete sync of your documents on OneDrive.</li>

<li>Monitor Sync Status Icons: Keep an eye on the synchronization status icons on the files stored in your OneDrive directories. These icons provide valuable feedback on the status of your files and folders. For more information, please refer to the PowerPoint presentation attached.</li>
</ol>
<p>If you encounter any challenges or have questions while performing these steps, please don't hesitate to reach out to our Service Desk at servicedesk@cifor-icraf.org. Our dedicated team is here to assist you promptly.</p>

<p>Ensuring that your OneDrive is actively synchronized is essential for safeguarding your data and maintaining seamless collaboration across the organization. We appreciate your cooperation in this matter.</p>

<p>Thank you for your attention to this important task, $user. Your proactive involvement helps us maintain the integrity of our data and keeps our workflows efficient.</p></br>

Best regards,
"@
    #Changed the 'to' parameter to the recipient variable when ready.
    Send-MailMessage -to $recepient -From $alertMailUserName -Subject 'OneDriveAlert' -Body $alert -BodyAsHtml -Attachments $onedrivePresentationDirectory -SmtpServer $smtpServer -Port 587 -UseSsl -Credential $alertCredentials
    Start-Sleep -s 60
   
    }

    #Send Alert to Users with Low Storage Utilization
    foreach ($users in $lowUtilizationOnly)
    {
    $userUtilization = [math]::Round($users.StorageUsageinGB,2)
    $recepient = $users.Email
    $user = $users.DisplayName
    $log_message += "User:"+$user+" User_Email_Address:"+$recepient+"  Alert: Low Utilization  Utilization:"+$userUtilization+"GB`n"
    Write-Host $log_message
    $alert = @"
Dear $user,</br></br>

<p>We've noticed that your <b>OneDrive storage utilization is $userUtilization GB which is below the expected threshold</b>. To ensure the safety of your documents and to maintain efficient collaboration within our organization, we kindly request your attention to the following steps:</p>
<ol>
<li>Check Your Document Storage: Please confirm that you are storing your documents in the designated OneDrive directory. This ensures that your files are included in the synchronization process.</li>

<li>Verify Synchronization Status: To check the synchronization status, click the OneDrive icon located in the taskbar of your computer. If you encounter any sync errors, make sure to follow the instructions provided alongside the errors to resolve them. Certain errors can impede a complete sync of your documents on OneDrive.</li>

<li>Monitor Sync Status Icons: Keep an eye on the synchronization status icons on the files stored in your OneDrive directories. These icons provide valuable feedback on the status of your files and folders. For more information, please refer to the PowerPoint presentation attached.</li>
</ol>
<p>If you encounter any challenges or have questions while performing these steps, please don't hesitate to reach out to our Service Desk at servicedesk@cifor-icraf.org. Our dedicated team is here to assist you promptly.</p>

<p>Ensuring that your OneDrive is actively synchronized is essential for safeguarding your data and maintaining seamless collaboration across the organization. We appreciate your cooperation in this matter.</p>

<p>Thank you for your attention to this important task, $user. Your proactive involvement helps us maintain the integrity of our data and keeps our workflows efficient.</p></br>

Best regards,
"@
    #Change the 'to' parameter to the recipient variable when ready.
    Send-MailMessage -to $recepient -From $alertMailUserName -Subject 'OneDriveAlert' -Body $alert -BodyAsHtml -Attachments $onedrivePresentationDirectory -SmtpServer $smtpServer -Port 587 -UseSsl -Credential $alertCredentials
    Start-Sleep -s 60
   
    }

    #Send Alert to Users with Old Activity Date and Low Utilization
    foreach ($users in $inactiveNLowUtilization)
    {
    $userUtilization = [math]::Round($users.StorageUsageinGB,2)
    $recepient = $users.Email
    $user = $users.DisplayName
    $log_message += "User:"+$user+" User_Email_Address:"+$recepient+"  Alert: Old Activity Date & Low Utilization  Utilization:"+$userUtilization+"GB`n"
    Write-Host $log_message
    $alert = @"
Dear $user,`r`n

<p>We've noticed that your <b>OneDrive storage utilization is 3 GB which is below the expected threshold</b> and <b>synchronization has not happened in the last 15 days</b>. To ensure the safety of your documents and to maintain efficient collaboration within our organization, we kindly request your attention to the following steps:</p>
<ol>
<li>Check Your Document Storage: Please confirm that you are storing your documents in the designated OneDrive directory. This ensures that your files are included in the synchronization process.</li>

<li>Verify Synchronization Status: To check the synchronization status, click the OneDrive icon located in the taskbar of your computer. If you encounter any sync errors, make sure to follow the instructions provided alongside the errors to resolve them. Certain errors can impede a complete sync of your documents on OneDrive.</li>

<li>Monitor Sync Status Icons: Keep an eye on the synchronization status icons on the files stored in your OneDrive directories. These icons provide valuable feedback on the status of your files and folders. For more information, please refer to the PowerPoint presentation attached.</li>
</ol>
<p>If you encounter any challenges or have questions while performing these steps, please don't hesitate to reach out to our Service Desk at servicedesk@cifor-icraf.org. Our dedicated team is here to assist you promptly.</p>

<p>Ensuring that your OneDrive is actively synchronized is essential for safeguarding your data and maintaining seamless collaboration across the organization. We appreciate your cooperation in this matter.</p>

<p>Thank you for your attention to this important task, $user. Your proactive involvement helps us maintain the integrity of our data and keeps our workflows efficient.</p></br>

Best regards,
"@
    #Change the 'to' parameter to the recipient variable when ready.
    Send-MailMessage -to $recepient -From $alertMailUserName -Subject 'OneDriveAlert' -Body $alert -BodyAsHtml -Attachments $onedrivePresentationDirectory -SmtpServer $smtpServer -Port 587 -UseSsl -Credential $alertCredentials
    Start-Sleep -s 60
   
    }
    $log_message += "Completed exeuction of the script`n`n"
    Log-Message $log_message

}
Export-ModuleMember -Function *