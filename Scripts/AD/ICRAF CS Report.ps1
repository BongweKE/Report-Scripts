$currentDate = Get-Date
$dateInterval = $currentDate.AddDays(14)
$accountPasswordsToExpire = @()
$ICRAFOU = 'OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG'
$ICRAF_ExcludedOU = @('Disabled accounts','ICRAF Kenya Service Accounts','ICRAF Kenya Shared Mailboxes','ICRAF Kenya General Accounts')

$ICRAFUsers = Get-ADUser -SearchBase 'OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG' -Properties msDS-UserPasswordExpiryTimeComputed -Filter *
foreach($users in $ICRAFUsers)
{
  
    $passWordExpiry = [datetime]::FromFileTime($users.'msDS-UserPasswordExpiryTimeComputed')
    
   
        if(($passWordExpiry -ge $currentDate) -and ($passWordExpiry -le $dateInterval)) 
        {
            $ErrorActionPreference = 'silentlycontinue'
            $passWordExpiry
            $reportObject = [PSCustomObject]@{
            Name = $users.Name
            UserName = $users.SamAccountName
            PasswordExpiryDate = $passWordExpiry
            }

            $accountPasswordsToExpire += $reportObject
        }
    

}


#Expired Accounts

$ICRAFExpiredAccounts = Search-ADAccount -SearchBase $ICRAFOU -AccountExpired | Select-Object Name,DistinguishedName,SamAccountName,AccountExpirationDate | Where 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|') 

#Accounts Without Expiry Date

$ICRAFNonExpiringAccounts = Get-ADUser -LDAPFilter '(|(accountExpires=0)(accountExpires=9223372036854775807))' -SearchBase $ICRAFOU | Select-Object Name,DistinguishedName,SamAccountName | Where 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|') 

#Accounts With Password Set Not To Expire

$ICRAFAccounts_PasswordsNeverExpire = Search-ADAccount -SearchBase $ICRAFOU -PasswordNeverExpires | Select-Object Name,DistinguishedName,SamAccountName | Where 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|') 

#Accounts About To Expire

$ICRAFAccountsAboutToExpire = Search-ADAccount -SearchBase $ICRAFOU -AccountExpiring -TimeSpan 14.00:00:00 | Select-Object Name,DistinguishedName,SamAccountName,AccountExpirationDate | Where 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|') 

#Export Reports


$accountPasswordsToExpire | Export-Csv -Path 'C:\Users\lkavoo\OneDrive - CIFOR-ICRAF\Documents\Auto Reports\Report Results\ICRAF CS Report\Accounts With Password About To Expire.csv' -NoTypeInformation
$ICRAFExpiredAccounts | Export-Csv -Path 'C:\Users\lkavoo\OneDrive - CIFOR-ICRAF\Documents\Auto Reports\Report Results\ICRAF CS Report\Expired Accounts.csv' -NoTypeInformation
$ICRAFNonExpiringAccounts | Export-Csv -Path 'C:\Users\lkavoo\OneDrive - CIFOR-ICRAF\Documents\Auto Reports\Report Results\ICRAF CS Report\Accounts Set Not To Expire.csv' -NoTypeInformation
$ICRAFAccounts_PasswordsNeverExpire | Export-Csv -Path 'C:\Users\lkavoo\OneDrive - CIFOR-ICRAF\Documents\Auto Reports\Report Results\ICRAF CS Report\Accounts With Never Expiring Passwords.csv' -NoTypeInformation
$ICRAFAccountsAboutToExpire | Export-Csv -Path 'C:\Users\lkavoo\OneDrive - CIFOR-ICRAF\Documents\Auto Reports\Report Results\ICRAF CS Report\Accounts About To Expire.csv' -NoTypeInformation

#Send Email to Recipients
$smtpServer = 'SMTP.Office365.com'
$alertMailUserName = 'CIFORICRAFAutoReport@cifor-icraf.org'
$alertMailPassword = ConvertTo-SecureString -String 'Winter2023' -AsPlainText -Force #Change to secure mode credential after testing
$mailCredential = New-Object System.Management.Automation.PSCredential($alertMailUserName,$alertMailPassword)
$reportDate = Get-Date -Format 'MMMM dd yyyy'
$subject = 'ICRAF Accounts Report - ' + $reportDate
$ICRAFReportRecipient = @('l.kavoo@cifor-icraf.org','l.kavoo@cgiar.org')
$attachments = @('C:\Users\lkavoo\OneDrive - CIFOR-ICRAF\Documents\Auto Reports\Report Results\ICRAF CS Report\Accounts About To Expire.csv','C:\Users\lkavoo\OneDrive - CIFOR-ICRAF\Documents\Auto Reports\Report Results\ICRAF CS Report\Expired Accounts.csv', 'C:\Users\lkavoo\OneDrive - CIFOR-ICRAF\Documents\Auto Reports\Report Results\ICRAF CS Report\Accounts Set Not To Expire.csv','C:\Users\lkavoo\OneDrive - CIFOR-ICRAF\Documents\Auto Reports\Report Results\ICRAF CS Report\Accounts With Never Expiring Passwords.csv', 'C:\Users\lkavoo\OneDrive - CIFOR-ICRAF\Documents\Auto Reports\Report Results\ICRAF CS Report\Accounts With Password About To Expire.csv')

$message = @"   
Greetings,

Please find ICRAF User Accounts report for the period of the past 2 weeks. `r`n
CIFOR ICRAF Auto Report
"@

Send-MailMessage -to $ICRAFReportRecipient -From $alertMailUserName -Subject $subject -Body $message -SmtpServer $smtpServer -Port 587 -UseSsl -Credential $mailCredential -Attachments $attachments