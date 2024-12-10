# Load the Import-Excel module
Import-Module -Name ImportExcel
$reportDate = Get-Date -Format 'MMMM yyyy'
# $location = "Kenya"
# $alertLocation = "ICRAF"
$selectObject = @('Name', 'samAccountName', 'ObjectClass', 'AccountExpirationDate', 'lastLogonDate', 'Enabled', 'PasswordNeverExpires')
$ICRAF_ExcludedOU = @('OCS Users', 'ICRAF BOT','ICRAF Spouses','Disabled accounts','ICRAF Kenya General Accounts','ICRAF Administrators','ICRAF MFI','ICRAF Meeting Rooms','ICRAF ICT','ICRAF OCS','EC Regreening Project','ICRAF Sharepoint Users','ICRAF Kenya Shared Mailboxes','ICRAF Kenya Service Accounts') #update with KE data
$ICRAFOU = 'OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG'
$ICRAFComputersOU = 'OU=Computers,OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG'
# $inactiveDisabledOU = 'OU=Disabled Due To Inactivity,OU=Disabled accounts,OU=ICRAF Kenya,OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG'
$csvo365UsageData = 'C:\Users\lkadmin\CIFOR-ICRAF\Information Communication Technology (ICT) - Reports Archive\AD Reports\O365 Exchange Mail Usage\'

# Get the most recent file in the $csvo365UsageData folder
$mostRecentFile = Get-ChildItem -Path $csvo365UsageData -File | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1

# Store the path of the most recent file in a string variable
$csvFileName = $mostRecentFile.FullName
###########################################################################################################
# Change from Test folder before submiting to Tasker
###########################################################################################################
# $reportDirectory = 'C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Results\AD\ICRAF\Test'
$reportDirectory = 'C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Results\AD\ICRAF\'
$reportDirectoryCurrent = $reportDirectory + $reportDate + '\'
$compressDirectory = $reportDirectoryCurrent + '*'
$compressedDirectory = $reportDirectoryCurrent + 'ADReport.zip'

<# 
###########################################################################################################
# DO NOT DELETE THESE COMMENTS, VERY IMPORTANT                                                            #
###########################################################################################################
Whenever credentials change for CIFORICRAFAutoReport@cifor-icraf.org, we will have to run the lines below
to ensure that the new credentials are accessible via the task scheduler
# Prompt for credentials and store them in a variable
$credential = Get-Credential

# Extract the username and password separately
$username = $credential.UserName
$password = $credential.GetNetworkCredential().Password


# Create a secure AES key and save it to a file: AES Ensures that we can share the file between Users
# of our server: i.e LKAdmin and POAdmin can both execute this script without storing the passwords 
# on the script. 

# NOTE: this is a precursor to having most of our scripts on a github account owned by one of our 
# service accounts like AutoReports, etc.
# However, we'll have to ensure that the account with the github account always has a strong pass


$keyPath = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Secure\MySecureKey.key"
$encryptionKey = New-Object Byte[] 32
[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($encryptionKey)
[IO.File]::WriteAllBytes($keyPath, $encryptionKey)

# Make The Password Secure Using our Encryption algo
$securePassword = $credential.Password | ConvertFrom-SecureString -Key ([System.IO.File]::ReadAllBytes($keyPath))

# Store the username and encrypted password in a custom XML file
$customCredential = New-Object PSObject -Property @{
    Username = $credential.UserName
    Password = $securePassword
}
$customCredential | Export-Clixml -Path "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Secure\SharedCredential.xml"

# Future: Save Both in separate locations. Essentially, this is what'll ensure that we have safety. 
# For now, this is just a test to ensure that our scripts can now all go Credential Free
###########################################################################################################
#>

<#
Add this back when you figure it out

# Paths to the stored key and the credentials file
$keyPath = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Secure\MySecureKey.key"
$credentialPath = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Secure\SharedCredential.xml"

# Read the encryption key from the file
$encryptionKey = [System.IO.File]::ReadAllBytes($keyPath)

# Import the XML containing the saved credentials
$savedCredential = Import-Clixml -Path $credentialPath

# Decrypt the password using the key
$securePassword = ConvertFrom-SecureString -SecureString $savedCredential.Password -Key $encryptionKey

$savedCredential.Password | ConvertTo-SecureString -Key ([System.IO.File]::ReadAllBytes($keyPath))

$securePassword = $savedCredential.Password | ConvertFrom-SecureString -Key ([System.IO.File]::ReadAllBytes($keyPath))

# Create a PSCredential object using the decrypted password
$mailCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $alertMailUserName, $alertMailPassword
#>

$alertMailUserName = 'CIFORICRAFAutoReport@cifor-icraf.org'
$alertMailPassword = ConvertTo-SecureString -String 'Winter2023' -AsPlainText -Force #Change to secure mode credential after testing
$mailCredential = New-Object System.Management.Automation.PSCredential($alertMailUserName,$alertMailPassword)

###########################################################################################################
# Types of A/Cs


#Expired Accounts

$ICRAFExpiredAccounts = Search-ADAccount -SearchBase $ICRAFOU -AccountExpired | Where-Object 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|')

#Expired Accounts Above Six Month

$ICRAFExpiredAccounts_over180 = ([array](Search-ADAccount -SearchBase $ICRAFOU -AccountExpired | Where-Object 'AccountExpirationDate' -lt ((get-date).AddDays(-180)) | Where-Object 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|'))) 

#Accounts Without Expiry Date

$ICRAFNonExpiringAccounts = ([array](Get-ADUser -LDAPFilter '(|(accountExpires=0)(accountExpires=9223372036854775807))' -SearchBase $ICRAFOU | Where-Object 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|')))

#Inactive Accounts: 30 dAYS

$ICRAFInactiveAccounts = @()

#Inactive Accounts For 180 Days

$ICRAFInactiveAccounts_over180 = @()

#Inactive Computers For 180 Days

$ICRAFInactiveComputers_over180 = Search-ADAccount -SearchBase $ICRAFComputersOU -AccountInactive -TimeSpan 180.00:00:00 -ComputersOnly | Where-Object 'Enabled' -eq 'True' 

#Accounts With Password Set Not To Expire

$ICRAFAccounts_PasswordsNeverExpire = Search-ADAccount -SearchBase $ICRAFOU -PasswordNeverExpires | Where-Object 'Enabled' -eq 'True' | Where-Object 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|')

#Accounts Disabled for 2 months and above

$ICRAFDisabledAccounts_over60 = Search-ADAccount -SearchBase $ICRAFOU -AccountDisabled | Where-Object "LastLogonDate" -lt ((get-date).AddDays(-60))

#Placeholder for accounts with null last activity
$ICRAFAccountsNoLastActivity = @()
###########################################################################################################
# IDEA: ADD THE CODE ABOVE TO MODULE FOR REUSE IN AD Report scripts
###########################################################################################################

#Get o365 csv File
$csvReport = Import-Csv -Path $csvFileName
$mailboxCount = 0
# Report Refresh Date,,,Is Deleted,Deleted Date,Last Activity Date,Send Count,Receive Count,Read Count,Meeting Created Count,Meeting Interacted Count,Assigned Products,Report Period
foreach ($Account in $csvReport)
{
    $emailAddress = $Account.'User Principal Name'
    # $name = $Account.'Display Name'
    $lastActivity = $Account.'Last Activity Date'
    try {
        $lastActivity = [datetime]::ParseExact($lastActivity, 'yyyy-MM-dd', $null)
        if ($lastActivity -lt (Get-Date).AddDays(-30)) {
        # USERS inactive for 30 days will be added to $ICRAFInactiveAccounts
        $ICRAFInactiveAccounts += Get-ADUser -Filter{UserPrincipalName -eq $emailAddress} -Properties Name, samAccountName, ObjectClass, AccountExpirationDate, lastLogonDate, Enabled, PasswordNeverExpires -SearchBase $ICRAFOU | Where-Object 'Enabled' -eq 'True' | Where-Object 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|')
        
        }
        elseIf($lastActivity -lt (Get-Date).AddDays(-180)) {
        # USERS inactive for 180 days will be added to $ICRAFInactiveAccounts_over180
        $ICRAFInactiveAccounts_over180 += Get-ADUser -Filter{UserPrincipalName -eq $emailAddress} -Properties Name, samAccountName, ObjectClass, AccountExpirationDate, lastLogonDate, Enabled, PasswordNeverExpires -SearchBase $ICRAFOU | Where-Object 'Enabled' -eq 'True' | Where-Object 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|')

        }
        $theUserMailbox = Get-ADUser -Filter{UserPrincipalName -eq $emailAddress} -Properties * -SearchBase $ICRAFOU
        if ($null -eq $theUserMailbox) {
            $mailboxCount += 1
        }
    } catch {
        # Note that fetch/write problems for these a/cs could be caused by lack of key details
        $ICRAFAccountsNoLastActivity += Get-ADUser -Filter{UserPrincipalName -eq $emailAddress} -Properties Name, samAccountName, ObjectClass, AccountExpirationDate, Enabled, PasswordNeverExpires -SearchBase $ICRAFOU | Where-Object 'Enabled' -eq 'True' | Where-Object 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|')
        # Write-Host $emailAddress" has invalid last activity date"
    }
}
###########################################################################################################
#AD vs outlook: Comparing AD and outlook accounts to find all inactive accounts (AD inactive)

<#
Explanation for future updates:
- We have fetched Inactive accounts above using email addresses from outlook/azure
- We now find all inactive accounts in AD and if they're not part of the list from outlook,
  we add them to the list from outlook.
#>
###########################################################################################################

$ICRAFInactiveAccountsAD = Search-ADAccount -SearchBase $ICRAFOU -AccountInactive -TimeSpan 30.00:00:00 -UsersOnly | Where-Object 'Enabled' -eq 'True' | Where-Object 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|')

# Column on the original csv which we want to use for comparison
$columnName = 'Display Name'


foreach ($account in $ICRAFInactiveAccountsAD)
{
    # We'll search for `Name` attribute of each ADAccount
    $searchValue = $account.'Name'
 
    $result = $csvReport| Where-Object { $_.$columnName -eq $searchValue }
 
    if ($null -eq $result) {
        # We add the account to the list of inactive ones if we don't 
        # find it's Name from outlook csv empty ones
         $ICRAFInactiveAccounts += $account
    }
}

$ICRAFInactiveAccountsAD_180 = Search-ADAccount -SearchBase $ICRAFOU -AccountInactive -TimeSpan 180.00:00:00 -UsersOnly | Where-Object 'Enabled' -eq 'True' | Where-Object 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|')


foreach ($account in $ICRAFInactiveAccountsAD_180)
{
    # We'll search for `Name` attribute of each ADAccount
    $searchValue = $account.'Name'
 
    $result = $csvReport| Where-Object { $_.$columnName -eq $searchValue }
 
    if ($null -eq $result) {
        # We add the account to the list of inactive ones if we don't 
        # find it's Name from outlook csv empty ones
        $ICRAFInactiveAccounts_over180 += $account
    }
}


###########################################################################################################
#Generate Report From Exchange Online Session
###########################################################################################################
$userAccountCount = (Get-ADUser -Filter * -SearchBase $ICRAFOU).count
$computerCount = (Get-ADComputer -Filter * -SearchBase $ICRAFComputersOU).count

# Archive Full Report in dashboard data source
###########################################################################################################
# Use test Excel file before we add this to the task scheduler (pipeline)
###########################################################################################################
$dashboardReportPath = 'C:\Users\lkadmin\CIFOR-ICRAF\Information Communication Technology (ICT) - Reports Archive\AD Reports\ad_dashboard_data.xlsx'

#$dashboardReportPath = 'C:\Users\lkadmin\CIFOR-ICRAF\Information Communication Technology (ICT) - Reports Archive\AD Reports\Test\ad_dashboard_data.xlsx'
#Generate Report Object

$ADFullReport = [pscustomobject]@{
'Report Date' = Get-Date -Format 'dd/MM/yyyy'
'No of MailBoxes' = $mailboxCount
'No of User Accounts' = $userAccountCount
'No of Computers' = $computerCount
'Expired Accounts' = $ICRAFExpiredAccounts.Count
'Expired Accounts Over 6 Months' = $ICRAFExpiredAccounts_over180.Count
'Accounts Without Expiry Date' = $ICRAFNonExpiringAccounts.Count
'Inactive Accounts' = $ICRAFInactiveAccounts.Count
'Inactive Accounts Over 6 Months' = $ICRAFInactiveAccounts_over180.Count
'Inactive Computers Over 6 Months' = $ICRAFInactiveComputers_over180.Count
'Accounts With Non Expiring Passwords' = $ICRAFAccounts_PasswordsNeverExpire.Count
'Disabled Accounts Over 2 Months' = $ICRAFDisabledAccounts_over60.Count
}

# Write-Output $ADFullReport


# Create a new Excel application object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# Open the Excel file
$workbook = $excel.Workbooks.Open($dashboardReportPath)
$sheet = $workbook.Worksheets.Item(1)

# Find the last used row in the first column
$rowMax = $sheet.UsedRange.Rows.Count + 1

# Add the new data to the next row
$sheet.Cells.Item($rowMax, 1).Value2 = Get-Date -Format 'MM/dd/yyyy'
$sheet.Cells.Item($rowMax, 2).Value2 = $mailboxCount.ToString()
$sheet.Cells.Item($rowMax, 3).Value2 = $userAccountCount.ToString()
$sheet.Cells.Item($rowMax, 4).Value2 = $computerCount.ToString()
$sheet.Cells.Item($rowMax, 5).Value2 = $ICRAFExpiredAccounts.Count.ToString()
$sheet.Cells.Item($rowMax, 6).Value2 = $ICRAFExpiredAccounts_over180.Count.ToString()
$sheet.Cells.Item($rowMax, 7).Value2 = $ICRAFNonExpiringAccounts.Count.ToString()
$sheet.Cells.Item($rowMax, 8).Value2 = $ICRAFInactiveAccounts.Count.ToString()
$sheet.Cells.Item($rowMax, 9).Value2 = $ICRAFInactiveAccounts_over180.Count.ToString()
$sheet.Cells.Item($rowMax, 10).Value2 = $ICRAFInactiveComputers_over180.Count.ToString()
$sheet.Cells.Item($rowMax, 11).Value2 = $ICRAFAccounts_PasswordsNeverExpire.Count.ToString()
$sheet.Cells.Item($rowMax, 12).Value2 = $ICRAFDisabledAccounts_over60.Count.ToString()


# Save and close the workbook
$workbook.Save()
$workbook.Close($true)
Start-Sleep -Seconds 2
# Quit Excel and release COM objects
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

#Store the report in a string
$reportBody = @"
No of MailBox: $mailboxCount
No of User Accounts: $userAccountCount
No of Computers: $computerCount
Expired Accounts: $($ICRAFExpiredAccounts.Count)
Expired Accounts Over 6 Months: $($ICRAFExpiredAccounts_over180.Count)
Accounts Without Expire Date: $($ICRAFNonExpiringAccounts.Count)
Inactive Accounts: $($ICRAFInactiveAccounts.Count)
Inactive Accounts Over 6 Months: $($ICRAFInactiveAccounts_over180.Count)
Inactive Computers Over 6 Months: $($ICRAFInactiveComputers_over180.Count)
Accounts without Last Activity: $($ICRAFAccountsNoLastActivity.Count)
Accounts With Non Expiring Passwords: $($ICRAFAccounts_PasswordsNeverExpire.Count)
Disabled Accounts Over 2 Months: $($ICRAFDisabledAccounts_over60.Count)
"@

#Generate Report Files
#Create Directory
mkdir $reportDirectoryCurrent -Force
$ICRAFExpiredAccounts | Select-Object $selectObject | Export-Csv ($reportDirectoryCurrent + 'Expired Accounts.csv') -NoTypeInformation
$ICRAFExpiredAccounts_over180 | Select-Object $selectObject | Export-Csv ($reportDirectoryCurrent + 'Expired Accounts Over 6 Months.csv') -NoTypeInformation
$ICRAFNonExpiringAccounts | Select-Object $selectObject | Export-Csv ($reportDirectoryCurrent + 'Accounts Without Expire Date.csv') -NoTypeInformation
$ICRAFInactiveAccounts | Select-Object $selectObject | Export-Csv ($reportDirectoryCurrent + 'Inactive Accounts.csv') -NoTypeInformation
$ICRAFInactiveAccounts_over180 | Select-Object $selectObject | Export-Csv ($reportDirectoryCurrent + 'Inactive Accounts Over 6 Months.csv') -NoTypeInformation
$ICRAFInactiveComputers_over180 | Select-Object $selectObject | Export-Csv ($reportDirectoryCurrent + 'Inactive Computers Over 6 Months.csv') -NoTypeInformation
$ICRAFAccounts_PasswordsNeverExpire | Select-Object $selectObject | Export-Csv ($reportDirectoryCurrent + 'Accounts With Non Expiring Passwords.csv') -NoTypeInformation
$ICRAFDisabledAccounts_over60 | Select-Object $selectObject | Export-Csv ($reportDirectoryCurrent + 'Disabled Accounts Over 2 Months.csv') -NoTypeInformation
$csvReport | Export-Csv ($reportDirectoryCurrent + 'Exchange Activity Data.csv') -NoTypeInformation
$ICRAFAccountsNoLastActivity | Export-Csv ($reportDirectoryCurrent + 'Accounts Without Last Activity on o365 Exchange.csv') -NoTypeInformation
$ADFullReport | Export-Csv ($reportDirectoryCurrent + 'Full Report.csv') -NoTypeInformation

#Compress The Directory
Compress-Archive -Path $compressDirectory -DestinationPath $compressedDirectory -Force



#The below strings are appended as they are (including the reportdate variable) in the message body. I have returned it as it was in line 277. You could find a fix for that in future.
$info_msg = "Please find ICRAF Active Directory report for $reportDate . You can use the attached report for further details."
# $test_info_msg = 'This is a routine test Please find ICRAF Active Directory report for $reportDate. You can use the attached TEST report for further details.'

#Send Email to Recipients
$smtpServer = 'SMTP.Office365.com'
$subject = 'ICRAF AD Report - ' + $reportDate
$ICRAFReportRecipient = @('r.kande@cifor-icraf.org', 'p.oyuko@cifor-icraf.org', 'c.mwangi@cifor-icraf.org', 'l.kavoo@cifor-icraf.org', 'servicedesk@cifor-icraf.org', 'b.obaga@cifor-icraf.org','g.kirimi@cifor-icraf.org')
# $ICRAFReportRecipient = @('p.oyuko@cifor-icraf.org', 'c.mwangi@cifor-icraf.org', 'l.kavoo@cifor-icraf.org', 'servicedesk@cifor-icraf.org', 'b.obaga@cifor-icraf.org','g.kirimi@cifor-icraf.org')
# $ICRAFReportRecipient = @('l.kavoo@cifor-icraf.org','b.obaga@cifor-icraf.org','g.kirimi@cifor-icraf.org')
# $ICRAFReportRecipient = @('b.obaga@cifor-icraf.org')
$attachments = @($compressedDirectory)
$message = @"   
Dear Administrator,

$info_msg `r`n
------------------------------------------------
$reportBody
------------------------------------------------

CIFOR ICRAF Auto Report
"@

Send-MailMessage -to $ICRAFReportRecipient -From $alertMailUserName -Subject $subject -Body $message -SmtpServer $smtpServer -Port 587 -UseSsl -Credential $mailCredential -Attachments $attachments
<#
Source File 
C:\Users\lkadmin\CIFOR-ICRAF\Information Communication Technology (ICT) - Reports Archive\AD Reports\O365 Exchange Mail Usage
Report Refresh Date,User Principal Name,Display Name,Is Deleted,Deleted Date,Last Activity Date,Send Count,Receive Count,Read Count,Meeting Created Count,Meeting Interacted Count,Assigned Products,Report Period


#Inactive Accounts

$ICRAFInactiveAccounts = Search-ADAccount -SearchBase $ICRAFOU -AccountInactive -TimeSpan 30.00:00:00 -UsersOnly | Where 'Enabled' -eq 'True' | Where 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|')

#Inactive Accounts For 180 Days

$ICRAFInactiveAccounts_over180 = Search-ADAccount -SearchBase $ICRAFOU -AccountInactive -TimeSpan 180.00:00:00 -UsersOnly | Where 'Enabled' -eq 'True' | Where 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|')
$ICRAFInactiveAccounts_over180 = $ICRAFInactiveAccounts_over180 | Get-ADUser -Properties WhenCreated | Where WhenCreated -lt (Get-Date).AddDays(-30)

Output Files
C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Results\AD\ICRAF\July 2024\
Inactive Accounts.csv
"Name","samAccountName","ObjectClass","AccountExpirationDate","lastLogonDate","Enabled","PasswordNeverExpires"

Inactive Accounts Over 6 Months.csv
"Name","samAccountName","ObjectClass","AccountExpirationDate","lastLogonDate","Enabled","PasswordNeverExpires"
 -Properties Name, samAccountName, ObjectClass, AccountExpirationDate, lastLogonDate, Enabled, PasswordNeverExpires
ForEach-Object { 
    "{0},{1},{2},{3:yyyy-MM-dd},{4:yyyy-MM-dd},{5},{6}" -f $_.Name, $_.samAccountName, $_.ObjectClass, $_.AccountExpirationDate, $_.lastLogonDate, $_.Enabled -eq $true, $_.PasswordNeverExpires -eq $true
  }
#>
