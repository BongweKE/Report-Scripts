# Load the Import-Excel module
Import-Module -Name ImportExcel
$reportDate = Get-Date -Format 'MMMM yyyy'
$AdminsGroup = 'Domain Admins'
$reportDirectory = 'C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Results\AD\ICRAF Admins\'
$reportDirectoryCurrent = $reportDirectory + $reportDate + '\'
$compressDirectory = $reportDirectoryCurrent + '*'
$compressedDirectory = $reportDirectoryCurrent + 'ADPriviledgedAccountsReport.zip'


#########################################################################
# RECHECK ON AD
#########################################################################
#Expired Admin Accounts
$ExpiredAdmin_Accounts = Get-ADGroupMember -Identity $AdminsGroup | Where-Object { $_.ObjectClass -eq 'user' -and $_.DistinguishedName -ne $null -and (Get-ADUser -Identity $_.DistinguishedName -Properties AccountExpirationDate).AccountExpirationDate -ne $null -and (Get-ADUser -Identity $_.DistinguishedName -Properties AccountExpirationDate).AccountExpirationDate -lt (Get-Date)}

#Active Admin Accounts
$ActiveAdminAccounts = Get-ADGroupMember -Identity $AdminsGroup | Where-Object { $_.ObjectClass -eq 'user' -and $_.DistinguishedName -ne $null -and (Get-ADUser -Identity $_.DistinguishedName -Properties LastLogonDate).LastLogonDate -gt (Get-Date).AddDays(-30) -and (Get-ADUser -Identity $_.DistinguishedName -Properties Enabled).Enabled -eq $true }

$dormantThreshold_30_days = (Get-Date).AddDays(-30)
$dormantThreshold_60_days = (Get-Date).AddDays(-60)

# Get all user accounts from the admin group with LastLogonDate and Enabled properties
$allAdminAccounts = Get-ADGroupMember -Identity $AdminsGroup | 
    Where-Object { $_.ObjectClass -eq 'user' -and $_.DistinguishedName -ne $null } | 
    Get-ADUser -Properties LastLogonDate, Enabled

# Filter for inactive accounts (regardless of enabled status)
$InactiveAdminAccounts = $allAdminAccounts | Where-Object { $_.LastLogonDate -lt $dormantThreshold_30_days }

# Filter for inactive and enabled accounts
$InactiveAdminAccountsEnabled = $InactiveAdminAccounts | Where-Object { $_.Enabled -eq $true }

# Filter for inactive and disabled accounts
$InactiveAdminAccountsDisabled = $InactiveAdminAccounts | Where-Object { $_.Enabled -eq $false }

# Filter for inactive accounts (60 days) regardless of enabled status
$InactiveAdminAccounts_60_days = $allAdminAccounts | Where-Object { $_.LastLogonDate -lt $dormantThreshold_60_days }

$InactiveAdminAccounts_60_days_Disabled = $InactiveAdminAccounts | Where-Object { $_.Enabled -eq $false }

$InactiveAdminAccounts_60_days_Enabled = $InactiveAdminAccounts | Where-Object { $_.Enabled -eq $true }
# No Expiry Admin Accounts
$NoExpiryAdminAccounts = Get-ADGroupMember -Identity $AdminsGroup | Where-Object { $_.ObjectClass -eq 'user' -and $_.DistinguishedName -ne $null -and (Get-ADUser -Identity $_.DistinguishedName -Properties accountExpires).accountExpires -eq 9223372036854775807 -and (Get-ADUser -Identity $_.DistinguishedName -Properties Enabled).Enabled -eq $true}

#Total Admin Accounts Count
$AdminAccounts = Get-ADGroupMember -Identity $AdminsGroup

#Generate Privileged Accounts Report Object
$ADPrivilegedAccountsReport = [pscustomobject]@{
'Report Date' = Get-Date -Format 'dd/MM/yyyy'
'No of Privileged Accounts' = $AdminAccounts.Count
'Expired Privileged Accounts' = $ExpiredAdmin_Accounts.Count
'Active Privileged Accounts' = $ActiveAdminAccounts.Count
'Dormant Privileged Accounts (30 Days)' = $InactiveAdminAccounts.Count
'Enabled Dormant Privileged Accounts (30 Days)' = $InactiveAdminAccountsEnabled.Count
'Disabled Dormant Privileged Accounts (30 Days)' = $InactiveAdminAccountsDisabled.Count
'Dormant Privileged Accounts (60 Days)' = $InactiveAdminAccounts_60_days.Count
'Enabled Dormant Privileged Accounts (60 Days)' = $InactiveAdminAccounts_60_days_Enabled.Count
'Disabled Dormant Privileged Accounts (60 Days)' = $InactiveAdminAccounts_60_days_Disabled.Count
'No Expiry Privileged Accounts' = $NoExpiryAdminAccounts.Count
}

# Archive Privileged Accounts Report in dashboard data source
$dashboardReportPath = 'C:\Users\lkadmin\CIFOR-ICRAF\Information Communication Technology (ICT) - Reports Archive\AD Reports\ad_privileged_accounts_dashboard_data.xlsx'

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
$sheet.Cells.Item($rowMax, 2).Value2 = $AdminAccounts.Count.ToString()
$sheet.Cells.Item($rowMax, 3).Value2 = $ExpiredAdmin_Accounts.Count.ToString()
$sheet.Cells.Item($rowMax, 4).Value2 = $ActiveAdminAccounts.Count.ToString()
$sheet.Cells.Item($rowMax, 5).Value2 = $InactiveAdminAccounts.Count.ToString()
$sheet.Cells.Item($rowMax, 6).Value2 = $InactiveAdminAccounts_60_days.Count.ToString()
$sheet.Cells.Item($rowMax, 7).Value2 = $NoExpiryAdminAccounts.Count.ToString()
$sheet.Cells.Item($rowMax, 8).Value2 = $InactiveAdminAccounts_60_days_Disabled.Count.ToString()
$sheet.Cells.Item($rowMax, 9).Value2 = $InactiveAdminAccounts_60_days_Enabled.Count.ToString()

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

#Store the report summary in a string
$reportBody_Summary = @"
No of Privileged Accounts: $($AdminAccounts.Count)`n
Expired Privileged Accounts: $($ExpiredAdmin_Accounts.Count)`n
Active Privileged Accounts: $($ActiveAdminAccounts.Count)`n
Dormant Privileged Accounts (Not used in the last 30 days): $($InactiveAdminAccounts.Count)`n
Dormant Privileged Accounts (Not used in the last 60 days): $($InactiveAdminAccounts_60_days.Count)`n
Enabled Dormant Privileged Accounts (Not used in the last 30 days): $($InactiveAdminAccountsEnabled.Count.ToString())`n
Disabled Dormant Privileged Accounts (Not used in the last 30 days): $($InactiveAdminAccountsDisabled.Count.ToString())`n
Enabled Dormant Privileged Accounts (Not used in the last 60 days): $($InactiveAdminAccounts_60_days_Enabled.Count.ToString())`n
Disabled Dormant Privileged Accounts (Not used in the last 60 days): $($InactiveAdminAccounts_60_days_Disabled.Count.ToString())`n
Privileged Accounts with No Expiry: $($NoExpiryAdminAccounts.Count)
"@

#Store the report expired accounts in a string
$reportBody_Expired = ""
foreach ($account in $ExpiredAdmin_Accounts)
{
   $reportBody_Expired += $account.name+"`n"
}

#Store the report active accounts in a string
$reportBody_Active = ""
foreach ($account in $ActiveAdminAccounts)
{
   $reportBody_Active += $account.name+"`n"
}

#Store the report inactive accounts (30 days) in a string
$reportBody_Inactive = ""
foreach ($account in $InactiveAdminAccounts)
{
   $reportBody_Inactive += $account.name+"`n"
}

#Store the report inactive accounts (60 days) in a string
$reportBody_Inactive_60_days = ""
foreach ($account in $InactiveAdminAccounts_60_days)
{
   $reportBody_Inactive_60_days += $account.name+"`n"
}

#Store the report no expiry accounts in a string
$reportBody_No_Expiry = ""
foreach ($account in $NoExpiryAdminAccounts)
{
   $reportBody_No_Expiry += $account.name+"`n"
}

$reportBody_Inactive_Enabled = ""
foreach ($account in $InactiveAdminAccountsEnabled)
{
   $reportBody_Inactive_Enabled += $account.name+"`n"
}


$reportBody_Inactive_Disabled = ""
foreach ($account in $InactiveAdminAccountsDisabled)
{
   $reportBody_Inactive_Disabled += $account.name+"`n"
}

$reportBody_Inactive_Enabled_60 = ""
foreach ($account in $InactiveAdminAccounts_60_days_Enabled)
{
   $reportBody_Inactive_Enabled_60 += $account.name+"`n"
}

$reportBody_Inactive_Disabled_60 = ""
foreach ($account in $InactiveAdminAccounts_60_days_Disabled)
{
   $reportBody_Inactive_Disabled_60 += $account.name+"`n"
}

#Generate Report Files
#Create Directory
mkdir $reportDirectoryCurrent -Force
$AdminAccounts | Select $selectObject | Export-Csv ($reportDirectoryCurrent + 'All Admin Accounts.csv') -NoTypeInformation
$ExpiredAdmin_Accounts | Select $selectObject | Export-Csv ($reportDirectoryCurrent + 'Expired Admin Accounts.csv') -NoTypeInformation
$ActiveAdminAccounts | Select $selectObject | Export-Csv ($reportDirectoryCurrent + 'Active Admin Accounts.csv') -NoTypeInformation
$InactiveAdminAccounts | Select $selectObject | Export-Csv ($reportDirectoryCurrent + 'Dormant Admin Accounts (Not used in the last 30 days).csv') -NoTypeInformation
$InactiveAdminAccountsEnabled | Select $selectObject | Export-Csv ($reportDirectoryCurrent + 'Enabled Dormant Admin Accounts (Not used in the last 30 days).csv') -NoTypeInformation
$InactiveAdminAccountsDisabled | Select $selectObject | Export-Csv ($reportDirectoryCurrent + 'Disabled Dormant Admin Accounts (Not used in the last 30 days).csv') -NoTypeInformation
$InactiveAdminAccounts_60_days | Select $selectObject | Export-Csv ($reportDirectoryCurrent + 'Dormant Admin Accounts (Not used in the last 60 days).csv') -NoTypeInformation
$InactiveAdminAccounts_60_days_Enabled | Select $selectObject | Export-Csv ($reportDirectoryCurrent + 'Enabled Dormant Admin Accounts (Not used in the last 60 days).csv') -NoTypeInformation
$InactiveAdminAccounts_60_days_Disabled | Select $selectObject | Export-Csv ($reportDirectoryCurrent + 'Disabled Dormant Admin Accounts (Not used in the last 60 days).csv') -NoTypeInformation
$NoExpiryAdminAccounts | Select $selectObject | Export-Csv ($reportDirectoryCurrent + 'Admin Accounts with No Expiry Date.csv') -NoTypeInformation
$ADPrivilegedAccountsReport | Export-Csv ($reportDirectoryCurrent + 'Privileged Accounts Full Report.csv') -NoTypeInformation

#Compress The Directory
Compress-Archive -Path $compressDirectory -DestinationPath $compressedDirectory -Force

#Send Email to Recipients
$smtpServer = 'SMTP.Office365.com'
$alertMailUserName = 'CIFORICRAFAutoReport@cifor-icraf.org'
$alertMailPassword = ConvertTo-SecureString -String 'Winter2023' -AsPlainText -Force #Change to secure mode credential after testing
$mailCredential = New-Object System.Management.Automation.PSCredential($alertMailUserName,$alertMailPassword)
$subject = 'CIFORICRAF Privileged Accounts AD Report - ' + $reportDate

$ICRAFReportRecipient = @('z.abidin@cifor-icraf.org','r.kande@cifor-icraf.org','p.oyuko@cifor-icraf.org','t.bandradi@cifor-icraf.org','c.mwangi@cifor-icraf.org','l.kavoo@cifor-icraf.org','i.dewantara@cifor-icraf.org','b.obaga@cifor-icraf.org')
# $ICRAFReportRecipient = @('l.kavoo@cifor-icraf.org', 'b.obaga@cifor-icraf.org')
# $ICRAFReportRecipient = @('b.obaga@cifor-icraf.org')

$attachments = @($compressedDirectory)
$message = @"   
Dear All,

Please find CIFORICRAF Privileged Accounts Active Directory report for $reportDate. You can use the attached report for further details. `r`n
Your action:- For your respective team:
i.	Review the membership - Incase of removal or addition, log a call for action at servicedesk@cifor-icraf.org for the amendments.
ii.	Review the dormancy report - Any dormant account should be disabled if dormant for 60days with the exception of system accounts.
iii.	Review the Non-expiry report – Any user admin account should have an expiry date consistent with their contract dates.
`r`n
------------------------------------------------
SUMMARY
------------------------------------------------
$reportBody_Summary


------------------------------------------------
EXPIRED ACCOUNTS
------------------------------------------------
$reportBody_Expired


------------------------------------------------
ACTIVE ACCOUNTS
------------------------------------------------
$reportBody_Active


------------------------------------------------
DORMANT ACCOUNTS (NOT USED IN THE LAST 30 DAYS)
------------------------------------------------
$reportBody_Inactive

------------------------------------------------
ENABLED DORMANT ACCOUNTS (NOT USED IN THE LAST 30 DAYS)
------------------------------------------------
$reportBody_Inactive_Enabled

------------------------------------------------
DISABLED DORMANT ACCOUNTS (NOT USED IN THE LAST 30 DAYS)
------------------------------------------------

$reportBody_Inactive_Disabled

------------------------------------------------
DORMANT ACCOUNTS (NOT USED IN THE LAST 60 DAYS)
------------------------------------------------
$reportBody_Inactive_60_days


------------------------------------------------
ENABLED DORMANT ACCOUNTS (NOT USED IN THE LAST 60 DAYS)
------------------------------------------------
$reportBody_Inactive_Enabled_60

------------------------------------------------
DISABLED DORMANT ACCOUNTS (NOT USED IN THE LAST 60 DAYS)
------------------------------------------------
$reportBody_Inactive_Disabled_60

------------------------------------------------
ACCOUNTS WITH NO EXPIRY DATE
------------------------------------------------
$reportBody_No_Expiry


CIFOR ICRAF Auto Report
"@

Send-MailMessage -to $ICRAFReportRecipient -From $alertMailUserName -Subject $subject -Body $message -SmtpServer $smtpServer -Port 587 -UseSsl -Credential $mailCredential -Attachments $attachments