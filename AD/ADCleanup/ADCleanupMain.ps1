Import-Module "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Scripts\AD\ADCleanup\ADCleanup-Disable.psm1" -Verbose -Force
Import-Module "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Scripts\AD\ADCleanup\ADCleanup-Alert.psm1" -Verbose -Force

$reportDate = Get-Date -Format 'MMMM yyyy'
$location = "Kenya"
$alertLocation = "ICRAF"
$ICRAF_ExcludedOU = @('ICRAF BOT','ICRAF Spouses','Disabled accounts','ICRAF Kenya General Accounts','ICRAF Administrators','ICRAF MFI','ICRAF Meeting Rooms','ICRAF ICT','ICRAF OCS','EC Regreening Project','ICRAF Sharepoint Users','ICRAF Kenya Shared Mailboxes','ICRAF Kenya Service Accounts', 'OCS Users', 'ICRAF Regions') #update with KE data
$MacUsers = (Get-ADGroup 'ICRAF MAC Clients').DistinguishedName
$Mac_usernames = Get-ADUser -Filter { memberof -eq $MacUsers -and enabled -eq $true} | Select-Object -ExpandProperty Name
$nonAdLoginUsers = (Get-ADGroup 'RemoteNonADLogins').DistinguishedName
$non_ad_login_usernames = Get-ADUser -Filter { memberof -eq $nonAdLoginUsers -and enabled -eq $true} | Select-Object -ExpandProperty Name
$ICRAFOU = 'OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG'
$inactiveDisabledOU = 'OU=Disabled Due To Inactivity,OU=Disabled accounts,OU=ICRAF Kenya,OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG'
$reportDirectory = 'C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Results\AD\ICRAF\'
$reportDirectoryCurrent = $reportDirectory + $reportDate + '\'


$reportRecipients = @('r.kande@cifor-icraf.org','servicedesk@cifor-icraf.org','c.mwangi@cifor-icraf.org','p.oyuko@cifor-icraf.org','l.kavoo@cifor-icraf.org', 'b.obaga@cifor-icraf.org', 'g.kirimi@cifor-icraf.org')
# $reportRecipients = @('servicedesk@cifor-icraf.org','c.mwangi@cifor-icraf.org','p.oyuko@cifor-icraf.org','l.kavoo@cifor-icraf.org')
# $reportRecipients = @('l.kavoo@cifor-icraf.org', 'b.obaga@cifor-icraf.org')
# $reportRecipients = @('b.obaga@cifor-icraf.org')
$ICRAFAlertRecipients = $reportRecipients 
$date_180_days_ago = (get-date).AddDays(-180)

#Inactive Accounts For 180 Days
$ICRAFInactiveAccounts_over180_1 = Get-ADUser -SearchBase $ICRAFOU -Filter {LastLogonDate -lt $date_180_days_ago} -Properties * | 
Where-Object 'Enabled' -eq 'True' |
Where-Object 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|') |
Where-Object Name -notin $Mac_usernames |
Where-Object Name -notin $non_ad_login_usernames |
Where-Object WhenCreated -lt (Get-Date).AddDays(-30) |
Select-Object DistinguishedName, Enabled, GivenName, Name, SamAccountName, Surname, UsePrincipalName, WhenCreated

if($ICRAFInactiveAccounts_over180_1.count -gt 0)
{
$firstNotficationDelay = (((get-date).AddDays(6)) - (get-date)).totalseconds
#$firstNotficationDelay = ((Get-Date).AddSeconds(30) - (get-date)).totalseconds
$reportString = $ICRAFInactiveAccounts_over180_1.Name

$reportString = $reportString -join "`n"


#Generate Inactive Accounts List To CSV
$inactiveUsers_Report = $reportDirectoryCurrent + 'Inactive Accounts Cleanup List.csv'

$ICRAFInactiveAccounts_over180_1 | Export-Csv -Path $inactiveUsers_Report -NoTypeInformation

#Send alert on inactive users. 
Send-Alert -recipients $ICRAFAlertRecipients -reportDirectory $inactiveUsers_Report -inactiveUsers $reportString -daysToCleanUp 7

#Pausing Script For 6 Days
Start-Sleep -Seconds $firstNotficationDelay

#Send Reminder
$ICRAFInactiveAccounts_over180_2 = Get-ADUser -SearchBase $ICRAFOU -Filter {LastLogonDate -lt $date_180_days_ago} -Properties * | 
Where-Object 'Enabled' -eq 'True' |
Where-Object 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|') |
Where-Object Name -notin $Mac_usernames |
Where-Object Name -notin $non_ad_login_usernames |
Where-Object WhenCreated -lt (Get-Date).AddDays(-30) |
Select-Object DistinguishedName, Enabled, GivenName, Name, SamAccountName, Surname, UsePrincipalName, WhenCreated

$reportString = $ICRAFInactiveAccounts_over180_2.Name
$reportString = $reportString -join "`n"
$secondNotificationDelay = ((get-date).AddDays(1) - (get-date)).totalseconds
#$secondNotificationDelay = ((Get-Date).AddSeconds(30) - (get-date)).totalseconds

#Generate Inactive Accounts List To CSV
$inactiveUsers_Report = $reportDirectoryCurrent + 'Inactive Accounts Cleanup List.csv'

$ICRAFInactiveAccounts_over180_2 | Export-Csv -Path $inactiveUsers_Report -NoTypeInformation

#Send reminder alert on inactive users. 
Send-Alert -recipients $ICRAFAlertRecipients -reportDirectory $inactiveUsers_Report -inactiveUsers $reportString -daysToCleanUp 1

#Pausing Script For 1 Day
Start-Sleep -Seconds $secondNotificationDelay

#Send Notice of Disabled Accounts
$ICRAFInactiveAccounts_over180_3 = Get-ADUser -SearchBase $ICRAFOU -Filter {LastLogonDate -lt $date_180_days_ago} -Properties * | 
Where-Object 'Enabled' -eq 'True' |
Where-Object 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|') |
Where-Object Name -notin $Mac_usernames |
Where-Object Name -notin $non_ad_login_usernames |
Where-Object WhenCreated -lt (Get-Date).AddDays(-30) |
Select-Object DistinguishedName, Enabled, GivenName, Name, SamAccountName, Surname, UsePrincipalName, WhenCreated

$reportString = $ICRAFInactiveAccounts_over180_3.Name
$reportString = $reportString -join "`n"

#Generate Inactive Accounts List To CSV
$inactiveUsers_Report = $reportDirectoryCurrent + 'Inactive Accounts Cleanup List.csv'

$ICRAFInactiveAccounts_over180_3 | Export-Csv -Path $inactiveUsers_Report -NoTypeInformation

#Send alert on user accounts being disabled. 
Send-Alert -recipients $ICRAFAlertRecipients -reportDirectory $inactiveUsers_Report -inactiveUsers $reportString -daysToCleanUp 0

<#
----------------------------------------------------------------------------------
Account-Disabling uses the $inactiveUsers_Report to disable said users
----------------------------------------------------------------------------------
#>
Account-Disabling -reportDirectory $inactiveUsers_Report
}
else
{
  Write-Output "No Inactive Accounts Found"
}
