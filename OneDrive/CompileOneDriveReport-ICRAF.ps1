Import-Module "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Reports\ESET and OneDrive Scripts\Get-OneDriveReport.psm1" -Verbose -Force
Import-Module "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Reports\ESET and OneDrive Scripts\Send-OneDriveReport.psm1" -Verbose -Force
Import-Module "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Reports\ESET and OneDrive Scripts\Send-OneDriveAlert.psm1" -Verbose -Force
Import-Module "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Reports\ESET and OneDrive Scripts\GetKaseyaMachineCount.psm1" -Verbose -Force

#Get Date
$reportDate = Get-Date -Format "MMMM dd yyyy"

#Define Report Recipients
# $reportRecipients = @('r.kande@cifor-icraf.org','servicedesk@cifor-icraf.org','c.mwangi@cifor-icraf.org','p.oyuko@cifor-icraf.org','l.kavoo@cifor-icraf.org')
# $reportRecipients = @('servicedesk@cifor-icraf.org','c.mwangi@cifor-icraf.org','p.oyuko@cifor-icraf.org','l.kavoo@cifor-icraf.org')
$reportRecipients = @('b.obaga@cifor-icraf.org')


$reportMonth = Get-Date -Format "MMMMyyyy"
#Define Other Report Directory
$reportDirectoryKaseya = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Sources\Kaseya\" + $reportMonth + "\"

#Define Other Report Filename
$kaseyaCountfn = "Machine Count - ICRAFHQ.xlsx"
$kaseyaCountfnR = "Machine Count - ICRAF Regions.xlsx"
$KaseyaCountODfn = "Machine Count (OneDrive Installed) - ICRAFHQ.xlsx"
$KaseyaCountODfnR = "Machine Count (OneDrive Installed) - ICRAF Regions.xlsx"
$KaseyaCountODMacfn = "Machine Count (OneDrive Installed - Mac) - ICRAFHQ.xlsx"
$KaseyaCountODMacfnR = "Machine Count (OneDrive Installed - Mac) - ICRAF Regions.xlsx"
$KaseyaCountODConfn = "Machine Count (OneDrive Configured) - ICRAFHQ.xlsx"
$KaseyaCountODConfnR = "Machine Count (OneDrive Configured) - ICRAF Regions.xlsx"

#Define Exported Report Directory
$reportDirectory = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Results\OneDrive\ICRAF\"
$reportDirectory = $reportDirectory + $reportDate + "\"
$compressDirectory = $reportDirectory + "*"
$compressedDirectory = $reportDirectory + "OneDriveReport.zip"

#Define Filter String
$skippedOU = @('Disabled accounts','ICRAF Administrators','ICRAF BOT','ICRAF ICT','ICRAF ICT Test','ICRAF Kenya General Accounts','ICRAF Meeting Rooms','ICRAF Kenya Folder Redirection','ICRAF Kenya Groups','ICRAF MFI')
$ODReport = Get-OneDriveReport

#Get Report Refresh Date
$refreshDate = $ODReport.ReportRefreshDate

#Compile Report
#Report From Kaseya
[string]$kaseyaCount = Get-KaseyaMachineCount -reportDirectory $reportDirectoryKaseya -fileName $kaseyaCountfn
[string]$kaseyaCountR = Get-KaseyaMachineCount -reportDirectory $reportDirectoryKaseya -fileName $kaseyaCountfnR
[string]$KaseyaCountWithOD = Get-KaseyaMachineCount -reportDirectory $reportDirectoryKaseya -fileName $KaseyaCountODfn
[string]$KaseyaCountWithODR = Get-KaseyaMachineCount -reportDirectory $reportDirectoryKaseya -fileName $KaseyaCountODfnR
[string]$KaseyaCountWithODMac = Get-KaseyaMachineCount -reportDirectory $reportDirectoryKaseya -fileName $KaseyaCountODMacfn
[string]$KaseyaCountWithODMacR = Get-KaseyaMachineCount -reportDirectory $reportDirectoryKaseya -fileName $KaseyaCountODMacfnR
[string]$KaseyaCountODConf = Get-KaseyaMachineCount -reportDirectory $reportDirectoryKaseya -fileName $KaseyaCountODConfn
[string]$KaseyaCountODConfR = Get-KaseyaMachineCount -reportDirectory $reportDirectoryKaseya -fileName $KaseyaCountODConfnR
[int]$KaseyaCountWithODTot = [int]$KaseyaCountWithOD + [int]$KaseyaCountWithODMac
[int]$KaseyaCountWithODTotR = [int]$KaseyaCountWithODR + [int]$KaseyaCountWithODMacR
[int]$KaseyaCountNoOD = [int]$kaseyaCount - [int]$KaseyaCountWithODTot
[int]$KaseyaCountNoODR = [int]$kaseyaCountR - [int]$KaseyaCountWithODTotR
[int]$KaseyaCountODNConf = [int]$KaseyaCountWithODTot - [int]$KaseyaCountODConf
[int]$KaseyaCountODNConfR = [int]$KaseyaCountWithODTotR - [int]$KaseyaCountODConfR

#Report From OneDrive
$ICRAFUsers = $ODReport.AllUsers | Where 'OU' -Match "ICRAF Kenya"
$ICRAFUsersCount = $ICRAFUsers.Count
$ICRAFUsersNoActivityDate = $ODReport.UsersNoActivityDate | Where 'OU' -Match 'ICRAF Kenya'
$ICRAFUsersNoActivityDateCount = $ICRAFUsersNoActivityDate.Count
$ICRAFUsersNoActivityDate_NotGeneral = $ICRAFUsersNoActivityDate | Where 'OU' -NotMatch ($skippedOU -join '|')
$ICRAFUsersNoActivityDate_NotGeneralCount = $ICRAFUsersNoActivityDate_NotGeneral.Count
$ICRAFUsersOldActivityDate = $ICRAFUsers | Where 'LastActivityDate' -lt ($refreshDate.AddDays(-15))
$ICRAFUsersOldActivityDateCount = $ICRAFUsersOldActivityDate.Count
$ICRAFUsersOldActivityDate_NotGeneral = $ICRAFUsersOldActivityDate | Where 'OU' -NotMatch ($skippedOU -join '|')
$ICRAFUsersOldActivityDate_NotGeneralCount = $ICRAFUsersOldActivityDate_NotGeneral.Count
$ICRAFUsersActiveAccounts = $ICRAFUsers | Where 'LastActivityDate' -GE ($refreshDate.AddDays(-15))
$ICRAFUsersActiveAccountsCount = $ICRAFUsersActiveAccounts.Count
$ICRAFUsersLowStorageUtilization = $ICRAFUsers | Where 'StorageUsageinGB' -lt 5
$ICRAFUsersLowStorageUtilization_NotGeneral = $ICRAFUsersLowStorageUtilization | Where 'OU' -NotMatch ($skippedOU -join '|')
$ICRAFUsersLowStorageUtilization_NotGeneralCount = $ICRAFUsersLowStorageUtilization_NotGeneral.Count
$ICRAFUsersGoodStorageUtilization = $ICRAFUsers | Where 'StorageUsageinGB' -GE 5
$ICRAFUsersGoodStorageUtilizationCount = $ICRAFUsersGoodStorageUtilization.Count

#ICRAF Regions 
$ICRAFRGUsers = $ODReport.AllUsers | Where 'OU' -Match "ICRAF Regions"
$ICRAFRGUsersCount = $ICRAFRGUsers.Count
$ICRAFRGUsersNoActivityDate = $ODReport.UsersNoActivityDate | Where 'OU' -Match 'ICRAF Regions'
$ICRAFRGUsersNoActivityDateCount = $ICRAFRGUsersNoActivityDate.Count
#$ICRAFRGUsersNoActivityDate_NotGeneral...commented for future use
#$ICRAFRGsersNoActivityDate_NotGeneralCount...commented for future use
$ICRAFRGUsersOldActivityDate = $ICRAFRGUsers | Where 'LastActivityDate' -lt ($refreshDate.AddDays(-15))
$ICRAFRGUsersOldActivityDateCount = $ICRAFRGUsersOldActivityDate.Count
#$ICRAFRGUsersOldActivityDate_NotGeneral...commented for future use
#$ICRAFRGUsersOldActivityDate_NotGeneralCount...commented for future use
$ICRAFRGUsersActiveAccounts = $ICRAFRGUsers | Where 'LastActivityDate' -GE ($refreshDate.AddDays(-15))
$ICRAFRGUsersActiveAccountsCount = $ICRAFRGUsersActiveAccounts.Count
$ICRAFRGUsersLowStorageUtilization = $ICRAFRGUsers | Where 'StorageUsageinGB' -lt 5
$ICRAFRGUsersLowStorageUtilizationCount = $ICRAFRGUsersLowStorageUtilization.Count
#$ICRAFRGUsersLowStorageUtilization_NotGeneral...commented for future use
#$ICRAFRGUsersLowStorageUtilization_NotGeneralCount...commented for future use
$ICRAFRGUsersGoodStorageUtilization = $ICRAFRGUsers | Where 'StorageUsageinGB' -GE 5
$ICRAFRGUsersGoodStorageUtilizationCount = $ICRAFRGUsersGoodStorageUtilization.Count


#Compile Report Body
$reportBody = @"
##### ICRAF OneDrive Report #####`r`n
##### Kaseya Report #####`r`n
No of managed computers (Kaseya) - HQ: $kaseyaCount`r`n
No of computers with OneDrive App Installed - HQ: $KaseyaCountWithODTot`r`n
No of computers without OneDrive App - HQ: $KaseyaCountNoOD`r`n
No of computers with OneDrive Configured - HQ: $KaseyaCountODConf`r`n
No of computers with OneDrive Not Configured - HQ: $KaseyaCountODNConf`r`n

##### General Report #####`r`n
Total Users Reported From CGNET: $($ODReport.TotalUsersCount)`r`n
Total Users marked as deleted from the CSV report (isDeleted = True): $($ODReport.DeletedUsers.Count)`r`n
Total active users from the report: $($ODReport.ActiveUsersCount)`r`n
Total active users that do not have principal name listed on the CSV report: $($ODReport.UsersNoPrimaryAddress.Count)`r`n
Total active users that do not have Last Activity Date on the CSV report: $($ODReport.UsersNoActivityDate.Count)`r`n
 
##### ICRAF Kenya Report #####`r`n
Total active users reported under ICRAF: $ICRAFUsersCount`r`n
Total active users that do not have Last Activity Date under ICRAF (Excluding General, Service, Disabled and Other user accounts): $ICRAFUsersNoActivityDate_NotGeneralCount`r`n
Total active users that have synchronized within the last 15 days under ICRAF: $ICRAFUsersActiveAccountsCount`r`n
Total active users that have not synchronized within the last 15 days under ICRAF (Excluding General, Service, Disabled and Other user accounts): $ICRAFUsersOldActivityDate_NotGeneralCount`r`n
Total active users that have storage utilization greater than 5GB: $ICRAFUsersGoodStorageUtilizationCount`r`n
Total active users that have storage utilization lower than 5GB under ICRAF (Excluding General, Service, Disabled and Other user accounts): $ICRAFUsersLowStorageUtilization_NotGeneralCount`r`n

#### ICRAF Regions Report ####`r`n
##### Kaseya Report #####`r`n
No of managed computers (Kaseya) - ICRAF Regions: $kaseyaCountR`r`n
No of computers with OneDrive App Installed - ICRAF Regions: $KaseyaCountWithODTotR`r`n
No of computers without OneDrive App - ICRAF Regions: $KaseyaCountNoODR`r`n
No of computers with OneDrive Configured - ICRAF Regions: $KaseyaCountODConfR`r`n
No of computers with OneDrive Not Configured - ICRAF Regions: $KaseyaCountODNConfR`r`n

Total active users reported under ICRAF Regions: $ICRAFRGUsersCount`r`n
Total active users that do not have Last Activity Date under ICRAF Regions: $ICRAFRGUsersNoActivityDateCount`r`n
Total active users that have synchronized within the last 15 days under ICRAF Regions: $ICRAFRGUsersActiveAccountsCount`r`n
Total active users that have not synchronized within the last 15 days under ICRAF Regions: $ICRAFRGUsersOldActivityDateCount`r`n
Total active users that have storage utilization greater than 5GB: $ICRAFRGUsersGoodStorageUtilizationCount`r`n
Total active users that have storage utilization lower than 5GB under ICRAF Regions: $ICRAFRGUsersLowStorageUtilizationCount`r`n
"@


#Export CSV Reports

mkdir $reportDirectory -Force
$ICRAFUsersLowStorageUtilization_NotGeneral | Export-Csv -Path ($reportDirectory + "ICRAFHQ-LowUtilization.csv") -NoTypeInformation
$ICRAFUsersNoActivityDate_NotGeneral | Export-Csv -Path ($reportDirectory + "ICRAFHQ-NoActivityDate.csv") -NoTypeInformation
$ICRAFUsersOldActivityDate_NotGeneral | Export-Csv -Path ($reportDirectory + "ICRAFHQ-OldActivityDate.csv") -NoTypeInformation
$ICRAFRGUsersLowStorageUtilization | Export-Csv -Path ($reportDirectory + "ICRAFRegions-LowUtilization.csv") -NoTypeInformation
$ICRAFRGUsersNoActivityDate | Export-Csv -Path ($reportDirectory + "ICRAFRegions-NoActivityDate.csv") -NoTypeInformation
$ICRAFRGUsersOldActivityDate | Export-Csv -Path ($reportDirectory + "ICRAFRegions-OldActivityDate.csv") -NoTypeInformation

#Compress The Directory
Compress-Archive -Path $compressDirectory -DestinationPath $compressedDirectory -CompressionLevel Optimal -Force

Send-OneDriveAlert z-UsersOldActivityDate_NotGeneral $ICRAFUsersOldActivityDate_NotGeneral -UsersLowStorageUtilization_NotGeneral $ICRAFUsersLowStorageUtilization_NotGeneral
Send-OneDriveReport -recipients $reportRecipients -reportBody $reportBody -reportAttach $compressedDirectory






