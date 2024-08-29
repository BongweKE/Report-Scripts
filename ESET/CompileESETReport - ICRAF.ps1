Import-Module "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Scripts\ESET\ESETReport.psm1" -Verbose -Force
Import-Module "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Scripts\OneDrive\GetKaseyaMachineCount.psm1" -Verbose -Force
Import-Module "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Scripts\OneDrive\Get-OneDriveReport.psm1"
Import-Module "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Scripts\ESET\SendESETReport.psm1" -Verbose -Force
#Import-Module "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Scripts\General\GetServiceNowMachineCount.psm1" -Verbose -Force

#Define OU & Location
$ou = "OU=Computers,OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG"
$location = "ICRAF HQ & Regions"

#Define Email Address of Recipients
$reportRecipients = @('b.obaga@cifor-icraf.org')

#Define Date
$reportDate = Get-Date
# $weekOfMonthNo = (Get-WmiObject Win32_LocalTime).WeekInMonth

$firstDayOfMonth = Get-Date -Day 1 -Month $reportDate.Month -Year $reportDate.Year
$firstDayOfWeekOfMonth = $firstDayOfMonth.DayOfWeek.value__

# Adjust for weeks starting on Sunday
if ($firstDayOfWeekOfMonth -eq 0) { $firstDayOfWeekOfMonth = 7 } 


#########################################################################################
# change back to auto calculate
#########################################################################################
$weekOfMonth = 1
#$weekOfMonth = [Math]::Ceiling(($date.Day + $firstDayOfWeekOfMonth - 1) / 7)

$reportSubFolder = Get-Date -Format "MMMMyyyy"

$reportDirectoryKaseya = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Sources\Kaseya\" + $reportSubFolder + "\"

$reportSubFolder = $reportSubFolder+"wk"+$weekOfMonth
#Define Directory Path
$reportDirectoryESET = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Sources\ESET\" + $reportSubFolder + "\"

$reportDirectoryServiceNow = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Sources\ServiceNow\"

#########################################################################################
# GET KASEYA DATA
#########################################################################################
# Define Various Machine Count Filenames

$kaseyaCountfn = "Machine Count - ICRAFHQ.xlsx"
$kaseyaCountfnR = "Machine Count - ICRAF Regions.xlsx"
#$serviceNowCountfn = "ICRAFHQ Machines.csv"

#########################################################################################
# Future: get Kandji Machine count reports using API
#########################################################################################

#Get Machine Counts
$kaseyaCount = Get-KaseyaMachineCount -reportDirectory $reportDirectoryKaseya -fileName $kaseyaCountfn

$kaseyaCount = $kaseyaCount[0]
$kaseyaCountR = Get-KaseyaMachineCount -reportDirectory $reportDirectoryKaseya -fileName $kaseyaCountfnR
$kaseyaCountR = $kaseyaCountR[0]
# $serviceNowCount = Get-ServiceNowMachineCount -reportDirectory $reportDirectoryServiceNow -fileName $serviceNowCountfn
# $ADCount = (Get-ADComputer -Filter * -SearchBase $ou).count

#########################################################################################
# ESET
#########################################################################################


#Define Grouped ESET Report Filenames
$lastConnectionfn = "ICRAF - Last Connected More than a week ago(Grouped By Office).csv"
$computerCountfn = "ICRAF - Computer Count(Grouped By Office).csv"
$lastUpdatefn = "ICRAF - Last Updated More than a week ago(Grouped By Office).csv"
$criticalMachinesfn = "ICRAF - Machine with Critical Threats (Grouped By Office).csv"
$threatCountfn = "ICRAF - Threat Count (Grouped by Office).csv"

#Get ESET Report
$esetReport = Get-ESETReport -reportDirectory $reportDirectoryESET -lastConnection $lastConnectionfn -lastUpdate $lastUpdatefn -threatCount $threatCountfn -criticalMachines $criticalMachinesfn -computerCount $computerCountfn


$lastConnectionCount = $esetReport.LastConnection
$computerCount = $esetReport.ComputerCount
$updatedComputerCount = $esetReport.UpdatedComputers
$notUpdatedComputerCount = $esetReport.NotUpdatedComputers
$criticalMachinesCount = $esetReport.CriticalMachines
$threatCountHQ = $esetReport.ThreatCountHQ
$threatCountRegions = $esetReport.ThreatCountRegions

#######################################################################################
# TBC:
# => Update scheduler to send group by data where necessary _/
# => Update ESETReport.psm1 to fetch group by data 
# => Remove unnecessary variables above
# => Finalize Excel Creation below
# => finalize sending report: Test first using your mail then add other
###############################################################################
$successfullConnection = $computerCount - $lastConnectionCount
$percentageSuccessfulConnection = ($successfullConnection/$computerCount).ToString("#.##%")
$percentageUpdatedComputers = ($updatedComputerCount/$computerCount).ToString("#.##%")
$percentageNotUpdatedComputers = ($notUpdatedComputerCount/$computerCount).ToString("#.##%")
$percentageComputersWithESET = ($computerCount/$kaseyaCount).ToString("#.##%")
$ComputersWithoutESET = [math]::Max(0,$kaseyaCount - $computerCount)
$percentageComputersWithoutESET = ($computersWithoutESET/$kaseyaCount).ToString("0.00%")
$percentageComputerWithKaseya = ($kaseyaCount/$ADCount).ToString("#.##%")
$computersWithoutKaseya = [math]::Max(0,$ADCount - $kaseyaCount)
$percentageComputerWithoutKaseya = ($computersWithoutKaseya/$ADCount).ToString("0.00%")

#Run Report Regions
$esetReportR = Get-ESETReport -reportDirectory $reportDirectoryESET -lastConnection $lastConnectionfnR -lastUpdate $lastUpdatefnR -threatCount $threatCountfnR -criticalMachines $criticalMachinesfnR -computerCount $computerCountfnR

#########################################################################################
# GET ONEDRIVE DATA
#########################################################################################

$ODReport = Get-OneDriveReport

#########################################################################################
# COMPILE ALL DATA
#########################################################################################
$lastConnectionCountR = $esetReportR.LastConnection
$computerCountR = $esetReportR.ComputerCount
$updatedComputerCountR = $esetReportR.UpdatedComputers
$notUpdatedComputerCountR = $esetReportR.NotUpdatedComputers
$criticalMachinesCountR = $esetReportR.CriticalMachines
$threatCountR = $esetReportR.ThreatCount
$successfullConnectionR = $computerCountR - $lastConnectionCountR
$percentageSuccessfulConnectionR = ($successfullConnectionR/$computerCountR).ToString("#.##%")
$percentageUpdatedComputersR = ($updatedComputerCountR/$computerCountR).ToString("#.##%")
$percentageNotUpdatedComputersR = ($notUpdatedComputerCountR/$computerCountR).ToString("#.##%")
$percentageComputersWithESETR = ($computerCountR/$kaseyaCountR).ToString("#.##%")
$ComputersWithoutESETR = [math]::Max(0,$kaseyaCountR - $computerCountR)
$percentageComputersWithoutESETR = ($computersWithoutESETR/$kaseyaCountR).ToString("0.00%")
$percentageComputerWithKaseyaR = ($kaseyaCount/$ADCount).ToString("#.##%")
$computersWithoutKaseyaR = [math]::Max(0,$ADCount - $kaseyaCount)
$percentageComputerWithoutKaseyaR = ($computersWithoutKaseya/$ADCount).ToString("0.00%")

##########################################################################################
# EXCEL Creation
# => Get the Data
# => Create Excel If not Exists else open
# => Map To columns
# => Close excel
#
#########################################################################################

<#
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

#>

#########################################################################################
# Email The report (occurs every fortnite as per Task Schduler)
#########################################################################################
# format date in standard format first
$reportDate = Get-Date -Format "MMMM dd yyyy"

#Compile Report Body
$reportBody = @"
Date of Report: $reportDate`r
Report For: $location`r
No of Computers on AD : $ADCount`r
No of Computers on Kaseya (HQ) : $KaseyaCount`r
No of Computers on Kaseya (Regions) : $KaseyaCountR`r
No of Computers on ServiceNow : $serviceNowCount`r
Client Count on AV Database : $computerCount`r
Client Count on AV Database (Regions) : $computerCountR`r
------------------------------------------------------
STATUS OF AV SIGNATURE DATABASE (HQ)
------------------------------------------------------
Computers Updated in the last 1 Week : $updatedComputerCount`r
Computers Not Updated in the last 1 Week : $notUpdatedComputerCount`r
------------------------------------------------------
STATUS OF AV SIGNATURE DATABASE (Regions)
------------------------------------------------------
Computers Updated in the last 1 Week : $updatedComputerCountR`r
Computers Not Updated in the last 1 Week : $notUpdatedComputerCountR`r
------------------------------------------------------
CONNECTION TO SERVER (HQ)
------------------------------------------------------
Computers Not Connected in over 1 Month : $lastConnectionCount`r
------------------------------------------------------
CONNECTION TO SERVER (Regions)
------------------------------------------------------
Computers Not Connected in over 1 Month : $lastConnectionCountR`r
------------------------------------------------------
THREAT LOG
------------------------------------------------------
Critical Machines : $criticalMachinesCount`r
No. of Threats : $threatCount`r
"@ -split '\n'

#Send The Report
Send-ESETReport -recipients $reportRecipients -reportBody $reportBody

<#
#Everything below this line is under testing
$body = @{
"Date" = $reportDate;
"Recorded_Date" = $weekOfMonthNo;
"No. of Computers on Kaseya" = [int]$kaseyaCount;
"No. of Computers on AD" = $ADCount;
"Client count on AV" = $computerCount;
"No. of Updated Clients in the last 1 week" = $updatedComputerCount;
"No. of Not Updated Clients in the last 1 week" = $notUpdatedComputerCount;
"No of Clients Not Connected Successfully within a month" = $lastConnectionCount;
"No of Clients Connected Successfully within a month" = $successfullConnection;
"Percentage of Computers Connected" = $percentageSuccessfulConnection;
"Percentage of Computers Updated" = $percentageUpdatedComputers;
"Percentage of Computers Not Updated" = $percentageNotUpdatedComputers;
"Percentage of Computers With ESET" = $percentageComputersWithESET;
"Percentage of Computers Without ESET" = $percentageComputersWithoutESET;
"Percentage of Computers With Kaseya" = $percentageComputerWithKaseya;
"Percentage of Computers Without Kaseya" = $percentageComputerWithoutKaseya;
"No. of Computers on Kaseya Regions" = [int]$kaseyaCountR;
"Client count on AV Regions" = $computerCountR;
"No. of Updated Clients in the last 1 week Regions" = $updatedComputerCountR;
"No. of Not Updated Clients in the last 1 week Regions" = $notUpdatedComputerCountR;
"No of Clients Not Connected Successfully within a month Regions" = $lastConnectionCountR;
"No of Clients Connected Successfully within a month Regions" = $successfullConnectionR;
"Percentage of Computers Connected Regions" = $percentageSuccessfulConnectionR;
"Percentage of Computers Updated Regions" = $percentageUpdatedComputersR;
"Percentage of Computers Not Updated Regions" = $percentageNotUpdatedComputersR;
"Percentage of Computers With ESET Regions" = $percentageComputersWithESETR;
"Percentage of Computers Without ESET Regions" = $percentageComputersWithoutESETR;
"Percentage of Computers With Kaseya Regions" = $percentageComputerWithKaseyaR;
"Percentage of Computers Without Kaseya Regions" = $percentageComputerWithoutKaseyaR;
}

$uri = "https://prod-113.westeurope.logic.azure.com:443/workflows/45aee51d72564a3d822e178e1f6f12f7/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=iRUvtXodOyQt894N6MW5N95x81VaGdqukBPjDgdO6o4"

Invoke-RestMethod -Uri $uri -Method Post -Body ($body | ConvertTo-Json) -ContentType "application/json"
#>