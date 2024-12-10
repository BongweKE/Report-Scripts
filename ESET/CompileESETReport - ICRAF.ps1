Import-Module "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Scripts\ESET\ESETReport.psm1" -Verbose -Force
Import-Module "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Scripts\OneDrive\GetKaseyaMachineCount.psm1" -Verbose -Force
Import-Module "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Scripts\OneDrive\Get-OneDriveReport.psm1"
Import-Module "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Scripts\ESET\SendESETReport.psm1" -Verbose -Force
#Import-Module "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Scripts\General\GetServiceNowMachineCount.psm1" -Verbose -Force

#Define OU & Location
$ou = "OU=Computers,OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG"
$location = "ICRAF HQ & Regions"

#Define Email Address of Recipients
# $reportRecipients = @('r.kande@cifor-icraf.org','servicedesk@cifor-icraf.org','c.mwangi@cifor-icraf.org','p.oyuko@cifor-icraf.org','l.kavoo@cifor-icraf.org', 'b.obaga@cifor-icraf.org', 'g.kirimi@cifor-icraf.org')
# $reportRecipients = @('servicedesk@cifor-icraf.org','c.mwangi@cifor-icraf.org','p.oyuko@cifor-icraf.org','l.kavoo@cifor-icraf.org')
$reportRecipients = @('l.kavoo@cifor-icraf.org', 'b.obaga@cifor-icraf.org', 'g.kirimi@cifor-icraf.org')
# $reportRecipients = @('b.obaga@cifor-icraf.org')

#Define Date
$reportDate = Get-Date
# $weekOfMonthNo = (Get-WmiObject Win32_LocalTime).WeekInMonth

$firstDayOfMonth = Get-Date -Day 1 -Month $reportDate.Month -Year $reportDate.Year
$firstDayOfWeekOfMonth = $firstDayOfMonth.DayOfWeek.value__

# Adjust for weeks starting on Sunday
if ($firstDayOfWeekOfMonth -eq 0) { $firstDayOfWeekOfMonth = 7 } 


#########################################################################################
# change back to auto calculate
# $weekOfMonth = 1
#########################################################################################

$weekOfMonth = [Math]::Ceiling(($date.Day + $firstDayOfWeekOfMonth - 1) / 7)

$reportSubFolder = (Get-Date).AddDays(-1).ToString("MMMMyyyy")

$reportDirectoryKaseya = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Sources\Kaseya\" + $reportSubFolder + "\"


# useful for some of the other services' reports like kaseya
$reportSubFolder = $reportSubFolder+"wk"+$weekOfMonth

#Define Directory Path
$reportDirectoryESET = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Sources\ESET"

$reportDirectoryESET = Get-ChildItem -Path $reportDirectoryESET -Directory | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1

$newreportDirESET = $reportDirectoryESET.FullName+"wk"+$weekOfMonth

if (-Not (Test-Path -Path $newreportDirESET)){
    New-Item -Path $newreportDirESET -ItemType Directory
}
$temp=$reportDirectoryESET.FullName+"\*"
$reportDirectoryESET = $newreportDirESET+"\"

# Copy items to a different  and make a record of it so we can differentiate reports we send on email
Copy-Item -Path $temp -Destination $reportDirectoryESET

# Currently we aren't getting Service now reports
# $reportDirectoryServiceNow = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Sources\ServiceNow\"
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
$kaseyaCountHQ = $kaseyaCount[0]
$kaseyaCountR = Get-KaseyaMachineCount -reportDirectory $reportDirectoryKaseya -fileName $kaseyaCountfnR
$kaseyaCountR = $kaseyaCountR[0]
# $serviceNowCount = Get-ServiceNowMachineCount -reportDirectory $reportDirectoryServiceNow -fileName $serviceNowCountfn
# $ADCount = (Get-ADComputer -Filter * -SearchBase $ou).count 

#########################################################################################
# ESET
#########################################################################################
#Define Grouped ESET Report Filenames
<# 
The filenames supplied below are from the PowerAutomate flow in charge of ESET data

Each CSV is divided into office data for the data point in question 
#>
$lastConnectionfn = "ICRAF - Last Connected More than a week ago(Grouped By Office).csv"
$computerCountfn = "ICRAF - Computer Count(Grouped By Office).csv"
$lastUpdatefn = "ICRAF - Last Updated More than a week ago(Grouped By Office).csv"
$criticalMachinesfn = "ICRAF - Machine with Critical Threats (Grouped By Office).csv"
$threatCountfn = "ICRAF - Threat Count (Grouped by Office).csv"
$scannedMachinesfn = "ICRAF - Computers Scanned This Month (Grouped By OU).csv"

$output = Get-ESETReport -reportDirectory $reportDirectoryESET -lastConnection $lastConnectionfn -lastUpdate $lastUpdatefn -threatCount $threatCountfn -criticalMachines $criticalMachinesfn -computerCount $computerCountfn -scannedMachines $scannedMachinesfn
$esetReport = $output.esetReport
$uniqueHighSeverityDetails = $output.uniqueHighSeverityDetails

$uniqueHighSeverityDetailsFp = $reportDirectoryESET + "uniqueHighSeverityDetails.xlsx"
# Export the data to an Excel file
$uniqueHighSeverityDetails | Export-Excel -Path $uniqueHighSeverityDetailsFp -WorkSheetname "High Severity" -AutoSize

# Get ESET Report
<#
#---------------#
1. Get ESET Report with Last update using Unique Count /
2. Change scripts to ensure new data reaches the excel file
3. Change excel file to include scannedMachines

# Export the unique entries to a new CSV file
$uniqueHighSeverityDetails | Export-Csv -Path $outputCsvPath -Delimiter ';' -NoTypeInformation -Encoding UTF8

#>

$lastConnectionCountR = $esetReport.LastConnectionRegions
$lastConnectionCountHQ = $esetReport.LastConnectionHQ

$computerCountR = $esetReport.ComputerCountRegions
$computerCountHQ = $esetReport.ComputerCountHQ


$updatedComputerCountR = $esetReport.UpdatedComputersRegions
$updatedComputersHQ = $esetReport.UpdatedComputersHQ


$notUpdatedComputerCountR = $computerCountR - $updatedComputerCountR
$notUpdatedComputerCountHQ = $computerCountHQ - $notUpdatedComputerCountHQ

$criticalMachinesCount = $esetReport.CriticalMachines
$criticalMachinesCountHQ = $esetReport.HighSeverity_CountHQ
$criticalMachinesCountR = $esetReport.HighSeverity_CountR

$threatCountHQ = $esetReport.ThreatCountHQ
$threatCountRegions = $esetReport.ThreatCountRegions

$scannedMachinesRegions = $esetReport.ScannedMachinesRegions
$scannedMachinesHQ = $esetReport.ScannedMachinesHQ
#######################################################################################
# TBC:
# => Update scheduler to send group by data where necessary _/
# => Update ESETReport.psm1 to fetch group by data  _/
# => Remove unnecessary variables above _/
# => Finalize Excel Creation below _/
# => finalize sending report: Test first using your mail then add other _/
###############################################################################


<#


--------------------------------
Recheck Kaseya Data and continue for regions

Then Test
----------------------------------

#>


$successfullConnectionHQ = $computerCountHQ - $lastConnectionCountHQ
$percentageSuccessfulConnectionHQ = ($successfullConnectionHQ/$computerCountHQ).ToString("#.##%")
$percentageUpdatedComputersHQ = ($updatedComputersHQ/$computerCountHQ).ToString("#.##%")
$percentageNotUpdatedComputersHQ = ($notUpdatedComputerCountHQ/$computerCountHQ).ToString("#.##%")
$percentageComputersWithESETHQ = ($computerCountHQ/$kaseyaCountHQ).ToString("#.##%")
$ComputersWithoutESETHQ = [math]::Max(0,$kaseyaCountHQ - $computerCountHQ)
$percentageComputersWithoutESETHQ = ($computersWithoutESETHQ/$kaseyaCountHQ).ToString("0.00%")
<#
$lastConnectionCountR
$computerCountR
$updatedComputerCountR
$notUpdatedComputerCountR
$threatCountRegions
$scannedMachinesRegions
#>

$successfullConnectionR = $computerCountR - $lastConnectionCountR
$percentageSuccessfulConnectionR = ($successfullConnectionR/$computerCountR).ToString("#.##%")
$percentageUpdatedComputersR = ($updatedComputerCountR/$computerCountR).ToString("#.##%")
$percentageNotUpdatedComputersR = ($notUpdatedComputerCountR/$computerCountR).ToString("#.##%")
$percentageComputersWithESETR = ($computerCountR/$kaseyaCountR).ToString("#.##%")
$ComputersWithoutESETR = [math]::Max(0,$kaseyaCountR - $computerCountR)
$percentageComputersWithoutESETR = ($computersWithoutESETR/$kaseyaCountR).ToString("0.00%")


$threatCount = $threatCountHQ + $threatCountRegions
#########################################################################################
# excel Report Creation Directory
#########################################################################################

$excelReport = "c:\Users\poadmin\CIFOR-ICRAF\Information Communication Technology (ICT) - Reports Archive\ESET Reports\eset_dashboard_data.xlsx"
# Create a new Excel application object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# Open the Excel file
$workbook = $excel.Workbooks.Open($excelReport)
# Access the worksheets (replace Sheet1 and Sheet2 with actual sheet names)
$sheetHQ = $workbook.Worksheets.Item("HQ")
$sheetR = $workbook.Worksheets.Item("Regions")

# Find the next available row in each sheet
$nextRowHQ = $sheetHQ.UsedRange.Rows.Count + 1
$nextRowR = $sheetR.UsedRange.Rows.Count + 1


# Populate data for HQ sheet
$sheetHQ.Cells.Item($nextRowHQ, 1) = (Get-Date).ToString("MMMM dd yyyy")  # Assuming 'Date' is the first column
$sheetHQ.Cells.Item($nextRowHQ, 2) = $computerCountHQ
$sheetHQ.Cells.Item($nextRowHQ, 3) = $computersWithoutESETHQ
$sheetHQ.Cells.Item($nextRowHQ, 4) = $percentageComputersWithoutESETHQ
$sheetHQ.Cells.Item($nextRowHQ, 5) = $successfullConnectionHQ
$sheetHQ.Cells.Item($nextRowHQ, 6) = $percentageSuccessfulConnectionHQ
$sheetHQ.Cells.Item($nextRowHQ, 7) = $updatedComputerCount
$sheetHQ.Cells.Item($nextRowHQ, 8) = $percentageUpdatedComputersHQ
$sheetHQ.Cells.Item($nextRowHQ, 9) = $notUpdatedComputerCount
$sheetHQ.Cells.Item($nextRowHQ, 10) = $percentageNotUpdatedComputersHQ
$sheetHQ.Cells.Item($nextRowHQ, 11) = $criticalMachinesCountHQ
$sheetHQ.Cells.Item($nextRowHQ, 12) = $scannedMachinesHQ

# $criticalMachinesCountHQ



# Populate data for Region sheet
$sheetR.Cells.Item($nextRowR, 1) = (Get-Date).ToString("MMMM dd yyyy")
$sheetR.Cells.Item($nextRowR, 2) = $computerCountR
$sheetR.Cells.Item($nextRowR, 3) = $ComputersWithoutESETR
$sheetR.Cells.Item($nextRowR, 4) = $percentageComputersWithoutESETR
$sheetR.Cells.Item($nextRowR, 5) = $successfullConnectionR
$sheetR.Cells.Item($nextRowR, 6) = $percentageSuccessfulConnectionR
$sheetR.Cells.Item($nextRowR, 7) = $updatedComputerCountR
$sheetR.Cells.Item($nextRowR, 8) = $percentageUpdatedComputersR
$sheetR.Cells.Item($nextRowR, 9) = $notUpdatedComputerCountR
$sheetR.Cells.Item($nextRowR, 10) = $percentageNotUpdatedComputersR
$sheetR.Cells.Item($nextRowR, 11) = $criticalMachinesCountR
$sheetR.Cells.Item($nextRowR, 12) = $scannedMachinesRegions

<#
$scannedMachinesRegions

#>

# Save and close the workbook
$workbook.Save()
$workbook.Close()
$excel.Quit()

# Clean up COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheetHQ)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheetR)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
#########################################################################################
# Email The report (occurs monthly as per Task Schduler)
#########################################################################################
$reportDate = Get-Date -Format "MMMM dd yyyy"

# Compile Executive Summary
$summary = @"
**Executive Summary**

* As of $reportDate, $percentageUpdatedComputersHQ of computers at HQ and $percentageUpdatedComputersR of computers in the regions have updated antivirus signatures within the last week.
* $percentageSuccessfulConnectionHQ of HQ computers and $percentageSuccessfulConnectionR of regional computers have successfully connected to the server.
* $percentageComputersWithoutESETHQ of HQ computers and $percentageComputersWithoutESETR of regional computers are not on the ESET antivirus.
* There are currently $criticalMachinesCount critical machines with a total of $threatCount threats detected.
* 

Now, Here's A detailed look into these numbers:
"@

# Compile Detailed Report Body 
$reportBody = @"
## Date of Report: $reportDate
## Report For: $location

**System Counts**

* No of Computers on Kaseya (HQ): $kaseyaCountHQ
* No of Computers on Kaseya (Regions): $kaseyaCountR
* Client Count on AV Database (HQ): $computerCountHQ
* Client Count on AV Database (Regions): $computerCountR

**STATUS OF AV SIGNATURE DATABASE (HQ)**

* Computers Updated in the last 1 Week: $updatedComputerCount
* Computers Not Updated in the last 1 Week: $notUpdatedComputerCount
* Percentage of Updated Computers: $percentageUpdatedComputersHQ
* Percentage of Not Updated Computers: $percentageNotUpdatedComputersHQ

**STATUS OF AV SIGNATURE DATABASE (Regions)**

* Computers Updated in the last 1 Week: $updatedComputerCountR
* Computers Not Updated in the last 1 Week: $notUpdatedComputerCountR
* Percentage of Updated Computers: $percentageUpdatedComputersR
* Percentage of Not Updated Computers: $percentageNotUpdatedComputersR

**CONNECTION TO SERVER (HQ)**

* Computers Not Connected in over 1 Month: $lastConnectionCountHQ
* Successful Connections: $successfullConnectionHQ 
* Percentage of Successful Connections: $percentageSuccessfulConnectionHQ

**CONNECTION TO SERVER (Regions)**

* Computers Not Connected in over 1 Month: $lastConnectionCountR
* Successful Connections: $successfullConnectionR
* Percentage of Successful Connections: $percentageSuccessfulConnectionR

**ESET COVERAGE**

* Computers With ESET (HQ): $computerCountHQ
* Computers Without ESET (HQ): $ComputersWithoutESETHQ
* Percentage of Computers With ESET (HQ): $percentageComputersWithESETHQ
* Percentage of Computers Without ESET (HQ): $percentageComputersWithoutESETHQ

* Computers With ESET (Regions): $computerCountR
* Computers Without ESET (Regions): $ComputersWithoutESETR
* Percentage of Computers With ESET (Regions): $percentageComputersWithESETR
* Percentage of Computers Without ESET (Regions): $percentageComputersWithoutESETR

* Machines Scanned this month (HQ): $scannedMachinesHQ
* Machines Scanned this month (Regions): $scannedMachinesRegions

**THREAT LOG**

* Critical Machines: $criticalMachinesCount
* No. of Threats: $threatCount
* No. of Threats(HQ): $threatCountHQ
* No. of Threats(Regions): $threatCountRegions

Please find attached a detailed list of machines with critical threats for investigation purposes.
"@

# Combine Summary and Detailed Report
$fullReport = $summary + $reportBody

# Send The Report
Send-ESETReport -recipients $reportRecipients -reportBody $fullReport -attachmentFp $uniqueHighSeverityDetailsFp

<#
#Everything below this line is under testing
$body = @{
"Date" = $reportDate;
"Recorded_Date" = $weekOfMonthNo;
"No. of Computers on Kaseya" = [int]$kaseyaCount;

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