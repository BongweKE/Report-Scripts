function Get-ESETReport{
param($reportDirectory,$lastConnection,$lastUpdate,$threatCount,$criticalMachines,$computerCount,$scannedMachines)
[hashtable]$esetReport = @{}

#Get Last Connection Details
$esetLastConnectionAbove30 = $reportDirectory + $lastConnection
$esetLastConnectionAbove30 = Import-Csv -Path $esetLastConnectionAbove30 -Delimiter ';'
# $esetLastConnectionAbove30_Count = $esetLastConnectionAbove30.'Count (Computer)' | Measure-Object -Sum | Select-Object -ExpandProperty Sum
$esetLastConnectionAbove30Regions = $esetLastConnectionAbove30 | Where-Object 'Group by (Static group name)' -ne 'ICRAF HQ' | Measure-Object -Sum 'Count (Computer)' | Select-Object -ExpandProperty Sum 
$esetLastConnectionAbove30HQ = $esetLastConnectionAbove30 | Where-Object 'Group by (Static group name)' -eq 'ICRAF HQ' | Measure-Object -Sum 'Count (Computer)' | Select-Object -ExpandProperty Sum 

#Get Threat Count
$esetThreatDetails = $reportDirectory + $threatCount
$esetThreatDetails = Import-Csv $esetThreatDetails -Delimiter ';'
$esetThreat_CountRegions =  $esetThreatDetails | Where-Object 'Group by (Static group name)' -ne 'ICRAF HQ' | Measure-Object -Sum 'Count (Threat name)' | Select-Object -ExpandProperty Sum
$esetThreat_CountHQ =  $esetThreatDetails | Where-Object 'Group by (Static group name)' -eq 'ICRAF HQ' | Measure-Object -Sum 'Count (Threat name)' | Select-Object -ExpandProperty Sum

#Get High Severity Threat Details
$esetHighSeverityDetails = $reportDirectory + $criticalMachines
$esetHighSeverityDetails = Import-Csv -Path $esetHighSeverityDetails -Delimiter ';'
$esetHighSeverity_Count = ($esetHighSeverityDetails | Select-Object "Computer name" | Get-Unique -AsString).Count
$uniqueHighSeverityDetails = $esetHighSeverityDetails | Select-Object -Property "Static group name", "Computer name", "Threat type", "Threat name", "Severity" -Unique

$esetHighSeverity_CountHQ = ($uniqueHighSeverityDetails | Where-Object { $_."Static group name" -eq "ICRAF HQ" }).count
$esetHighSeverity_CountR = ($uniqueHighSeverityDetails | Where-Object { $_."Static group name" -ne "ICRAF HQ" }).count

<#
# Get unique entries based on the specified columns
$uniqueHighSeverityDetails = $esetHighSeverityDetails | Select-Object -Property "Static group name", "Computer name", "Threat type", "Threat name", "Severity" -Unique

# Export the unique entries to a new CSV file
$uniqueHighSeverityDetails | Export-Csv -Path $outputCsvPath -Delimiter ';' -NoTypeInformation -Encoding UTF8
#>

#Get Computer  (Old report)
$esetComputerCount = $reportDirectory + $computerCount
$esetComputerCount = Import-Csv $esetComputerCount -Delimiter ';'
$esetComputerCountfnR = $esetComputerCount | Where-Object 'Group by (Static group name)' -ne 'ICRAF HQ' | Measure-Object -Sum 'Count (Computer)' | Select-Object -ExpandProperty Sum 
$esetComputerCountHQ = $esetComputerCount | Where-Object 'Group by (Static group name)' -eq 'ICRAF HQ' | Measure-Object -Sum 'Count (Computer)' | Select-Object -ExpandProperty Sum

#Get Update Details: Count of updated machines
#########################################################################################
# Change request: Update Data so that we get update details in two grouped reports
#########################################################################################
$esetUpdateDetails = $reportDirectory + $lastUpdate
$esetUpdateDetails = Import-Csv -Path $esetUpdateDetails -Delimiter ';'
$esetUpdatedComputers_CountHQ = $esetUpdateDetails | Where-Object 'Group by (Static group name)' -eq 'ICRAF HQ' | Measure-Object -Sum 'Unique count (Computer)' | Select-Object -ExpandProperty Sum
# Regions Computer Counts
$esetUpdatedComputers_CountR = $esetUpdateDetails | Where-Object 'Group by (Static group name)' -ne 'ICRAF HQ' | Measure-Object -Sum 'Unique count (Computer)' | Select-Object -ExpandProperty Sum
$esetNotUpdatedComputers_CountR = $esetComputerCountfnR - $esetUpdatedComputers_CountR

# Get Machines that have been scanned
$esetScannedMachines = $reportDirectory + $scannedMachines
$esetScannedMachines = Import-Csv $esetScannedMachines -Delimiter ';'
$esetScannedMachinesfnR = $esetScannedMachines | Where-Object 'Group by (Static group name)' -ne 'ICRAF HQ' | Measure-Object -Sum 'Unique count (Computer)' | Select-Object -ExpandProperty Sum 
$esetScannedMachinesHQ = $esetScannedMachines | Where-Object 'Group by (Static group name)' -eq 'ICRAF HQ' | Measure-Object -Sum 'Unique count (Computer)' | Select-Object -ExpandProperty Sum

# Add to the reports Object
$esetReport.LastConnectionRegions = [int]$esetLastConnectionAbove30Regions
$esetReport.LastConnectionHQ = [int]$esetLastConnectionAbove30HQ

$esetReport.UpdatedComputersRegions = [int]$esetNotUpdatedComputers_CountR
$esetReport.UpdatedComputersHQ = [int]$esetUpdatedComputers_CountHQ

$esetReport.ThreatCountRegions = [int]$esetThreat_CountRegions
$esetReport.ThreatCountHQ = [int]$esetThreat_CountHQ

$esetReport.CriticalMachines = [int]$esetHighSeverity_Count

$esetReport.HighSeverity_CountHQ = [int]$esetHighSeverity_CountHQ
$esetReport.HighSeverity_CountR = [int]$esetHighSeverity_CountR

$esetReport.ComputerCountRegions = [int]$esetComputerCountfnR
$esetReport.ComputerCountHQ = [int]$esetComputerCountHQ

$esetReport.ScannedMachinesRegions = [int]$esetScannedMachinesfnR
$esetReport.ScannedMachinesHQ = [int]$esetScannedMachinesHQ


<#

$esetHighSeverity_CountHQ

$esetHighSeverity_CountR

    $updatedComputerCount = $esetReport.UpdatedComputers
    $notUpdatedComputerCount = $esetReport.NotUpdatedComputers
    $criticalMachinesCount = $esetReport.CriticalMachines
#>
return [PSCustomObject]@{
    esetReport = $esetReport
    uniqueHighSeverityDetails = $uniqueHighSeverityDetails
}
}

Export-ModuleMember -Function *


