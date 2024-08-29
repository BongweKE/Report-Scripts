function Get-ESETReport{
param($reportDirectory,$lastConnection,$lastUpdate,$threatCount,$criticalMachines,$computerCount)
[hashtable]$esetReport = @{}

#Get Last Connection Details
$esetLastConnectionAbove30 = $reportDirectory + $lastConnection
$esetLastConnectionAbove30 = Import-Csv -Path $esetLastConnectionAbove30 -Delimiter ';'
# $esetLastConnectionAbove30_Count = $esetLastConnectionAbove30.'Count (Computer)' | Measure-Object -Sum | Select -ExpandProperty Sum
$esetLastConnectionAbove30Regions = $esetLastConnectionAbove30 | Where 'Group by (Static group name)' -ne 'ICRAF HQ' | Measure-Object -Sum 'Count (Computer)' | Select -ExpandProperty Sum 
$esetLastConnectionAbove30HQ = $esetLastConnectionAbove30 | Where 'Group by (Static group name)' -eq 'ICRAF HQ' | Measure-Object -Sum 'Count (Computer)' | Select -ExpandProperty Sum 

#Get Update Details
#########################################################################################
# Change request: Update Data so that we get update details in two grouped reports
#########################################################################################
$esetUpdateDetails = $reportDirectory + $lastUpdate
$esetUpdateDetails = Import-Csv -Path $esetUpdateDetails -Delimiter ';'
$esetUpdatedComputers_Count = $esetUpdateDetails | Where 'Variable interval list (Time of occurrence)' -ne '> 7 days' | Measure-Object -Sum 'Count (Computer)' | Select -ExpandProperty Sum
$esetNotUpdatedComputers_Count = $esetUpdateDetails | Where 'Variable interval list (Time of occurrence)' -eq '> 7 days' | Measure-Object -Sum 'Count (Computer)' | Select -ExpandProperty Sum


#Get Threat Count
$esetThreatDetails = $reportDirectory + $threatCount
$esetThreatDetails = Import-Csv $esetThreatDetails -Delimiter ';'
$esetThreat_CountRegions =  $esetThreatDetails | Where 'Group by (Static group name)' -ne 'ICRAF HQ' | Measure-Object -Sum 'Count (Threat name)' | Select -ExpandProperty Sum
$esetThreat_CountHQ =  $esetThreatDetails | Where 'Group by (Static group name)' -eq 'ICRAF HQ' | Measure-Object -Sum 'Count (Threat name)' | Select -ExpandProperty Sum

#Get High Severity Threat Details
$esetHighSeverityDetails = $reportDirectory + $criticalMachines
$esetHighSeverityDetails = Import-Csv -Path $esetHighSeverityDetails -Delimiter ';' 
$esetHighSeverity_Count = ($esetHighSeverityDetails | Select "Computer name" | Get-Unique -AsString).Count
$esetHighSeverity_Count

#Get Computer  (Old report)
$esetComputerCount = $reportDirectory + $computerCount
$esetComputerCount = Import-Csv $esetComputerCount -Delimiter ';'
$esetComputerCountfnR = $esetComputerCount | Where 'Group by (Static group name)' -ne 'ICRAF HQ' | Measure-Object -Sum 'Count (Computer)' | Select -ExpandProperty Sum 
$esetComputerCountHQ = $esetComputerCount | Where 'Group by (Static group name)' -eq 'ICRAF HQ' | Measure-Object -Sum 'Count (Computer)' | Select -ExpandProperty Sum


# Add to the reports Object
$esetReport.LastConnectionRegions = [int]$esetLastConnectionAbove30Regions
$esetReport.LastConnectionHQ = [int]$esetLastConnectionAbove30HQ
$esetReport.UpdatedComputers = [int]$esetUpdatedComputers_Count
$esetReport.NotUpdatedComputers = [int]$esetNotUpdatedComputers_Count
$esetReport.ThreatCountRegions = [int]$esetThreat_CountRegions
$esetReport.ThreatCountHQ = [int]$esetThreat_CountHQ
$esetReport.CriticalMachines = [int]$esetHighSeverity_Count
$esetReport.ComputerCountRegions = [int]$esetComputerCountfnR
$esetReport.ComputerCountHQ = [int]$esetComputerCountHQ

<#



$updatedComputerCount = $esetReport.UpdatedComputers
$notUpdatedComputerCount = $esetReport.NotUpdatedComputers
$criticalMachinesCount = $esetReport.CriticalMachines


#>
return $esetReport
}

Export-ModuleMember -Function *


