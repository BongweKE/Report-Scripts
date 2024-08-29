function Get-ServiceNowMachineCount {
param($reportDirectory,$fileName)

$getMachineCount = Import-Csv -Path ($reportDirectory + $fileName) 
$getMachineCount = $getMachineCount.Count

return $getMachineCount
}

Export-ModuleMember -Function *