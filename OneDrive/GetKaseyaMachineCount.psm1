function Get-KaseyaMachineCount {
param($reportDirectory,$fileName)

$getMachineCount = Import-Excel -Path ($reportDirectory + $fileName) -StartRow 2
$getMachineCount = $getMachineCount | Select "Agents Count" -ExpandProperty "Agents Count"

return $getMachineCount
}

Export-ModuleMember -Function *