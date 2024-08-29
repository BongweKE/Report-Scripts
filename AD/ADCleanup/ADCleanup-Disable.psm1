function Account-Disabling
{
    param ($reportDirectory)
    $test_cleanup_list = 'C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Results\AD\ICRAF\February 2024\Test Cleanup List.csv'
    
    $year = Get-Date -Format 'yyyy'
    $month = Get-Date -Format 'MMM'

    Import-CSV $reportDirectory | 
    Foreach {
        Get-ADUser -Identity $_.SamAccountName | 
        Move-ADObject -TargetPath "OU=Mar,OU=2024,OU=Disabled accounts,OU=ICRAF Kenya,OU=ICRAFHUB,DC=cifor-icraf,DC=org" -PassThru | 
        Disable-ADAccount
    }
}
Export-ModuleMember -Function 'Account-Disabling'