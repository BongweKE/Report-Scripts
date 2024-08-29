function Get-OneDriveReport{
param($ODReport)

$StorageUsageinGB = 0

#Deleted Users Variable
$usersDeleted = @()
$usersDeletedCount = 0

#Variable for holding user accounts that were not found on AD when resolving by Principal Name
$userAccountsNotFoundOnAD = @()
$userAccountsNotFoundOnADCount = 0

#Variable for holding total users from the CSV list
$usersReport = @()
$usersCount = 0

#Variable for holding user accounts that were not found in any of the OUs
$usersNoOUReport = @()
$usersNoOUCount = 0

#Variable for holding user accounts without Primary Address from the CSV list
$usersNoPrimaryAddress = @()
$usersNoPrimaryAddressCount = 0

#Variable for holding user accounts without "Last Activity Date" from the CSV list
$usersNoActivityDate = @()
$usersNoActivityDateCount = 0

#Variable for holding reports for ICRAF
$ICRAFUsersICRAF = @()
$icrafUsersCount = 0

#Variable for holding all reports in a hashtable
[hashtable]$oneDriveReportAll = @{}

#Get Report Directory
$ReportDate = (Get-Date).AddDays(-1).ToString("yyyy-MM-dd")
$reportDirectory = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Sources\OneDrive\CIFOR-ICRAF OneDrive Bi-Weekly Usage Report "+ $ReportDate +".csv"

#Get Report File
$csvReport = Import-Csv -Path $reportDirectory
$UPN = 'Owner Principal Name' ###changed from 'PrincipalName' to 'Owner Principal Name' as per the CGNET changes to email address attribute

#Get Report Date From File 
[DateTime]$reportRefreshDate = $csvReport[0].'Report Refresh Date'

$totalUsers = $csvReport.count

foreach ($Users in $csvReport)
{
    $userEmail = $Users.$UPN
    #$userEmail
    if ($Users.'Is Deleted' -eq "True")
    {
        if ($Users.'Last Activity Date' -ne "") {
            try {
                $lastActDate = [datetime]::ParseExact($Users.'Last Activity Date','M/dd/yyyy',[cultureinfo]::CurrentCulture)
            } catch {
                $lastActDate = [datetime]::ParseExact($Users.'Last Activity Date','M/d/yyyy',[cultureinfo]::CurrentCulture)
            }
        }
        else {
            $lastActDate = 'No Activity Date'
        }
        $userData = New-Object psobject -Property @{
            DisplayName = $Users.'Owner Display Name'
            Email = $Users.$UPN
            OU = 'N/A'
            isAccountDeleted = $Users.'Is Deleted'
            LastActivityDate = $lastActDate            #[datetime]::ParseExact($Users.'Last Activity Date','MM/dd/yyyy',[cultureinfo]::CurrentCulture)
            ActiveFileCount = $Users.'Active File Count'
            StorageUsage = $Users.'Storage Used (Byte)'
            StorageUsageinGB = $Users.'Storage Used (Byte)'/1073741824  ### Changed that from 1000000000 to 1073741824 for accuracy since the conversion factor of bytes is 1024 and not 1000
            }
        
            $usersDeleted += $userData
            $usersDeletedCount++
    }
    elseIf ($Users.$UPN -eq "")
    {
        $userData = New-Object psobject -Property @{
            DisplayName = $Users.'Owner Display Name'
            Email = 'Prinicipal Name not listed'
            OU = 'N/A, Prinicipal Name not listed'
            isAccountDeleted = $Users.'Is Deleted'
            LastActivityDate = 'No Activity Date'            #[datetime]::ParseExact($Users.'Last Activity Date','MM/dd/yyyy',[cultureinfo]::CurrentCulture)
            ActiveFileCount = $Users.'Active File Count'
            StorageUsage = $Users.'Storage Used (Byte)'
            StorageUsageinGB = $Users.'Storage Used (Byte)'/1073741824  ### Changed that from 1000000000 to 1073741824 for accuracy since the conversion factor of bytes is 1024 and not 1000
        }
            $usersNoPrimaryAddress += $userData
            $usersNoPrimaryAddressCount++
        
    }
    else
    {
        if($Users.'Last Activity Date' -ne "")
            {
           try {
                $lastActDate = [datetime]::ParseExact($Users.'Last Activity Date','M/dd/yyyy',[cultureinfo]::CurrentCulture)
            } catch {
                $lastActDate = [datetime]::ParseExact($Users.'Last Activity Date','M/d/yyyy',[cultureinfo]::CurrentCulture)
            }
            $userOU = Get-ADUser -Filter {userPrincipalName -eq $userEmail} | Select-Object -ExpandProperty DistinguishedName   
            $userData = New-Object psobject -Property @{
                DisplayName = $Users.'Owner Display Name'
                Email = $Users.$UPN
                OU = $userOU
                isAccountDeleted = $Users.'Is Deleted'
                LastActivityDate = $lastActDate            #[datetime]::ParseExact($Users.'Last Activity Date','MM/dd/yyyy',[cultureinfo]::CurrentCulture)
                ActiveFileCount = $Users.'Active File Count'
                StorageUsage = $Users.'Storage Used (Byte)'
                StorageUsageinGB = $Users.'Storage Used (Byte)'/1073741824  ### Changed that from 1000000000 to 1073741824 for accuracy since the conversion factor of bytes is 1024 and not 1000
                }
                $usersReport += $userData
                $usersCount++
    
            }
        else {
            $userOU = Get-ADUser -Filter {userPrincipalName -eq $userEmail} | Select-Object -ExpandProperty DistinguishedName
            $userData = New-Object psobject -Property @{
                DisplayName = $Users.'Owner Display Name'
                Email = $Users.$UPN
                OU = $userOU
                isAccountDeleted = $Users.'Is Deleted'
                LastActivityDate = 'No Activity Date'            #[datetime]::ParseExact($Users.'Last Activity Date','MM/dd/yyyy',[cultureinfo]::CurrentCulture)
                ActiveFileCount = $Users.'Active File Count'
                StorageUsage = $Users.'Storage Used (Byte)'
                StorageUsageinGB = $Users.'Storage Used (Byte)'/1073741824  ### Changed that from 1000000000 to 1073741824 for accuracy since the conversion factor of bytes is 1024 and not 1000
                }
                $usersNoActivityDate += $userData
                $usersNoActivityDateCount++
            }
    }
  
          
}


$oneDriveReportAll.AllUsers = $usersReport
$oneDriveReportAll.TotalUsersCount = $totalUsers
$oneDriveReportAll.ActiveUsersCount = $usersCount
$oneDriveReportAll.DeletedUsers = $usersDeleted
$oneDriveReportAll.UsersNoPrimaryAddress = $usersNoPrimaryAddress
$oneDriveReportAll.UsersNoActivityDate = $usersNoActivityDate
$oneDriveReportAll.ReportRefreshDate = $reportRefreshDate

return $oneDriveReportAll
}

Export-ModuleMember -Function *
