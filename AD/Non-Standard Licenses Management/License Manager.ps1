Import-Module "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Non-Standard Licenses Management\License Expiry Alert.psm1" -Verbose -Force
# Load the Import-Excel module
Import-Module -Name ImportExcel

# Define the paths for the Excel file and the CSV file
$excelFilePath = "C:\Users\lkadmin\CIFOR-ICRAF\Information Communication Technology (ICT) - Microsoft 365 Licensing\Non-Standard License Data.xlsx"
$archiveCsvFilePath = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Non-Standard Licenses Management\License Data Archive.csv"
$securityGroupsFilePath = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Non-Standard Licenses Management\License Security Group Map.csv"

# Read the Excel file
$excelData = Import-Excel -Path $excelFilePath
# Read the security group license mapper csv
$securityGroups = Import-Csv -Path $securityGroupsFilePath


# Expiry dates for accounts that are to be notified of the expiry
$date_in_three_days = (Get-Date).AddDays(3).ToString("dd/MM/yyyy 00:00:00")
$date_in_two_days = (Get-Date).AddDays(2).ToString("dd/MM/yyyy 00:00:00")
$date_in_one_day = (Get-Date).AddDays(1).ToString("dd/MM/yyyy 00:00:00")

# Expiry date for account to be notified of the license expiration
$date_today = (Get-Date).ToString("dd/MM/yyyy 00:00:00")

$expired_licenses = @()
$accounts_nearing_expiry = @()

foreach ($row in $excelData) {
   if($row."End".ToString("dd/MM/yyyy 00:00:00") -eq $date_today)
   {
       $emailAddress = $row."Email Address"
       $license = $row."License"
       $user = Get-ADUser -Filter {UserPrincipalName -eq $emailAddress} -Properties GivenName, Surname

       # Disable exclude the user from the license group
       foreach($securityGroup in $securityGroups)
       {
           if($securityGroup."License" -eq $license)
           {
               Remove-ADGroupMember -Identity $securityGroup."Security Group" -Members $user -Confirm:$false
           }
       }
       <#
       # Delete rows that match the email address
       $new_data = $excelData | Where-Object { $_."Email Address" -ne $emailAddress }

       # Save the changes
       $new_data | Export-Excel -Path $excelFilePath -AutoSize -WorksheetName "Sheet1" -ClearSheet
       #>
       $data = [PSCustomObject]@{
           "First Name"= $user.GivenName
           "Last Name"= $user.Surname
           "Email Address"= $emailAddress
           "Institution"= $row."Institution"
           "License"= $license
           "Start"= $row."Start"
           "End"= $row."End"       
       }
       Send-LicenseExpiryAlert -accounts_to_alert $data
       $expired_licenses += $data
   }
   elseif($row."End".ToString("dd/MM/yyyy 00:00:00") -eq $date_in_three_days)
   {
       $emailAddress = $row."Email Address"
       $user = Get-ADUser -Filter {UserPrincipalName -eq $emailAddress} -Properties GivenName, Surname
       $data = [PSCustomObject]@{
           "First Name"= $user.GivenName
           "Last Name"= $user.Surname
           "Start"= $row."Start"
           "End"= $row."End"
           "License"= $row."License"
           "Email Address"= $emailAddress
           "Expiry"= "3 days"
       }
       Send-LicenseExpiryWarning -accounts_to_alert $data
   }
   elseif($row."End".ToString("dd/MM/yyyy 00:00:00") -eq $date_in_two_days)
   {
       $emailAddress = $row."Email Address"
       $user = Get-ADUser -Filter {UserPrincipalName -eq $emailAddress} -Properties GivenName, Surname
       $data = [PSCustomObject]@{
           "First Name"= $user.GivenName
           "Last Name"= $user.Surname
           "Start"= $row."Start"
           "End"= $row."End"
           "License"= $row."License"
           "Email Address"= $emailAddress
           "Expiry"= "2 days"
       }
       Send-LicenseExpiryWarning -accounts_to_alert $data
   }
   elseif($row."End".ToString("dd/MM/yyyy 00:00:00") -eq $date_in_one_day)
   {
       $emailAddress = $row."Email Address"
       $user = Get-ADUser -Filter {UserPrincipalName -eq $emailAddress} -Properties GivenName, Surname
       $data = [PSCustomObject]@{
           "First Name"= $user.GivenName
           "Last Name"= $user.Surname
           "Start"= $row."Start"
           "End"= $row."End"
           "License"= $row."License"
           "Email Address"= $emailAddress
           "Expiry"= "1 day"
       }
       Send-LicenseExpiryWarning -accounts_to_alert $data
   }
}