Import-Module "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Account Expiry Alert\Account Expiry Alert.psm1" -Verbose -Force

<#
--------------------------------------------------------------
Prep for Student Email Filtering
--------------------------------------------------------------
#>

# Import the Active Directory module
Import-Module ActiveDirectory

# Define the suffixes array
$suffixes = @("Students", "Interns")

# Initialize an empty array to store matching OUs
$matchingOUs = @()

# Loop through each suffix
foreach ($suffix in $suffixes) {
    # Find OUs that match the current suffix
    $matchingOUs += Get-ADOrganizationalUnit -Filter * | Where-Object { $_.Name -like "*$suffix" }
}

<#
Get-StudentStatus: A Function to Check if a given User Object matches the OUs of Students in CIFOR-ICRAF
#>
function Get-StudentStatus {
    [CmdletBinding()]
    param (
        # The user object to check
        [Parameter(Mandatory)]
        [Microsoft.ActiveDirectory.Management.ADUser]
        $User,

        # List of OUs to compare against
        [Parameter(Mandatory)]
        [Array]
        $MatchingOUs
    )

    # Get the DistinguishedName of the user
    $userDN = $User.DistinguishedName

    # Check if the user's DN ends with any of the matching OU distinguished names
    $isInMatchingOU = $MatchingOUs | ForEach-Object {
        if ($userDN.EndsWith($_.DistinguishedName)) { return $true }
    }

    return $isInMatchingOU -eq $true
}


<#
--------------------------------------------------------------
Prep for All Accounts
--------------------------------------------------------------
#>

$selectObject = @('Name', 'samAccountName', 'ObjectClass', 'AccountExpirationDate', 'lastLogonDate', 'Enabled', 'PasswordNeverExpires')
$ICRAF_ExcludedOU = @('ICRAF BOT','ICRAF Spouses','Disabled accounts','ICRAF Kenya General Accounts','ICRAF Administrators','ICRAF MFI','ICRAF Meeting Rooms','ICRAF ICT','ICRAF OCS','EC Regreening Project','ICRAF Sharepoint Users','ICRAF Kenya Shared Mailboxes','ICRAF Kenya Service Accounts') #update with KE data
$ICRAFOU = 'OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG'
$hr_focal_point_file_path = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Scripts\AD\Account Expiry Alert\HR Focal Points.csv"
$hr_focal_points_data = Import-Csv -Path $hr_focal_point_file_path
$ICRAFExcludedOURegex = ($ICRAF_ExcludedOU | ForEach-Object { [regex]::Escape($_) }) -join '|'
[datetime]$In30Days = (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(30)
[datetime]$In14Days = (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(14)
[datetime]$In7Days = (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(7)
[datetime]$In1Day = (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(1)

<#------------- Sometimes I change date For Testing Purposes -------------------------------------#>
# $today = (Get-Date).AddDays(-1)
$today = Get-Date

#Accounts expiring in 1 month

$date_before_one_month = ($today).AddDays(29).ToString("yyyy-MM-dd")
$date_after_one_month = ($today).AddDays(31).ToString("yyyy-MM-dd")
$ICRAFExpiring_1Month =  Get-ADUser -SearchBase $ICRAFOU -Filter "Enabled -eq `$true -and AccountExpirationDate -gt '$date_before_one_month' -and AccountExpirationDate -lt '$date_after_one_month'" -Properties GivenName, Surname, UserPrincipalName, AccountExpirationDate, Manager, co | Where-Object { -not ($_.DistinguishedName -match $ICRAFExcludedOURegex) }

#Accounts expiring in 2 weeks
$date_before_two_weeks = ($today).AddDays(13).ToString("yyyy-MM-dd")
$date_after_two_weeks = ($today).AddDays(15).ToString("yyyy-MM-dd")
$ICRAFExpiring_2Weeks =  Get-ADUser -SearchBase $ICRAFOU -Filter "Enabled -eq `$true -and AccountExpirationDate -gt '$date_before_two_weeks' -and AccountExpirationDate -lt '$date_after_two_weeks'" -Properties GivenName, Surname, UserPrincipalName, AccountExpirationDate, Manager, co | Where-Object { -not ($_.DistinguishedName -match $ICRAFExcludedOURegex) }

#Accounts expiring in 1 week
$date_before_one_week = ($today).AddDays(6).ToString("yyyy-MM-dd")
$date_after_one_week = ($today).AddDays(8).ToString("yyyy-MM-dd")
$ICRAFExpiring_1Week = Get-ADUser -SearchBase $ICRAFOU -Filter "Enabled -eq `$true -and AccountExpirationDate -gt '$date_before_one_week' -and AccountExpirationDate -lt '$date_after_one_week'" -Properties GivenName, Surname, UserPrincipalName, AccountExpirationDate, Manager, co | Where-Object { -not ($_.DistinguishedName -match $ICRAFExcludedOURegex) }

#Accounts expiring in 1 day
$date_before_one_day = ($today).ToString("yyyy-MM-dd")
$date_after_one_day = ($today).AddDays(2).ToString("yyyy-MM-dd")
$ICRAFExpiring_1Day = Get-ADUser -SearchBase $ICRAFOU -Filter "Enabled -eq `$true -and AccountExpirationDate -gt '$date_before_one_day' -and AccountExpirationDate -lt '$date_after_one_day'" -Properties GivenName, Surname, UserPrincipalName, AccountExpirationDate, Manager, co | Where-Object { -not ($_.DistinguishedName -match $ICRAFExcludedOURegex) }

$Accounts_Nearing_Expiry = @()
$hq_focal_point = 'w.wambugu@cifor-icraf.org'
$kic_focal_point = 'I.Ingumba@cifor-icraf.org'
    
<#
--------------------------------------------------------------
Action on each group
--------------------------------------------------------------
#>

$ICRAFExpiring_1Month | ForEach-Object {
    $hr_focal_point = $hq_focal_point # Set Default to walter
    if ($_.co -ne $null) {
        foreach ($focal in $hr_focal_points_data) {
            foreach ($country in $focal."Country/Region" | ConvertFrom-Json) {
                if ($country -eq $_.co) {
                    $hr_focal_point = $focal.Email
                }
            }  
        }    
    } Else {
        # If the user Object's Country Attribute is Null
        $hr_focal_point = 'w.wambugu@cifor-icraf.org'
    }
    If ($_.Manager -ne $null) {
        # Students and employees who have Supervisor as part of their AD deets
        $SupervisorEmail = (get-aduser $_.Manager -properties * | Select UserPrincipalName).UserPrincipalName
    }
    else {
        # Students and Employees who's Supervisor deets are missing
        If (Get-StudentStatus -User $_ -MatchingOUs $matchingOUs) {
            $SupervisorEmail = $kic_focal_point
        } else {
            $SupervisorEmail = $hr_focal_point
        }
        
    }

    # Choose Focal Point depending on Whether we have a student or Staff
    If (Get-StudentStatus -User $_ -MatchingOUs $matchingOUs) {
        $focal_point = $kic_focal_point
    } 
    Else {
        $focal_point = $hr_focal_point 
    }

    $accounts_expiring_in_1Month = [PSCustomObject]@{
        GivenName = $_.GivenName
        Surname = $_.Surname
        EmailAddress = $_.UserPrincipalName
        AccountExpiry = "30 days"
        AccountExpirationDate = $_.AccountExpirationDate.ToString("dd-MM-yyyy")
        SupervisorEmailAddress = $SupervisorEmail
        HrFocalPoint = $focal_point 
    }
    $Accounts_Nearing_Expiry += $accounts_expiring_in_1Month
}

$ICRAFExpiring_2Weeks | ForEach-Object {
    $hr_focal_point = $hq_focal_point # Set Default to walter
    if ($_.co -ne $null) {
        foreach ($focal in $hr_focal_points_data) {
            foreach ($country in $focal."Country/Region" | ConvertFrom-Json) {
                if ($country -eq $_.co) {
                    $hr_focal_point = $focal.Email
                }
            }  
        }    
    } Else {
        # If the user Object's Country Attribute is Null
        $hr_focal_point = 'w.wambugu@cifor-icraf.org'
    }
    if ($_.Manager -ne $null) {
        $SupervisorEmail = (get-aduser $_.Manager -properties * | Select UserPrincipalName).UserPrincipalName
    }
    else {
        # Students and Employees who's Supervisor deets are missing
        If (Get-StudentStatus -User $_ -MatchingOUs $matchingOUs) {
            $SupervisorEmail = $kic_focal_point
        } else {
            $SupervisorEmail = $hr_focal_point
        }
        
    }

    # Choose Focal Point depending on Whether we have a student or Staff
    If (Get-StudentStatus -User $_ -MatchingOUs $matchingOUs) {
        $focal_point = $kic_focal_point
    } 
    Else {
        $focal_point = $hr_focal_point 
    }

    $accounts_expiring_in_2Weeks = [PSCustomObject]@{
        GivenName = $_.GivenName
        Surname = $_.Surname
        EmailAddress = $_.UserPrincipalName
        AccountExpiry = "14 days"
        AccountExpirationDate = $_.AccountExpirationDate.ToString("dd-MM-yyyy")
        SupervisorEmailAddress = $SupervisorEmail
        HrFocalPoint = $focal_point
    }
    $Accounts_Nearing_Expiry += $accounts_expiring_in_2Weeks
}

$ICRAFExpiring_1Week | ForEach-Object {
    $hr_focal_point = $hq_focal_point # Set Default to walter
    if ($_.co -ne $null) {
        foreach ($focal in $hr_focal_points_data) {
            foreach ($country in $focal."Country/Region" | ConvertFrom-Json) {
                if ($country -eq $_.co) {
                    $hr_focal_point = $focal.Email
                }
            }  
        }    
    } Else {
        # If the user Object's Country Attribute is Null
        $hr_focal_point = 'w.wambugu@cifor-icraf.org'
    }

    if ($_.Manager -ne $null) {
        $SupervisorEmail = (get-aduser $_.Manager -properties * | Select UserPrincipalName).UserPrincipalName
    }
    else {
        # Students and Employees who's Supervisor deets are missing
        If (Get-StudentStatus -User $_ -MatchingOUs $matchingOUs) {
            $SupervisorEmail = $kic_focal_point
        } else {
            $SupervisorEmail = $hr_focal_point
        }
        
    }

    # Choose Focal Point depending on Whether we have a student or Staff
    If (Get-StudentStatus -User $_ -MatchingOUs $matchingOUs) {
        $focal_point = $kic_focal_point
    } 
    Else {
        $focal_point = $hr_focal_point 
    }

    $accounts_expiring_in_1Week = [PSCustomObject]@{
        GivenName = $_.GivenName
        Surname = $_.Surname
        EmailAddress = $_.UserPrincipalName
        AccountExpiry = "7 days"
        AccountExpirationDate = $_.AccountExpirationDate.ToString("dd-MM-yyyy")
        SupervisorEmailAddress = $SupervisorEmail
        HrFocalPoint = $focal_point
    }
    $Accounts_Nearing_Expiry += $accounts_expiring_in_1Week
}


$ICRAFExpiring_1Day | ForEach-Object {
    $hr_focal_point = $hq_focal_point # Set Default to walter
    if ($_.co -ne $null) {
        # For Objects for which the Country Is Available, we use this info to choose Focal Point
        foreach ($focal in $hr_focal_points_data) {
            foreach ($country in $focal."Country/Region" | ConvertFrom-Json) {
                if ($country -eq $_.co) {
                    $hr_focal_point = $focal.Email
                }
            }  
        }    
    } Else {
        # If the user Object's Country Attribute is Null
        $hr_focal_point = 'w.wambugu@cifor-icraf.org'
    }

    if ($_.Manager -ne $null) {
        $SupervisorEmail = (get-aduser $_.Manager -properties * | Select UserPrincipalName).UserPrincipalName
    }
    else {
        # Students and Employees who's Supervisor deets are missing
        If (Get-StudentStatus -User $_ -MatchingOUs $matchingOUs) {
            $SupervisorEmail = $kic_focal_point
        } else {
            $SupervisorEmail = $hr_focal_point
        }
        
    }

    # Choose Focal Point depending on Whether we have a student or Staff
    If (Get-StudentStatus -User $_ -MatchingOUs $matchingOUs) {
        $focal_point = $kic_focal_point
    } 
    Else {
        $focal_point = $hr_focal_point 
    }

    $accounts_expiring_in_1Day = [PSCustomObject]@{
        GivenName = $_.GivenName
        Surname = $_.Surname
        EmailAddress = $_.UserPrincipalName
        AccountExpiry = "1 day"
        AccountExpirationDate = $_.AccountExpirationDate.ToString("dd-MM-yyyy")
        SupervisorEmailAddress = $SupervisorEmail
        HrFocalPoint = $focal_point
    }
    $Accounts_Nearing_Expiry += $accounts_expiring_in_1Day
}

# $Accounts_Nearing_Expiry | Select-Object EmailAddress, SupervisorEmailAddress, HrFocalPoint, AccountExpiry, OU | Format-Table -AutoSize

Send-AccountExpiryAlert -accounts_to_alert $Accounts_Nearing_Expiry


<#
-------------------------------------
TESTING ONE USERS FOCAL POINT
-------------------------------------
$hr_focal_point_file_path = "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Scripts\AD\Account Expiry Alert\HR Focal Points.csv"
$hr_focal_points_data = Import-Csv -Path $hr_focal_point_file_path
$x1 = Get-ADUser -Identity GReynoso -Properties co, Manager
foreach ($focal in $hr_focal_points_data) {
    foreach ($country in $focal."Country/Region" | ConvertFrom-Json) {
        if ($country -eq $x1.co) {
            $hr_focal_point = $focal.Email
        }
    }  
}   

$hr_focal_point
#>