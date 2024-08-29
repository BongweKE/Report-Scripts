Import-Module "C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Account Expiry Alert\Account Expiry Alert.psm1" -Verbose -Force

$selectObject = @('Name', 'samAccountName', 'ObjectClass', 'AccountExpirationDate', 'lastLogonDate', 'Enabled', 'PasswordNeverExpires')
$ICRAF_ExcludedOU = @('ICRAF BOT','ICRAF Spouses','Disabled accounts','ICRAF Kenya General Accounts','ICRAF Administrators','ICRAF MFI','ICRAF Meeting Rooms','ICRAF ICT','ICRAF OCS','EC Regreening Project','ICRAF Sharepoint Users','ICRAF Kenya Shared Mailboxes','ICRAF Kenya Service Accounts') #update with KE data
$ICRAFOU = 'OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG'
$hr_focal_point_file_path = "C:\Users\lkadmin\Desktop\Account Expiry Alert\HR Focal Points.csv"
$hr_focal_points_data = Import-Csv -Path $hr_focal_point_file_path
$ICRAFExcludedOURegex = ($ICRAF_ExcludedOU | ForEach-Object { [regex]::Escape($_) }) -join '|'
[datetime]$In30Days = (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(30)
[datetime]$In14Days = (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(14)
[datetime]$In7Days = (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(7)
[datetime]$In1Day = (Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(1)

#Accounts expiring in 1 month
$date_before_one_month = (Get-Date).AddDays(29).ToString("yyyy-MM-dd")
$date_after_one_month = (Get-Date).AddDays(31).ToString("yyyy-MM-dd")
$ICRAFExpiring_1Month =  Get-ADUser -SearchBase $ICRAFOU -Filter "Enabled -eq `$true -and AccountExpirationDate -gt '$date_before_one_month' -and AccountExpirationDate -lt '$date_after_one_month'" -Properties GivenName, Surname, UserPrincipalName, AccountExpirationDate, Manager, co | Where-Object { -not ($_.DistinguishedName -match $ICRAFExcludedOURegex) }

#Accounts expiring in 2 weeks
$date_before_two_weeks = (Get-Date).AddDays(13).ToString("yyyy-MM-dd")
$date_after_two_weeks = (Get-Date).AddDays(15).ToString("yyyy-MM-dd")
$ICRAFExpiring_2Weeks =  Get-ADUser -SearchBase $ICRAFOU -Filter "Enabled -eq `$true -and AccountExpirationDate -gt '$date_before_two_weeks' -and AccountExpirationDate -lt '$date_after_two_weeks'" -Properties GivenName, Surname, UserPrincipalName, AccountExpirationDate, Manager, co | Where-Object { -not ($_.DistinguishedName -match $ICRAFExcludedOURegex) }

#Accounts expiring in 1 week
$date_before_one_week = (Get-Date).AddDays(6).ToString("yyyy-MM-dd")
$date_after_one_week = (Get-Date).AddDays(8).ToString("yyyy-MM-dd")
$ICRAFExpiring_1Week = Get-ADUser -SearchBase $ICRAFOU -Filter "Enabled -eq `$true -and AccountExpirationDate -gt '$date_before_one_week' -and AccountExpirationDate -lt '$date_after_one_week'" -Properties GivenName, Surname, UserPrincipalName, AccountExpirationDate, Manager, co | Where-Object { -not ($_.DistinguishedName -match $ICRAFExcludedOURegex) }

#Accounts expiring in 1 day
$date_before_one_day = (Get-Date).ToString("yyyy-MM-dd")
$date_after_one_day = (Get-Date).AddDays(2).ToString("yyyy-MM-dd")
$ICRAFExpiring_1Day = Get-ADUser -SearchBase $ICRAFOU -Filter "Enabled -eq `$true -and AccountExpirationDate -gt '$date_before_one_day' -and AccountExpirationDate -lt '$date_after_one_day'" -Properties GivenName, Surname, UserPrincipalName, AccountExpirationDate, Manager, co | Where-Object { -not ($_.DistinguishedName -match $ICRAFExcludedOURegex) }

$Accounts_Nearing_Expiry = @()
$hr_focal_point = 'w.wambugu@cifor-icraf.org'
    
$ICRAFExpiring_1Month | ForEach-Object {
    if ($_.co -ne $null) {
        foreach ($focal in $hr_focal_points_data) {
            foreach ($country in $focal."Country/Region" | ConvertFrom-Json) {
                if ($country -eq $_.co) {
                    $hr_focal_point = $focal.Email
                }
            }  
        }    
    }
    if ($_.Manager -ne $null) {
        $SupervisorEmail = (get-aduser $_.Manager -properties * | Select UserPrincipalName).UserPrincipalName
    }
    else {
        $SupervisorEmail = $hr_focal_point
    }
    $accounts_expiring_in_1Month = [PSCustomObject]@{
        GivenName = $_.GivenName
        Surname = $_.Surname
        EmailAddress = $_.UserPrincipalName
        AccountExpiry = "30 days"
        AccountExpirationDate = $_.AccountExpirationDate.ToString("dd-MM-yyyy")
        SupervisorEmailAddress = $SupervisorEmail
        HrFocalPoint = $hr_focal_point 
    }
    $Accounts_Nearing_Expiry += $accounts_expiring_in_1Month
}

$ICRAFExpiring_2Weeks | ForEach-Object {
    if ($_.co -ne $null) {
        foreach ($focal in $hr_focal_points_data) {
            foreach ($country in $focal."Country/Region" | ConvertFrom-Json) {
                if ($country -eq $_.co) {
                    $hr_focal_point = $focal.Email
                }
            }  
        }    
    }
    if ($_.Manager -ne $null) {
        $SupervisorEmail = (get-aduser $_.Manager -properties * | Select UserPrincipalName).UserPrincipalName
    }
    else {
        $SupervisorEmail = $hr_focal_point
    }
    $accounts_expiring_in_2Weeks = [PSCustomObject]@{
        GivenName = $_.GivenName
        Surname = $_.Surname
        EmailAddress = $_.UserPrincipalName
        AccountExpiry = "14 days"
        AccountExpirationDate = $_.AccountExpirationDate.ToString("dd-MM-yyyy")
        SupervisorEmailAddress = $SupervisorEmail
        HrFocalPoint = $hr_focal_point
    }
    $Accounts_Nearing_Expiry += $accounts_expiring_in_2Weeks
}

$ICRAFExpiring_1Week | ForEach-Object {
    if ($_.co -ne $null) {
        foreach ($focal in $hr_focal_points_data) {
            foreach ($country in $focal."Country/Region" | ConvertFrom-Json) {
                if ($country -eq $_.co) {
                    $hr_focal_point = $focal.Email
                }
            }  
        }    
    }
    if ($_.Manager -ne $null) {
        $SupervisorEmail = (get-aduser $_.Manager -properties * | Select UserPrincipalName).UserPrincipalName
    }
    else {
        $SupervisorEmail = $hr_focal_point
    }
    $accounts_expiring_in_1Week = [PSCustomObject]@{
        GivenName = $_.GivenName
        Surname = $_.Surname
        EmailAddress = $_.UserPrincipalName
        AccountExpiry = "7 days"
        AccountExpirationDate = $_.AccountExpirationDate.ToString("dd-MM-yyyy")
        SupervisorEmailAddress = $SupervisorEmail
        HrFocalPoint = $hr_focal_point
    }
    $Accounts_Nearing_Expiry += $accounts_expiring_in_1Week
}

$ICRAFExpiring_1Day | ForEach-Object {
    if ($_.co -ne $null) {
        foreach ($focal in $hr_focal_points_data) {
            foreach ($country in $focal."Country/Region" | ConvertFrom-Json) {
                if ($country -eq $_.co) {
                    $hr_focal_point = $focal.Email
                }
            }  
        }    
    }
    if ($_.Manager -ne $null) {
        $SupervisorEmail = (get-aduser $_.Manager -properties * | Select UserPrincipalName).UserPrincipalName
    }
    else {
        $SupervisorEmail = $hr_focal_point
    }
    $accounts_expiring_in_1Day = [PSCustomObject]@{
        GivenName = $_.GivenName
        Surname = $_.Surname
        EmailAddress = $_.UserPrincipalName
        AccountExpiry = "1 day"
        AccountExpirationDate = $_.AccountExpirationDate.ToString("dd-MM-yyyy")
        SupervisorEmailAddress = $SupervisorEmail
        HrFocalPoint = $hr_focal_point
    }
    $Accounts_Nearing_Expiry += $accounts_expiring_in_1Day
}

Send-AccountExpiryAlert -accounts_to_alert $Accounts_Nearing_Expiry