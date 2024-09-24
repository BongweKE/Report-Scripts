# Load the Import-Excel module
Import-Module ImportExcel
$reportDate = Get-Date -Format 'MMMM yyyy'
# $location = "Kenya"
# $alertLocation = "ICRAF"
$selectObject = @('Name', 'samAccountName', 'ObjectClass', 'AccountExpirationDate', 'lastLogonDate', 'Enabled', 'PasswordNeverExpires')
$ICRAF_ExcludedOU = @('OCS Users', 'ICRAF BOT','ICRAF Spouses','Disabled accounts','ICRAF Kenya General Accounts','ICRAF Administrators','ICRAF MFI','ICRAF Meeting Rooms','ICRAF ICT','ICRAF OCS','EC Regreening Project','ICRAF Sharepoint Users','ICRAF Kenya Shared Mailboxes','ICRAF Kenya Service Accounts', 'ICRAFHQ Servers') #update with KE data
$ICRAFOU = 'OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG'
$ICRAFComputersOU = 'OU=Computers,OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG'
# $inactiveDisabledOU = 'OU=Disabled Due To Inactivity,OU=Disabled accounts,OU=ICRAF Kenya,OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG'
$csvo365UsageData = 'C:\Users\lkadmin\CIFOR-ICRAF\Information Communication Technology (ICT) - Reports Archive\AD Reports\O365 Exchange Mail Usage\'
###########################################################################################################
# Change date below to conform with report once this is automated by tasker
###########################################################################################################
#$csvFileDate = (Get-Date).AddDays(-1).ToString("yyyy-MM-dd")
$csvFileDate = "2024-08-30"
$csvFileName = $csvo365UsageData + $csvFileDate + '.csv'
###########################################################################################################
# Change from Test folder before submiting to Tasker
###########################################################################################################
$reportDirectory = 'C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Results\AD\ICRAF\Test'
# $reportDirectory = 'C:\Users\lkadmin\OneDrive - CIFOR-ICRAF\Desktop\Auto Reports\Report Results\AD\ICRAF\'
$reportDirectoryCurrent = $reportDirectory + $reportDate + '\'
$compressDirectory = $reportDirectoryCurrent + '*'
$compressedDirectory = $reportDirectoryCurrent + 'AuditReport.zip'
# $excUsername = 'CIFORICRAFAutoReport@cifor-icraf.org'
# $excPassword = ConvertTo-SecureString -String 'Winter2023' -AsPlainText -Force #Change to secure mode credential after testing
# $excCreds = New-Object System.Management.Automation.PSCredential($excUsername,$excPassword) 

###########################################################################################################
# Get o365 csv File
###########################################################################################################
$csvReport = Import-Csv -Path $csvFileName
$mailboxCount = 0

###########################################################################################################
# START
###########################################################################################################

###########################################################################################################
# new dormant accounts that had not logged in for the past 6 months (180 days) 
# 12 dormant accounts that had not logged in for the past 6 months (180 days) 
###########################################################################################################
$dormantThreshold = 180
$dormantDate = (Get-Date).AddDays(-$dormantThreshold)

# Construct the regex for excluded OUs
$excludedOURegex = $ICRAF_ExcludedOU -join '|'
# Get the users that meet the criteria
# $ICRAFDormantAccountsAD = Get-ADUser -Filter { Enabled -eq $true } -SearchBase $ICRAFOU -Properties * | Where-Object {
#     $_.DistinguishedName -notmatch $excludedOURegex -and
#     $_.LastLogonDate -lt $dormantDate
# }

$ICRAFDormantAccountsAD = Get-ADUser -Filter { Enabled -eq $true } -SearchBase $ICRAFOU -Properties * | Where-Object {
    $_.DistinguishedName -notmatch $excludedOURegex -and
    $_.LastLogonDate -lt $dormantDate
}
# Get-ADUser -Filter '(PasswordLastSet -lt $d) -or (LastLogonTimestamp -lt $d)'
$d = [DateTime]::Today.AddDays(-180)
$ICRAFDormantAccountsAD = Get-ADUser -Filter { Enabled -eq $true } -SearchBase $ICRAFOU -Properties * | Where-Object {
    $_.DistinguishedName -notmatch $excludedOURegex -and
    $_.LastLogonDate -and # is not null
    $_.LastLogonTimestamp -and # is not null
    $_.LastLogonDate -lt $dormantDate -and
    [DateTime]::$_.LastLogonTimestamp -lt $d 
}

$ICRAFDormantAccountsAD.count # 99

# $ICRAFDormantAccountsAD | Select-Object SamAccountName, Name, LastLogonDate, EmailAddress
###########################################################################################################
# active accounts with blank last logon dates. 
# 91 active accounts with blank last logon dates. 

###########################################################################################################
$ICRAFBlankLogonAD = Get-ADUser -Filter { Enabled -eq $true } -SearchBase $ICRAFOU -Properties * | Where-Object {
    $_.DistinguishedName -notmatch $excludedOURegex -and
    $null -eq $_.LastLogonDate
}

$ICRAFBlankLogonAD.count # 

#####################################################################################################
# active accounts that had not changed their passwords in the last 90 days
# 84 active accounts with expired passwords. : Here we go with expired passwords
#####
# Newest
$ICRAFExpiredPwd = Get-ADUser -Filter { Enabled -eq $true } -SearchBase $ICRAFOU -Properties * | Where-Object {
    $_.DistinguishedName -notmatch $excludedOURegex -and 
    (($_.LastLogonDate -and $_.LastLogonDate -lt (Get-Date).AddDays(-90)) -or
    ($_.LastLogonTimestamp -and [DateTime]::$_.LastLogonTimestamp -lt [DateTime]::Today.AddDays(-180))) -and # active
    $_.PasswordExpired  -eq $true
}

$ICRAFExpiredPwd | Select-Object -First 3 # 113

#####################################################################################################
$Days = 90
$Date = (Get-Date).AddDays(-$Days)

# Construct the regex for excluded OUs
$excludedOURegex = $ICRAF_ExcludedOU -join '|'

# Get the users that meet the criteria
$inactiveUsers = Get-ADUser -Filter { Enabled -eq $true } -SearchBase $ICRAFOU -Properties msDS-UserPasswordExpiryTimeComputed, DistinguishedName | Where-Object {
    $_.DistinguishedName -notmatch $excludedOURegex -and
    $_.'msDS-UserPasswordExpiryTimeComputed' -and
    ([datetime]::FromFileTime($_.'msDS-UserPasswordExpiryTimeComputed')) -lt $Date
}

$inactiveUsers.Count

$today = (Get-Date).AddDays(-1)
# Get accounts whose passwords have expired
$expiredAccounts = Get-ADUser -Filter {
    Enabled -eq $True -and 
    PasswordNeverExpires -eq $False
} -SearchBase $ICRAFOU -Properties "DisplayName", "msDS-UserPasswordExpiryTimeComputed" |
Select-Object -Property "DisplayName",
    @{Name="ExpiryDate"; Expression={
        if ($_.'msDS-UserPasswordExpiryTimeComputed') {
            [datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")
        } else {
            $null  # No expiry date available
        }
    }} | Where-Object {
        $_.ExpiryDate -and $_.ExpiryDate -lt $today -and
        $_.DistinguishedName -notmatch $excludedOURegex
    }


$expiredAccounts | Format-Table -AutoSize

$expiredAccounts.count


$temp = Get-ADUser -Filter * | Where-Object {
    [datetime]::FromFileTime($_.'msDS-UserPasswordExpiryTimeComputed') -gt $today
} | Select-Object SamAccountName, Name, LastLogonDate

$temp | Select-Object -Last 5 | Format-Table -AutoSize 


$today = Get-Date

# Get users with expired passwords
$expiredAccounts = Get-ADUser -Filter {
    Enabled -eq $True -and 
    PasswordNeverExpires -eq $False
} -Properties * |
Select-Object -Property "DisplayName",
    @{Name="ExpiryDate"; Expression={
        # Ensure msDS-UserPasswordExpiryTimeComputed is not null before converting
        if ($_.'msDS-UserPasswordExpiryTimeComputed') {
            [datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")
        } else {
            $null  # Return null if the attribute is not set
        }
    }} | Where-Object {
        $_.ExpiryDate -ne $null -and $_.ExpiryDate -lt $today
    }

# Output the list of expired accounts
$expiredAccounts | Format-Table -AutoSize
# Output the expired accounts



# Also check that 



#####################################################################################################
#Accounts With Password Set Not To Expire
# 12 active accounts with passwords that do not expire. 
#####################################################################################################
$ICRAFAccounts_PasswordsNeverExpireAD = Search-ADAccount -SearchBase $ICRAFOU -PasswordNeverExpires | Where 'Enabled' -eq 'True' | Where 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|')
$ICRAFAccounts_PasswordsNeverExpire = @()
foreach ($account in $ICRAFAccounts_PasswordsNeverExpireAD) {
    $searchValue = $account.'Name'
    $result = $csvReport | Where-Object { $_.$columnName -eq $searchValue }

    # Check if result exists and if the Last Activity Date is more than 90 days ago
    if ($result -and $result.'Is Delete' -eq 'FALSE') {
        $ICRAFAccounts_PasswordsNeverExpire += $account
    }
}

#####################################################################################################

<#

103 active accounts that had not changed their passwords in the last 90 days. 
#>

#####################################################################################################
# Assuming Expired means AccountExpired
$expiredAccounts = Search-ADAccount -SearchBase $ICRAFOU -AccountExpired | Where 'DistinguishedName' -NotMatch ($ICRAF_ExcludedOU -join '|')

<#
$compressedDirectory
# organization overview
$userAccountCount
$computerCount

$ICRAFDormantAccounts # dormant accounts that had not logged in for the past 6 months (180 days) 
$ICRAFAccountsNoLastActivity # active accounts with blank last logon dates. 
$Expired Passwords # active accounts that had not changed their passwords in the last 90 days
$ICRAFAccounts_PasswordsNeverExpireAD # AD Accounts With Password Set Not To Expire
$ICRAFAccounts_PasswordsNeverExpire # Onedrive Accounts With Password Set Not To Expire
$expiredAccounts # expired AD accounts
#>
#####################################################################################################

###########################################################################################################
#Generate Report From Exchange Online Session
###########################################################################################################
$userAccountCount = (Get-ADUser -Filter * -SearchBase $ICRAFOU).count
$computerCount = (Get-ADComputer -Filter * -SearchBase $ICRAFComputersOU).count

# Archive Full Report in dashboard data source
###########################################################################################################
# Use test Excel file before we add this to the task scheduler (pipeline)
###########################################################################################################

# $ICRAFReportRecipient = @('l.kavoo@cifor-icraf.org','servicedesk@cifor-icraf.org','c.mwangi@cifor-icraf.org','b.obaga@cifor-icraf.org','g.kirimi@cifor-icraf.org','s.mariwa@cifor-icraf.org','p.oyuko@cifor-icraf.org','r.kande@cifor-icraf.org')
# $ICRAFReportRecipient = @('l.kavoo@cifor-icraf.org', 'b.obaga@cifor-icraf.org', 'g.kirimi@cifor-icraf.org')
$ICRAFReportRecipient = @('l.kavoo@cifor-icraf.org', 'b.obaga@cifor-icraf.org')
#Send Email to Recipients
$smtpServer = 'SMTP.Office365.com'
$alertMailUserName = 'CIFORICRAFAutoReport@cifor-icraf.org'
$alertMailPassword = ConvertTo-SecureString -String 'Winter2023' -AsPlainText -Force #Change to secure mode credential after testing
$mailCredential = New-Object System.Management.Automation.PSCredential($alertMailUserName,$alertMailPassword)
$emailSubject = "AD Audit Report"



$emailBody = @"

Audit Report Overview:

- Total User Accounts: $userAccountCount
- Total Computer Accounts: $computerCount

- Dormant Accounts (not logged in for 180 days): $($ICRAFDormantAccounts.Count)
- Active Accounts with Blank Last Logon Dates: $($ICRAFAccountsNoLastActivity.Count)
- Active Accounts Not Changed Password in 90 Days: $($Expired Passwords.Count)
- AD Accounts with Passwords Set Not to Expire: $($ICRAFAccounts_PasswordsNeverExpireAD.Count)
- OneDrive Accounts with Passwords Set Not to Expire: $($ICRAFAccounts_PasswordsNeverExpire.Count)
- Expired AD Accounts: $($expiredAccounts.Count)

Please find the detailed report attached.

"@

$ExcelFileDate = (Get-Date).AddDays(-1).ToString("yyyy-MM-dd")

$workbookPath = 'c:\Users\poadmin\CIFOR-ICRAF\Information Communication Technology (ICT) - Reports Archive\Audit Reports-CIFORICRAFRASVR\' + $ExcelFileDate + '.xlsx'

# Define the structure of the workbook
$excelStructure = @(
    @{SheetName="DormantAccounts"; Columns=@("Name", "LastLogonDate", "DistinguishedName")},
    @{SheetName="Blank Last Logon Date"; Columns=@("Name", "LastLogonDate", "DistinguishedName")},
    @{SheetName="Expired Passwords"; Columns=@("Name", "PasswordLastSet", "DistinguishedName")},
    @{SheetName="Non-expiring passwords"; Columns=@("Name", "PasswordNeverExpires", "DistinguishedName")},
    @{SheetName="Blank Last Logon Date"; Columns=@("Name", "AccountExpirationDate", "DistinguishedName")}
)

# Create the workbook with pre-named sheets and columns
foreach ($sheet in $excelStructure) {
    # Create a dummy row with headers to set up the columns
    $dummyRow = @{}
    foreach ($col in $sheet.Columns) {
        $dummyRow.$col = $null
    }

    # Export the dummy row to create the sheet and columns
    $dummyRow | Export-Excel -Path $workbookPath -WorksheetName $sheet.SheetName -AutoSize 
}


# Function to get the last row index in an Excel worksheet
function Get-LastRow($worksheet) {
    $usedRange = $worksheet.UsedRange
    return $usedRange.Rows.Count
}

# Open the workbook (or create it if it doesn't exist)
if (-not (Test-Path $workbookPath)) {
    # If the workbook doesn't exist, create it with pre-named sheets and columns
    $excelStructure = @(
        @{SheetName="DormantAccounts"; Columns=@("Name", "LastLogonDate", "DistinguishedName")},
        @{SheetName="Blank Last Logon Date"; Columns=@("Name", "LastLogonDate", "DistinguishedName")},
        @{SheetName="Expired Passwords"; Columns=@("Name", "PasswordLastSet", "DistinguishedName")},
        @{SheetName="Non-expiring passwords"; Columns=@("Name", "PasswordNeverExpires", "DistinguishedName")},
        @{SheetName="Blank Last Logon Date"; Columns=@("Name", "AccountExpirationDate", "DistinguishedName")}
    )

    foreach ($sheet in $excelStructure) {
        $dummyRow = @{}
        foreach ($col in $sheet.Columns) {
            $dummyRow.$col = $null
        }
        $dummyRow | Export-Excel -Path $workbookPath -WorksheetName $sheet.SheetName -AutoSize -NoClobber
    }
}

# Function to support Data Append 
# Specifically gets a given worksheet
function Get-ExcelWorkSheet {
    [OutputType([OfficeOpenXml.ExcelWorksheet])]
    [cmdletBinding()]
    param (
        [OfficeOpenXml.ExcelPackage]  $ExcelDocument,
        [string] $Name
    )
    $Data = $ExcelDocument.Workbook.Worksheets | Where { $_.Name -eq $Name }
    return $Data
}


# Function to append data to an existing worksheet starting from the last row
function Append-DataToSheet {
    param (
        [string]$sheetName,
        [array]$data
    )

    # Get the last row index in the worksheet
    $lastRow = Get-LastRow (Open-ExcelPackage -Path $workbookPath | Get-ExcelWorksheet -WorksheetName $sheetName)
    $startRow = $lastRow + 1

    # Append data to the worksheet starting from the last row
    $data | Export-Excel -Path $workbookPath -WorksheetName $sheetName -StartRow $startRow -AutoSize
}

# Append data to the workbook
Append-DataToSheet -sheetName "DormantAccounts" -data ($ICRAFDormantAccountsAD | Select Name, LastLogonDate, DistinguishedName)
Append-DataToSheet -sheetName "Blank Last Logon Date" -data ($ICRAFAccountsNoLastActivity | Select Name, LastLogonDate, DistinguishedName)
Append-DataToSheet -sheetName "Expired Passwords" -data ($Expired Passwords | Select Name, PasswordLastSet, DistinguishedName)
Append-DataToSheet -sheetName "Non-expiring passwords" -data ($ICRAFAccounts_PasswordsNeverExpireAD | Select Name, PasswordNeverExpires, DistinguishedName)
Append-DataToSheet -sheetName "Blank Last Logon Date" -data ($expiredAccounts | Select Name, AccountExpirationDate, DistinguishedName)



# Send Email with attachment
Send-MailMessage -To $ICRAFReportRecipient -From $alertMailUserName -Subject $emailSubject -Body $emailBody -SmtpServer $smtpServer -Port 587 -UseSsl -Credential $mailCredential -Attachments $workbookPath

