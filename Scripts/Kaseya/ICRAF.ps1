#Import AD Module
Import-Module ActiveDirectory

#ErrorPreference
#$ErrorActionPreference = “SilentlyContinue”

#VariableDeclaration
$UserCount = ""
$UserProc = 0
$AddCount = 0
$SkippedUser = 0

## Enter your AD Domain here
$Domain = 'dc=CIFOR-ICRAF,dc=AD'
 
## Enter the name of your Dynamic Security group here
$Groupname = 'ICRAF-O365-MFAUsers'
$GroupDN = 'CN=ICRAF-O365-MFAUsers,OU=Groups,OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG'
$GroupMember = Get-ADGroupMember -Identity $Groupname -Recursive | Select-Object -ExpandProperty samAccountName 

## Enter the OU's you want to search
$OUsToSearch = @(
  "OU=ICRAF Regions,OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG",
  "OU=ICRAF Kenya Employees,OU=ICRAF Kenya,OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG",
  "OU=ICRAF MFI,OU=ICRAF Kenya,OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG",
  "OU=ICRAF ICT,OU=ICRAF Kenya,OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG",
  "OU=ICRAF Administrators,OU=ICRAF Kenya,OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG",
  "OU=ICRAF Kenya Students-New,OU=ICRAF Kenya,OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG",
  "OU=ICRAF Kenya Consultants and Temporary Staff,OU=ICRAF Kenya,OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG",
  "OU=ICRAF Kenya Hosted Organizations,OU=ICRAF Kenya,OU=ICRAFHUB,DC=CIFOR-ICRAF,DC=ORG"
)

# Create empty array 
$Users = @()
 
# Loop through OUs and search for users
foreach($Path in $OUsToSearch){
    $Users += Get-ADUser -SearchBase $Path -Filter *
    $Users.count
    #$UserCount += (Get-ADUser -SearchBase $Path -Filter *).count
    #$TotalCount += $UserCount
       
}

# Loop through OUs and search for users
foreach($User in $Users)
{
if ($GroupMember -contains $User.SamAccountName){
$SkippedUser++
}
else
{
Add-ADGroupMember -Identity $groupname -members $User.SamAccountName -ErrorAction SilentlyContinue
$AddCount++
}
}

#Remove Contact Objects From Group

$ContactObj = Get-ADObject -filter {(memberof -eq $GroupDN) -and (objectClass -eq "contact")}  

Foreach ($cObj in $ContactObj)
{
Set-ADGroup -Identity $GroupName -Remove @{'member'=$cObj.DistinguishedName}
}

#Generate Log
$UserCount = $Users.count
$GroupMemberCount = $GroupMember.count


$LogReport = @"
No of Users from AD Group = $UserCount
Users Added to MFA Group = $AddCount
Total Users in MFA Group = $GroupMemberCount
"@

$LogReport | Out-File 'C:\Kworking\MFALog\MFALog-ICRAF.txt'

#$SpaLog = New-Object psobject
#$SpaLog | Add-Member -MemberType NoteProperty -Name "No of Users from AD Group" -Value $Users.Count
#$SpaLog | Add-Member -MemberType NoteProperty -Name "Users Added to Spanning Group" -Value $AddCount
#$SpaLog | Add-Member -MemberType NoteProperty -Name "Total Users in Spanning Group" -Value $GroupMember.Count
#$SpaLog | Export-Csv -Path 'C:\KWorking\SpanningLog\SpanningLog-ILRI.csv' -NoTypeInformation



