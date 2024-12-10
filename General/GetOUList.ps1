# Import the Active Directory module
Import-Module ActiveDirectory

function Is-Student {
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

# Define the suffixes array
$suffixes = @("Students", "Interns")

# Initialize an empty array to store matching OUs
$matchingOUs = @()

# Loop through each suffix
foreach ($suffix in $suffixes) {
    # Find OUs that match the current suffix
    $matchingOUs += Get-ADOrganizationalUnit -Filter * | Where-Object { $_.Name -like "*$suffix" }
}

# Display results
$matchingOUs | Select-Object Name, DistinguishedName | Format-Table -AutoSize

# Fetch a user object
$user = Get-ADUser -Identity "bobaga" -Properties DistinguishedName

# Call the function to check if the user is a student
if (Is-Student -User $user -MatchingOUs $matchingOUs) {
    Write-Output "The user is part of a student OU."
} else {
    Write-Output "The user is NOT part of a student OU."
}