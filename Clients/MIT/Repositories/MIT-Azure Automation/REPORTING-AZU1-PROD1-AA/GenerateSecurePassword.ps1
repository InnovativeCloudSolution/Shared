<#

Mangano IT - Generate Secure Password
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with an Azure Automation runbook.

#>
param (
    $Length = 12,
    $DinopassCalls = 1
)

$DinopassUrl = "https://www.dinopass.com/password/strong"

Function Test-PasswordForDomain {
    param ([Parameter(Mandatory=$true)][string]$Password)

    if ($Password.Length -lt $Length) { return $false }
    if ($Password -like "*+*") { return $false }

    if (($Password -cmatch "[A-Z\p{Lu}\s]") -and ($Password -cmatch "[a-z\p{Ll}\s]")) { 
        return $true
    } else {
        return $false
    }
}

# Create password
do {
    # Fetch password from Dinopass
    $Password = Invoke-RestMethod -Uri $DinopassUrl

    # Add extra Dinopass calls if requested
    for ($i = 1; $i -lt $DinopassCalls; $i++) { 
        $Password += '-'
        $Password += Invoke-RestMethod -Uri $DinopassUrl
    }

    # Test password to see if it'll be allowed
    $CheckPassword = Test-PasswordForDomain $Password
} while ($CheckPassword -eq $False)

# Send back to outer script
Write-Output $Password