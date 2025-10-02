<#

Mangano IT - Active Directory - Search for Users
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be run from a Power Automate flow.

#>

param(
    [Parameter(Mandatory)]$Users,
    [Parameter(Mandatory)][string]$OUPath
)

## GET USERS ##

foreach($User in $Users) {
    $MatchedUser = .\AD-GetUserByFirstAndLastName.ps1 -GivenName $User.GivenName -Surname $User.Surname | ConvertFrom-Json
    if ($null -ne $MatchedUser) { $User += @{ ADUser = $MatchedUser.User.Enabled.ToString(); } }
    else { $User += @{ ADUser = "N/A" } }
    $UpdatedUsers += $User
}

## WRITE BACK TO FLOW ##

Write-Output $UpdatedUsers | ConvertTo-Json -Depth 100