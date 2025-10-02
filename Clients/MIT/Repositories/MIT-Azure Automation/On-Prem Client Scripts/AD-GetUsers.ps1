<#

Mangano IT - Active Directory - Get Users
Created by: Gabriel Nugent
Version: 1.5

This runbook is designed to be run from a Power Automate flow.

#>

param(
    [Parameter(Mandatory)][string]$OUPaths, # Separated by ; for multiple paths
    [bool]$AdminSearch = $false
)

## SCRIPT VARIABLES ##

$Log = ''
$Accounts = @()
$AccountsUpdated = @()

## GET ACCOUNTS FROM EACH PROVIDED OU ##

foreach ($OUPath in $OUPaths.split(';')) {
    try {
        $Log += "Looking for users in $OUPath...`n" 
        $Accounts += Get-ADUser -Filter * -SearchBase $OUPath -ResultPageSize 0 -Prop lastLogonTimestamp, MemberOf, Manager | `
        Select-Object Name, Enabled, @{n="lastLogonDate";e={[datetime]::FromFileTime($_.lastLogonTimestamp)}}, MemberOf, DistinguishedName, `
        UserPrincipalName, GivenName, Surname, Manager, SamAccountName
        $Log += "SUCCESS: Located users in $OUPath.`n`n"
        Write-Warning "SUCCESS: Located users in $OUPath."
    } catch {
        $Log += "ERROR: Unable to look for users in $OUPath.`nERROR DETAILS: " + $_
        Write-Error "Unable to look for users in $OUPath : $_"
        $Accounts = $null
    }
}

if ($AdminSearch) {
    try {
        $Log += "Looking for users that are not in $OUPath...`n" 
        $UserAccounts = Get-ADUser -Filter * -ResultPageSize 0 -Prop MemberOf, Manager | Where-Object { $_.DistinguishedName -notlike "*$OUPath*" }| `
        Select-Object Name, Enabled, MemberOf, DistinguishedName, UserPrincipalName, GivenName, Surname, Manager, SamAccountName
        $Log += "SUCCESS: Located users not in $OUPath.`n`n"
        Write-Warning "SUCCESS: Located users not in $OUPath."
    } catch {
        $Log += "ERROR: Unable to look for users in $OUPath.`nERROR DETAILS: " + $_
        Write-Error "Unable to look for users in $OUPath : $_"
        $UserAccounts = $null
    }
} else {$UserAccounts = $null}

## SORT THROUGH ACCOUNTS ##

foreach ($User in $Accounts) {
	# Clear variables
	$UserUpdated = $null
	$NonAdminUser = $null

    # Replace null values with empty/filler strings
    if ($null -eq $User.lastLogonDate) { $LastLogonDate = "N/A" }
    else { $LastLogonDate = $User.lastLogonDate }
    if ($null -eq $User.UserPrincipalName) { $UserPrincipalName = '' } else { $UserPrincipalName = $User.UserPrincipalName }
    if ($null -eq $User.GivenName) { $GivenName = '' } else { $GivenName = $User.GivenName }
    if ($null -eq $User.Surname) { $Surname = '' } else { $Surname = $User.Surname }
    $SamAccountName = $User.SamAccountName

    # Get a friendly name for the manager
    if ($null -eq $User.Manager) { $Manager = '' }
    else { $Manager = $User.Manager.Substring(3, $User.Manager.IndexOf(',') - 3) }
        
    # Get a friendly name for each group user is a memberof
    $GroupList = @()
    foreach ($Group in $User.MemberOf) { $GroupList += (Get-ADGroup $Group).name }  
    $GroupListString = $GroupList -join ", " # Convert group object to string

    # Get path from DN
    $Index = $User.distinguishedName.IndexOf("OU=")
    $Path = $User.distinguishedName.Substring($Index, $User.distinguishedName.Length - $Index)

    # Create object for export
    $UserUpdated = @{
        'Name' = $User.name
        'UserPrincipalName' = $UserPrincipalName
        'SamAccountName' = $SamAccountName
        'GivenName' = $GivenName
        'Surname' = $Surname
        'Enabled' = $User.enabled
        'LastLogonDate' = $LastLogonDate
        'Manager' = $Manager
        'MemberOf' = $GroupListString
        'OUPath' = $Path
        'ADUser' = "N/A"
    }

    # If admin search, check for non-admin user
    if ($null -ne $UserAccounts) {
        foreach ($NonAdminUser in $UserAccounts) {
            if ($NonAdminUser.GivenName -eq $GivenName -and $NonAdminUser.Surname -eq $Surname) {
                $UserUpdated.ADUser = $NonAdminUser.Enabled.ToString();
                break
            }
        }
    }

    # Add user to output
    $AccountsUpdated += $UserUpdated
    $GroupList = $null
}

## SEND DATA TO FLOW ##

Write-Output $AccountsUpdated | ConvertTo-Json -Depth 100