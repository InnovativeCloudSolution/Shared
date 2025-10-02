<#

Mangano IT - Active Directory - Get User by First and Last Name
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be run from a Power Automate flow.

#>

param(
    [Parameter(Mandatory)][string]$OUPath,
    [Parameter(Mandatory)][string]$GivenName,
    [Parameter(Mandatory)][string]$Surname
)

## SCRIPT VARIABLES ##

$Log = ''

## GET ACCOUNTS FROM EACH PROVIDED OU ##

try {
    $Log += "Looking for $GivenName $Surname in $OUPath...`n" 
    $User = Get-ADUser -Filter "GivenName -like '$GivenName' -and Surname -like '$Surname'" -SearchBase $OUPath -ResultSetSize 1 `
    -Prop lastLogonTimestamp, MemberOf, Manager | Select-Object Name, Enabled, @{n="lastLogonDate";e={[datetime]::FromFileTime($_.lastLogonTimestamp)}}, `
    MemberOf, DistinguishedName, UserPrincipalName, GivenName, Surname, Manager, SamAccountName
    $Log += "SUCCESS: Located $GivenName $Surname in $OUPath.`n`n"
    Write-Warning "SUCCESS: Located $GivenName $Surname in $OUPath."
} catch {
    $Log += "ERROR: Unable to look for users in $OUPath.`nERROR DETAILS: " + $_
    Write-Error "Unable to look for users in $OUPath : $_"
    $User = $null
}

## SORT THROUGH ACCOUNT DETAILS ##

# Replace null values with empty/filler strings
if ($null -eq $User.lastLogonDate) { $LastLogonDate = "N/A" }
else { $LastLogonDate = $User.lastLogonDate }
if ($null -eq $User.UserPrincipalName) { $User.UserPrincipalName = '' }
if ($null -eq $User.GivenName) { $User.GivenName = '' }
if ($null -eq $User.Surname) { $User.Surname = '' }

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
$UserUpdated += @{
    'Name' = $User.name
    'UserPrincipalName' = $User.UserPrincipalName
    'SamAccountName' = $User.SamAccountName
    'GivenName' = $User.GivenName
    'Surname' = $User.Surname
    'Enabled' = $User.enabled
    'LastLogonDate' = $LastLogonDate
    'Manager' = $Manager
    'MemberOf' = $GroupListString
    'OUPath' = $Path
}

## SEND DATA TO FLOW ##

$Output = @{
    User = $UserUpdated
    Log = $Log
}

Write-Output $Output | ConvertTo-Json -Depth 100