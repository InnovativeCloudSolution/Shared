<#

Mangano IT - Azure Active Directory - Set Usage Location (AAD)
Created by: Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
	[string]$BearerToken,
    [string]$TenantUrl,
    [Parameter(Mandatory=$true)][string]$UserPrincipalName = '',
    [Parameter(Mandatory=$true)][string]$UsageLocation
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

## CHECKS AND BALANCES FOR USER/GROUP DETAILS ##

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    $Log += "Bearer token not supplied. Getting bearer token for $TenantUrl...`n`n"
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

## SCRIPT VARIABLES ##

$Result = $false

$ApiArguments = @{
    Uri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName"
    Method = 'PATCH'
    Headers = @{ 'Content-Type'="application/json";'Authorization'="Bearer $BearerToken";'ConsistencyLevel'='eventual' }
    Body = @{ "usageLocation" = $UsageLocation } | ConvertTo-Json
    UseBasicParsing = $true
}

## SET USAGE LOCATION ##

# Contacts Graph API to add member to group - only runs if user was not already found
try {
    $Log += "Setting usage location to $UsageLocation for $UserPrincipalName...`n"
    Invoke-WebRequest @ApiArguments | Out-Null
    $Result = $true
    $Log += "SUCCESS: Set usage location to $UsageLocation for $UserPrincipalName."
    Write-Warning "SUCCESS: Set usage location to $UsageLocation for $UserPrincipalName."
} catch {
    $Result = $false
    $Log += "ERROR: Unable to set usage location to $UsageLocation for $UserPrincipalName.`nERROR DETAILS:" + $_
    Write-Error "Unable to set usage location to $UsageLocation for $UserPrincipalName. : $_"
}

## SEND DETAILS BACK TO FLOW ##

$Output = @{
    Result = $Result
    Log = $Log
}

Write-Output $Output | ConvertTo-Json