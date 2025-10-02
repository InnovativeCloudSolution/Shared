<#

Mangano IT - Azure Active Directory - Delete MFA for AAD User
Created by: Gabriel Nugent
Version: 0.9

NOT CURRENTLY WORKING - BEARER TOKEN DOESN'T HAVE THE RIGHT PERMS

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$UserId,
    [string]$UserPrincipalName,
    [string]$BearerToken,
    [string]$TenantUrl
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

# Get user ID if not provided
if ($UserId -eq '') {
    $User = .\AAD-CheckUserExists.ps1 -UserPrincipalName $UserPrincipalName -BearerToken $BearerToken | ConvertFrom-Json
    $UserId = $User.UserId
    $Log += $User.Log + "`n`n"
}

## GET LIST OF MFA OPTIONS ##

$GetMfaArguments = @{
    Uri = "https://graph.microsoft.com/v1.0/users/$UserId/revokeSignInSessions"
    Method = 'POST'
    Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
}

try {
    $Log += "Fetching MFA options for $UserId...`n"
    $MfaOptions = Invoke-WebRequest @GetMfaArguments | ConvertFrom-Json
    $Log += "SUCCESS: Fetched MFA options for $UserId.`n`n"
    Write-Warning "SUCCESS: Fetched MFA options for $UserId."
    $Result = $true
} catch {
    $Log += "ERROR: Unable to fetch MFA options for $UserId.`nERROR DETAILS: "
    Write-Error "Unable to fetch MFA options for $UserId : $_"
    $Result = $false
}

## DELETE MFA OPTIONS ##

## SEND DETAILS TO FLOW ##

$Output = @{
    Result = $Result
    Log = $Log
}

Write-Output $Output | ConvertTo-Json -Depth 100