<#

Mangano IT - Azure Active Directory - Unassign Role from User (AAD)
Created by: Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
	[string]$BearerToken,
    [string]$TenantUrl,
    [string]$UserPrincipalName = '',
    [string]$UserId = '',
    [string]$AppRoleId = ''
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

$Result = $false

## CHECKS AND BALANCES FOR USER/GROUP DETAILS ##

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    $Log += "Bearer token not supplied. Getting bearer token for $TenantUrl...`n`n"
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

# If UserId is null, fetch UserId via Graph API request
if ($UserId -eq '') {
    $GetUserArguments = @{
        Uri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName"
        Method = 'GET'
        Headers = @{ 'Authorization'="Bearer $BearerToken";'ConsistencyLevel'='eventual' }
        UseBasicParsing = $true
    }
    
    try {
        $Log += "Attempting to get user details for $UserPrincipalName..."
        $GetUserResponse = Invoke-WebRequest @GetUserArguments | ConvertFrom-Json
        $UserId = $GetUserResponse.id
        $Log += "`nSUCCESS: Fetched user details for $UserPrincipalName.`nUser ID: $UserId`n`n"
        Write-Warning "User ID: $UserId"
    } catch {
        $Log += "`nERROR: Unable to fetch user details for $UserPrincipalName.`nERROR DETAILS: " + $_ + "`n`n"
        Write-Error "Unable to fetch user details for $UserPrincipalName : $_"
    }
}

$ApiArguments = @{
    Uri = "https://graph.microsoft.com/v1.0/users/$UserId/appRoleAssignments/$AppRoleId"
    Method = 'DELETE'
    Headers = @{ 'Content-Type'="application/json";'Authorization'="Bearer $BearerToken";'ConsistencyLevel'='eventual' }
    UseBasicParsing = $true
}

## ADD TO SECURITY GROUP ##

# Contacts Graph API to remove member from role
try {
    $ApiResponse = Invoke-WebRequest @ApiArguments
} catch {
    $Result = $false
    $Log += "ERROR: Unable to remove $UserId from $AppRoleId.`nERROR DETAILS:" + $_
    Write-Error "Unable to remove $UserId from $AppRoleId : $_"
}
if ($ApiResponse.StatusCode -eq 204) { 
    $Log += "SUCCESS: $UserId has been removed from $AppRoleId."
    Write-Warning "SUCCESS: $UserId has been removed from $AppRoleId."
    $Result = $true 
}

## SEND DETAILS BACK TO FLOW ##

$Output = @{
    Result = $Result
    AppRoleId = $AppRoleId
    UserId = $UserId
    Log = $Log
}

Write-Output $Output | ConvertTo-Json