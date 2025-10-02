<#

Mangano IT - Azure Active Directory - Get all Assigned Roles for User
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$BearerToken,
	[string]$TenantUrl,
    [Parameter(Mandatory=$true)][string]$UserId
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    $Log += "Bearer token not supplied. Getting bearer token for $TenantUrl...`n`n"
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

$ApiArguments = @{
    Uri = "https://graph.microsoft.com/v1.0/users/$UserId/appRoleAssignments"
    Method = 'GET'
    Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken";ConsistencyLevel='eventual'}
    UseBasicParsing = $true
}

## GET ALL LICENSES ASSIGNED TO USER ##

$Response = Invoke-WebRequest @ApiArguments | ConvertFrom-Json

## WRITE OUTPUT TO FLOW ##

Write-Output $Response.value | ConvertTo-Json -Depth 100