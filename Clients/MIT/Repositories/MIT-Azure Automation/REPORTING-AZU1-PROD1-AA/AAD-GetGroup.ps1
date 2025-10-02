<#

Mangano IT - Azure Active Directory - Get Group
Created by: Gabriel Nugent
Version: 1.0.2

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$BearerToken,
	[string]$TenantUrl,
    [string]$DisplayName,
    [string]$GroupId
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

## GET GROUP ##

$ApiArguments = @{
    Uri = 'https://graph.microsoft.com/v1.0/groups'
    Method = 'GET'
    Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken";'ConsistencyLevel'='eventual'}
    UseBasicParsing = $true
}

# Filter by ID/display name if provided
if ($GroupId -ne '' -and $null -ne $GroupId) {
    $ApiArguments.Uri += "/$GroupId"
} elseif ($DisplayName -ne '' -and $null -ne $GroupId) {
    $ApiArguments.Uri += '?$search="' + "displayName:$DisplayName" + '"&$top=1'
}

# Fetch group

try {
    $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
    Write-Warning "SUCCESS: Fetched group."
    $Group += $ApiResponse.value[0]
} catch {
    Write-Error "Unable to fetch group : $_"
}

## WRITE OUTPUT TO FLOW ##

Write-Output $Group | ConvertTo-Json -Depth 100