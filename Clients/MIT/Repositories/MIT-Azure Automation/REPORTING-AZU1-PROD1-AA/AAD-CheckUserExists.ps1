<#

Mangano IT - Azure Active Directory - Check User Exists in AAD
Created by: Gabriel Nugent
Version: 1.6.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$BearerToken,
	[string]$TenantUrl,
    [string]$UserId,
    [string]$UserPrincipalName,
    [bool]$SkipExtraAttempts = $false
)

## SCRIPT VARIABLES ##

$UserId = 'null'

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
    Uri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName" + '?$select=id,userPrincipalName,accountEnabled'
    Method = 'GET'
    Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
    UseBasicParsing = $true
}

# Replace URL if UserId provided
if ($UserId -ne '' -and $UserPrincipalName -eq '') {
    $ApiArguments.Uri = "https://graph.microsoft.com/v1.0/users/$UserId" + '?$select=id,userPrincipalName,accountEnabled'
}

$Result = $false
$AccountEnabled = $false
$Attempts = 0

# Change number of attempts if skipping
if ($SkipExtraAttempts) {
    $Attempts = 19
}

## CHECK IF USER EXISTS ##

while ($UserId -eq 'null' -and $Attempts -lt 20) {
    try {
        $Log += "Attempting to locate user...`n"
        $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
        $Result = $true
        $UserId = $ApiResponse.id
        $AccountEnabled = $ApiResponse.accountEnabled
        $Log += "SUCCESS: $UserPrincipalName ($UserId) has been located."
        Write-Warning "SUCCESS: $UserPrincipalName ($UserId) has been located."
    } catch {
        $Log += "INFO: $UserPrincipalName not located. Waiting 30 seconds before trying again.`n"
        Write-Warning "INFO: $UserPrincipalName not located. Waiting 30 seconds before trying again."
        $Attempts += 1
        $UserId = 'null'
        Start-Sleep -Seconds 30
    }
}

if ($UserId -eq 'null') {
    $Log += "`nERROR: $UserPrincipalName not located."
    Write-Error "$UserPrincipalName not located."
}

## SEND DETAILS BACK TO FLOW ##

$Output = @{
    Result = $Result
    UserId = $UserId
    UserPrincipalName = $UserPrincipalName
    AccountEnabled = $AccountEnabled
    Log = $Log
}

Write-Output $Output | ConvertTo-Json