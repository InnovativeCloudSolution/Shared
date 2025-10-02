<#

Mangano IT - Azure Active Directory - Check if License is Assigned to User
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$BearerToken,
	[string]$TenantUrl,
    [Parameter(Mandatory=$true)][string]$SkuPartNumber,
    [Parameter(Mandatory=$true)][string]$UserId
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''
$Result = $false

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    $Log += "Bearer token not supplied. Getting bearer token for $TenantUrl...`n`n"
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

$ApiArguments = @{
    Uri = "https://graph.microsoft.com/v1.0/users/$UserId/licenseDetails" + '?$select=id, skuId, skuPartNumber'
    Method = 'GET'
    Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken";'ConsistencyLevel'='eventual'}
    UseBasicParsing = $true
}

## PULL ALL LICENSES ##

try {
    $Log += "Attempting to pull all licenses assigned to $UserId...`n"
    $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
    $Log += "SUCCESS: License details have been pulled.`n`n"
} catch {
    $Log += "ERROR: License details not pulled.`nERROR DETAILS: " + $_
    Write-Error "License details not pulled : $_"
    $ApiResponse = $null
}

## FIND REQUESTED LICENSE ##

$Log += "Attempting to locate license $SkuPartNumber...`n"
foreach ($License in $ApiResponse.value) {
	if ($License.skuPartNumber -eq $SkuPartNumber) {
		$Log += "SUCCESS: License details located."
        Write-Warning "SUCCESS: License details located."
		$LocatedLicense = $License
        $Result = $true
        break
	}
}

## SEND DETAILS BACK TO FLOW ##

$Output = @{
    Result = $Result
    UserId = $UserId
	SkuPartNumber = $LocatedLicense.skuPartNumber
	Id = $LocatedLicense.id
	SkuId = $LocatedLicense.skuId
	Log = $Log
}

Write-Output $Output | ConvertTo-Json -Depth 100