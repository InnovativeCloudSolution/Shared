<#

Mangano IT - Azure Active Directory - Find License Details for Provided SKU Part Number
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$BearerToken,
	[string]$TenantUrl,
    [string]$SkuPartNumber
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
    Uri = 'https://graph.microsoft.com/v1.0/subscribedSkus/?$select=id,skuId,skuPartNumber,consumedUnits,prepaidUnits'
    Method = 'GET'
    Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
	UseBasicParsing = $true
}

## PULL ALL LICENSES ##

try {
    $Log += "Attempting to pull all licenses for tenant...`n"
    $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
    $Log += "SUCCESS: License details have been pulled.`n`n"
	Write-Warning "SUCCESS: License details have been pulled."
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
		$LocatedLicense = $License
	}
}

## SEND DETAILS BACK TO FLOW ##

if ($LocatedLicense.consumedUnits -ge $LocatedLicense.prepaidUnits.enabled) {
	$LicenseAvailable = $false
} else { $LicenseAvailable = $true }

$Output = @{
	SkuPartNumber = $LocatedLicense.skuPartNumber
	LicenseAvailable = $LicenseAvailable
	Id = $LocatedLicense.id
	SkuId = $LocatedLicense.skuId
	ConsumedUnits = $LocatedLicense.consumedUnits
	EnabledUnits = $LocatedLicense.prepaidUnits.enabled
	SuspendedUnits = $LocatedLicense.prepaidUnits.suspended
	WarningUnits = $LocatedLicense.prepaidUnits.warning
	Log = $Log
}

Write-Output $Output | ConvertTo-Json -Depth 100