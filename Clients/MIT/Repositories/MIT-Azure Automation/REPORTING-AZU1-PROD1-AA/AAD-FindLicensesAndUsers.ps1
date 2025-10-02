<#

Mangano IT - Azure Active Directory - Find License Details and User List
Created by: Gabriel Nugent
Version: 1.1

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

## PULL LICENSE INFO ##

$License = .\AAD-FindLicenseDetails.ps1 -BearerToken $BearerToken -SkuPartNumber $SkuPartNumber | ConvertFrom-Json
$SkuId = $License.SkuId

$ApiArguments = @{
    Uri = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,assignedLicenses&$filter=assignedLicenses/any(u:u/skuId eq ' + "$SkuId)"
    Method = 'GET'
    Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
    UseBasicParsing = $true
}

## PULL USER LIST ##

try {
    $Log += "Attempting to pull all users with license $SkuPartNumber...`n"
    while ($null -ne $ApiArguments.Uri) {
        $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
        if ($ApiResponse.value) { $Users += $ApiResponse.value }
        $ApiArguments.Uri = $ApiResponse.'@odata.nextlink' # Repeat if there are more users
    }
    
    $Log += "SUCCESS: User details have been pulled.`n`n"
    Write-Warning "SUCCESS: User details have been pulled."
} catch {
    $Log += "ERROR: User details not pulled.`nERROR DETAILS: " + $_
    Write-Error "User details not pulled : $_"
    $ApiResponse = $null
}

## SEND OUTPUT TO FLOW ##

$Output = @{
    SkuPartNumber = $License.SkuPartNumber
	Id = $License.Id
	SkuId = $License.SkuId
	ConsumedUnits = $License.ConsumedUnits
    Users = $Users
	Log = $Log
}

Write-Output $Output | ConvertTo-Json -Depth 100