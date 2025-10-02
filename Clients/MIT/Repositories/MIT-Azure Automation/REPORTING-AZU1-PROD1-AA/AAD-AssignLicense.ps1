<#

Mangano IT - Azure Active Directory - Assign License to User
Created by: Gabriel Nugent
Version: 1.2

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

## GET LICENSE DETAILS ##

$License = .\AAD-FindLicenseDetails.ps1 -BearerToken $BearerToken -SkuPartNumber $SkuPartNumber | ConvertFrom-Json

## ASSIGN LICENSE ##

# Build request body
$ApiBody = @{
    addLicenses = @( 
        @{ skuId = $License.SkuId } 
    )
    removeLicenses = @()
}

$ApiArguments = @{
    Uri = "https://graph.microsoft.com/v1.0/users/$UserId/assignLicense"
    Method = 'POST'
    Body = ConvertTo-Json -InputObject $ApiBody -Depth 100
    UseBasicParsing = $true
    Headers = @{
        'Content-Type' = "application/json"
        Authorization = "Bearer $BearerToken"
        ConsistencyLevel = 'eventual'
    }
}

try {
    $Log += "Attempting to assign $SkuPartNumber to $UserId...`n"
    $Response = Invoke-WebRequest @ApiArguments
	if ($null -ne $Response) {
		$Log += "SUCCESS: $SkuPartNumber has been assigned to $UserId."
        Write-Warning "SUCCESS: $SkuPartNumber has been assigned to $UserId."
    	$Result = $true
	} else {
        $CheckAssigned = .\AAD-CheckIfLicenseAssigned.ps1 -BearerToken $BearerToken -SkuPartNumber $SkuPartNumber -UserId $UserId | ConvertFrom-Json
        if ($CheckAssigned.Result) {
            $Log += "SUCCESS: The request returned an error, but a separate check has confirmed that $SkuPartNumber is assigned to $UserId."
            Write-Warning "SUCCESS: The request returned an error, but a separate check has confirmed that $SkuPartNumber is assigned to $UserId."
            $Result = $true
        } else { throw }
    }
} catch {
    $Log += "ERROR: $SkuPartNumber has not been assigned to $UserId.`nERROR DETAILS: " + $_
    Write-Error "$SkuPartNumber has not been assigned to $UserId : $_"
    $Result = $false
}

## SEND DETAILS BACK TO FLOW ##

if ($License.ConsumedUnits -ge $License.EnabledUnits) {
	$LicenseAvailable = $false
} else { $LicenseAvailable = $true }

$Output = @{
    Result = $Result
    LicenseAvailable = $LicenseAvailable
    ConsumedUnits = $License.ConsumedUnits
	EnabledUnits = $License.EnabledUnits
	Log = $Log
}

Write-Output $Output | ConvertTo-Json -Depth 100