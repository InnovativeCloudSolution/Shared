<#

Mangano IT - Azure Active Directory - Unassign Licenses from User
Created by: Gabriel Nugent
Version: 1.7.6

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$BearerToken,
	[string]$TenantUrl,
    [Parameter(Mandatory=$true)][string]$UserId,
    [string]$UserPrincipalName
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''
$Result = $false
$LicenseList = ''

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    $Log += "Bearer token not supplied. Getting bearer token for $TenantUrl...`n`n"
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

## GET ALL LICENSES ASSIGNED TO USER ##

$Licenses = .\AAD-FindUsersAssignedLicenses.ps1 -BearerToken $BearerToken -UserId $UserId | ConvertFrom-Json

## UNASSIGN LICENSES ##

# Build request body
$ApiBody = @{
    addLicenses = @()
    removeLicenses = @()
}

# Add licenses to API body and make list
foreach ($License in $Licenses) {
    $SkuPartNumber = $License.skuPartNumber
    $SkuId = $License.skuId
    $LicenseList += "`n- $SkuPartNumber ($SkuId)"
    $ApiBody.removeLicenses += $SkuId
}

# Remove licenses from user
if ($ApiBody.removeLicenses -ne @() -and $ApiBody.removeLicenses -ne @($null) -and $null -ne $ApiBody.removeLicenses) {
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

    # Remove licenses - built-in retry for web server errors
    try {
        $Log += "Attempting to unassign licenses from $UserPrincipalName...`n"
        Write-Warning "Attempting to unassign licenses from $UserPrincipalName..."
        do { $Response = Invoke-WebRequest @ApiArguments } until ( $Response.StatusCode -ne 503 )
        if ($null -ne $Response) {
            $Log += "SUCCESS: Licenses have been unassigned from $UserPrincipalName."
            Write-Warning "SUCCESS: Licenses have been unassigned from $UserPrincipalName."
            $Result = $true
        }
    } catch {
        $Log += "ERROR: Licenses have not been unassigned from $UserPrincipalName.`nERROR DETAILS: " + $_
        Write-Error "Licenses have not been unassigned from $UserPrincipalName : $_"
        $Result = $false
    }
} else {
    $Log += "INFO: No licenses to unassign for $UserPrincipalName."
    Write-Warning "INFO: No licenses to unassign for $UserPrincipalName."
    $LicenseList = ''
    $Result = $true
}

## SEND DETAILS BACK TO FLOW ##

$Output = @{
    Result = $Result
    LicenseList = $LicenseList
	Log = $Log
}

Write-Output $Output | ConvertTo-Json -Depth 100