<#

Mangano IT - Azure Active Directory - Get All Enterprise Applications (Service Principals)
Created by: Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$BearerToken,
	[string]$TenantUrl,
    [bool]$AppsWithCertificates = $false
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

$Output = @()

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    $Log += "Bearer token not supplied. Getting bearer token for $TenantUrl...`n`n"
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

## GET ALL GROUPS ##

$ApiUrl = 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=id,appId,displayName,keyCredentials,servicePrincipalType&$top=999'

while ($null -ne $ApiUrl) {
    $ApiArguments = @{
        Uri = $ApiUrl
        Method = 'GET'
        Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
        UseBasicParsing = $true
    }

    try {
        $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
        Write-Warning "SUCCESS: Fetched list of enterprise applications for $TenantUrl."
        $EnterpriseApplications += $ApiResponse.value
        $ApiUrl = $ApiResponse.'@odata.nextlink'
        Start-Sleep -Seconds 3
    } catch {
        Write-Error "Unable to fetch list of enterprise applications for $TenantUrl : $_"
    }
}

# List only enterprise applications with certificates
if ($AppsWithCertificates) {
    foreach ($EnterpriseApplication in $EnterpriseApplications) {
        if ($null -ne $EnterpriseApplication.keyCredentials -and $EnterpriseApplication.keyCredentials -ne @()) {
            $Output += $EnterpriseApplication
        }
    }
} else {
    $Output = $EnterpriseApplications
}

## WRITE OUTPUT TO FLOW ##

Write-Output $Output | ConvertTo-Json -Depth 4 -Compress