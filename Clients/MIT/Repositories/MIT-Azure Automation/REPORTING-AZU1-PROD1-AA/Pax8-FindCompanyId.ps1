<#

Mangano IT - Pax8 - Find Company ID
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$BearerToken,
    [Parameter(Mandatory=$true)][string]$CompanyName
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    $Log += "Bearer token not supplied. Getting bearer token for Pax8...`n`n"
	$BearerToken = .\Pax8-GetBearerToken.ps1
}

$ContentType = 'application/json'
$CompanyId = ''
$ApiUrl = .\KeyVault-GetSecret.ps1 -SecretName 'Pax8-ApiUrl'

## GET LIST OF COMPANIES ##

$ApiArguments = @{
	Uri = "$ApiUrl/companies"
	Method = 'GET'
    ContentType = $ContentType
    Headers = @{'Authorization'="Bearer $BearerToken"}
	Body = @{
        status = "Active"
        size = 200
    }
    UseBasicParsing = $true
}

# Fetch list of companies to find company ID
try {
    $Log += "Fetching companies from Pax8...`n"
    $Companies = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
    $Log += "SUCCESS: Fetched companies from Pax8.`n"
}
catch {
    $Log += "ERROR: Unable to fetch companies from Pax8.`nERROR DETAILS: " + $_
    Write-Error "Unable to fetch companies from Pax8 : $_"
}

# Find matching company ID
foreach ($Company in $Companies.content) {
    if ($Company.name -eq $CompanyName) {
        $CompanyId = $Company.id
        $Log += "INFO: Located company ID in Pax8 for $CompanyName. Company ID: $CompanyId"
        Write-Warning "Company ID: $CompanyId"
    }
}

## SEND DETAILS TO FLOW ##

$Output = @{
    CompanyId = $CompanyId
    Log = $Log
}

Write-Output $Output | ConvertTo-Json