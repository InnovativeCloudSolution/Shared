<#

Mangano IT - ConnectWise Manage - Get Site
Created by: Gabriel Nugent
Version: 1.0.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param (
    [Parameter(Mandatory)][int]$CompanyId,
    [int]$SiteId,
    [string]$SiteName,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$ContentType = 'application/json'
$SiteObject = $false

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## TEMPORARY QHL OVERRIDE ##

if ($CompanyId -eq 3393 -and $SiteName -eq "BRIS1") {
    $SiteName = 'BNE1 - Brisbane (George Street)'
}

## SETUP API VARIABLES ##

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

$ApiArguments = @{
    Uri = "$($CWMApiUrl)/company/companies/$($CompanyId)/sites"
    Method = 'GET'
    ContentType = $ContentType
    Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
    UseBasicParsing = $true
}

# Add site ID if provided, or search for site
if ($SiteId -ne 0) {
    $ApiArguments.Uri += "/$($SiteId)"
    $SiteObject = $true
} elseif ($SiteName -ne '') {
    $ApiArguments += @{
        Body = @{
            conditions = "name like '$($SiteName)'"
            pageSize = 1000
        }
    }
}

## FETCH DETAILS FROM TICKET ##

try { 
    $Site = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
    Write-Warning "SUCCESS: Site grabbed."
} catch { 
    Write-Error "Site unable to be grabbed : $($_)"
}

## SEND DETAILS TO FLOW ##

if ($SiteObject) {
    Write-Output $Site | ConvertTo-Json -Depth 100
} else {
    Write-Output $Site[0] | ConvertTo-Json -Depth 100
}