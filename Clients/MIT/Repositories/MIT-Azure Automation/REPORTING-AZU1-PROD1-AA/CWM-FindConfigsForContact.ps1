<#

Mangano IT - ConnectWise Manage - Find Configurations for Contact
Created by: Gabriel Nugent
Version: 1.2.2

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param (
    [int]$ContactId,
	[string]$CompanyName,
    [string]$FirstName,
    [string]$LastName,
    [string]$EmailAddress,
    [bool]$ParentOnly = $true,
    [bool]$ActiveOnly = $true,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$ContentType = 'application/json'
$ApiResponse = $null
$Output = @()

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## GET CONTACT IF NOT PROVIDED ##

if ($ContactId -eq 0) {
    if ($EmailAddress -ne '') {
        $ContactParams = @{
            EmailAddress = $EmailAddress
            CompanyName = $CompanyName
            ApiSecrets = $ApiSecrets
        }
        $Contact = .\CWM-FindContactByEmail.ps1 @ContactParams | ConvertFrom-Json
    } else {
        $ContactParams = @{
            FirstName = $FirstName
            LastName = $LastName
            CompanyName = $CompanyName
            ApiSecrets = $ApiSecrets
        }
        $Contact = .\CWM-FindContact.ps1 @ContactParams | ConvertFrom-Json
    }
    
    $ContactId = $Contact.id
}

## SETUP API VARIABLES ##

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

$ApiArguments = @{
    Uri = "$CWMApiUrl/company/configurations"
    Method = 'GET'
    Body = @{
        pageSize = 1000
        conditions = "contact/id = $ContactId"
    }
    ContentType = $ContentType
    Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
    UseBasicParsing = $true
}

# Add filter if only active configs
if ($ActiveOnly) {
    $ApiArguments.Body.conditions += " AND status/name = 'Active'"
}

## GET CONFIG DETAILS ##

try {
    Write-Warning "Fetching configs for $ContactId..."
    $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
    Write-Warning "SUCCESS: Fetched configs for $ContactId."
} catch {
    Write-Error "Unable to find configs for $ContactId : $_"
}

## ORGANISE IF NECESSARY ##

# Add config to the list if it's a parent config
if ($ParentOnly) {
    foreach ($Config in $ApiResponse) {
        if ($null -eq $Config.parentConfigurationId) {
            $Output += @($Config)
        }
    }
} else {
    $Output = $ApiResponse
}

Write-Output $Output | ConvertTo-Json -Depth 100