<#

Mangano IT - ConnectWise Manage - Find Contact by Email Address
Created by: Gabriel Nugent
Version: 1.0.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param (
	[Parameter(Mandatory=$true)][string]$CompanyName,
    [Parameter(Mandatory=$true)][string]$EmailAddress,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$ContentType = 'application/json'
$ApiResponse = $null

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## SETUP API VARIABLES ##

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

$ApiArguments = @{
    Uri = "$CWMApiUrl/company/contacts"
    Method = 'GET'
    Body = @{ 
        conditions = "company/name = '$CompanyName'"
        childConditions = "communicationItems/value like '$EmailAddress' AND communicationItems/communicationType = 'Email'"
    }
    ContentType = $ContentType
    Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
    UseBasicParsing = $true
}

## GET CONTACT DETAILS ##

try { $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json } catch {}

Write-Output $ApiResponse | ConvertTo-Json -Depth 100