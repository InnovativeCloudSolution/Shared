<#

Mangano IT - ConnectWise Manage - Find ConnectWise Manage Agreement
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param (
	[Parameter(Mandatory)][string]$CompanyName,
	[string]$AgreementType,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$ContentType = 'application/json'

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## GET AGREEMENT FROM COMPANY NAME ##

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

$ApiArguments = @{
	Uri = "$CWMApiUrl/finance/agreements"
	Method = 'GET'
	Body = @{ 
		conditions = "company/name = '$CompanyName' AND agreementStatus = 'Active'"
		fields = 'id,name,type,company,contact,agreementStatus'
	}
	ContentType = $ContentType
	Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
	UseBasicParsing = $true
}

# Specify agreement type if provided
if ($AgreementType -ne '') {
	$ApiArguments.Body.conditions += "type/name = $AgreementType"
}

## GET COMPANY ID ##

try { $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json } catch {}

# Send back agreement details

Write-Output $ApiResponse | ConvertTo-Json -Depth 100