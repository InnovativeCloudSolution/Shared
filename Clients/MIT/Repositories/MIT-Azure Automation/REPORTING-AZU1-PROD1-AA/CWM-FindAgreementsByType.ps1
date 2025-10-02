<#

Mangano IT - ConnectWise Manage - Find ConnectWise Manage Agreement by Type
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param (
	[Parameter(Mandatory)][string]$AgreementTypes,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$ContentType = 'application/json'

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## GET AGREEMENT BY TYPE ##

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

foreach ($AgreementType in $AgreementTypes.split(';')) {
	$ApiArguments = @{
		Uri = "$CWMApiUrl/finance/agreements"
		Method = 'GET'
		Body = @{ 
			conditions = "type/name = '$AgreementType' AND agreementStatus = 'Active'"
			fields = 'id,name,type,company,contact,agreementStatus'
		}
		ContentType = $ContentType
		Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
		UseBasicParsing = $true
	}
	
	try { $ApiResponse += Invoke-WebRequest @ApiArguments | ConvertFrom-Json } catch {}
}

# Send back agreement details

Write-Output $ApiResponse | ConvertTo-Json -Depth 100