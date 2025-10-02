<#

Mangano IT - ConnectWise Manage - Update Agreement Addition
Created by: Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param (
	[Parameter(Mandatory)][int]$AgreementId,
	[Parameter(Mandatory)][int]$AdditionId,
	[Parameter(Mandatory)][int]$Value,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$ContentType = 'application/json'
$Output = @()
$Log = ''

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## UPDATE AGREEMENT ADDITION WITH NEW VALUE ##

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

$ApiBody = @(
	@{
		op = 'replace'
		path = '/quantity'
		value = $Value
	}
)

$ApiArguments = @{
	Uri = "$CWMApiUrl/finance/agreements/$AgreementId/additions/$AdditionId"
	Method = 'PATCH'
	Body = ConvertTo-Json -InputObject $ApiBody -Depth 100
	ContentType = $ContentType
	Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
	UseBasicParsing = $true
}

try {
	$Log += "Attempting to update $AdditionId...`n"
	Invoke-WebRequest @ApiArguments | Out-Null
	$Log += "SUCCESS: Updated $AdditionId."
	Write-Warning "SUCCESS: Updated $AdditionId."
	$Result = $true
} catch {
	$Log += "ERROR: Unable to update $AdditionId.`nERROR DETAILS: " + $_
	Write-Error "Unable to update $AdditionId : $_"
	$Result = $false
}

## WRITE OUTPUT TO FLOW ##

$Output = @{
	Result = $Result
	Log = $Log
}

Write-Output $Output | ConvertTo-Json -Depth 100