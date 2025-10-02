<#

Mangano IT - ConnectWise Manage - Rename Configurations in ConnectWise Manage
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be run independently.

#>

param(
    [Parameter(Mandatory)][int]$CompanyId,
    [Parameter(Mandatory)][string]$CharactersToReplace,
    [Parameter(Mandatory)][string]$CharactersToUse,
    [string]$StartMiddleOrEnd = 'Middle',
	$ApiSecrets = $null
)

## SCRIPT VARIABLES ##

[string]$Log = ''
$ContentType = 'application/json'
$MaxConfigsInOneRequest = 1000

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## SETUP API VARIABLES ##

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

## GET CONFIGS FOR COMPANY ##

$StartMiddleOrEnd = $StartMiddleOrEnd.ToLower();

switch($StartMiddleOrEnd) {
    'start' {
        $Log += "Given string will be searched for at the start of the config name.`n`n"
        $ApiBody = @{
            conditions = "company/id=$CompanyId AND name like '$CharactersToReplace*'"
            pageSize = 1000
            page=1
        }
    }
    'middle' {
        $Log += "Given string will be searched for anywhere in the config name.`n`n"
        $ApiBody = @{
            conditions = "company/id=$CompanyId AND name like '*$CharactersToReplace*'"
            pageSize=1000
            page=1
        }
    }
    'end' {
        $Log += "Given string will be searched for at the end of the config name.`n`n"
        $ApiBody = @{
            conditions = "company/id=$CompanyId AND name like '*$CharactersToReplace'"
            pageSize=1000
            page=1
        }
    }
}

$GetConfigsArguments = @{
    Uri = "$CWMApiUrl/company/configurations"
    Method = 'GET'
    Body = $ApiBody
    ContentType = $ContentType
    Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
    UseBasicParsing = $true
}

try {
    $Log += "Fetching configs for company $CompanyId...`n"
    $Configs = Invoke-WebRequest @GetConfigsArguments | ConvertFrom-Json
    $Log += "SUCCESS: Fetched configs for company $CompanyId.`n"
    while ($Configs.Count -eq $MaxConfigsInOneRequest) {
        $Log += "Array count suggests there are more configs to fetch. Fetching configs for company $CompanyId...`n"
        $GetConfigsArguments.Body.page += 1
        $Configs += Invoke-WebRequest @GetConfigsArguments | ConvertFrom-Json
        $Log += "SUCCESS: Fetched more configs for company $CompanyId.`n"
    }
    $Log += "INFO: Config count: " + $Configs.Count + "`n`n"
} catch {
    $Log += "Unable to fetch configs for $CompanyId.`nERROR DETAILS: " + $_
    $Configs = $null
}

## UPDATE CONFIGS ##

foreach ($Config in $Configs) {
	# Update config details to use the new name
	$ConfigId = $Config.id
    $ConfigName = $Config.name
    $NewConfigName = $ConfigName.Replace($CharactersToReplace, $CharactersToUse)

    $ApiBody = @(
        @{
            op = 'replace'
            path = '/name'
            value = $NewConfigName
        }
    )

    $UpdateConfigsArguments = @{
        Uri = "$CWMApiUrl/company/configurations/$ConfigId"
        Method = 'PATCH'
        Body = ConvertTo-Json -InputObject $ApiBody -Depth 100
        ContentType = $ContentType
        Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
        UseBasicParsing = $true
    }

	# Push config over the top of the old one
    try {
        $Log += "Replacing $ConfigName with $NewConfigName...`n"
        Invoke-WebRequest @UpdateConfigsArguments | Out-Null
        $Log += "SUCCESS: Replaced $ConfigName with $NewConfigName...`n`n"
    } catch {
        $Log += "Unable to replace $ConfigName with $NewConfigName.`nERROR DETAILS: " + $_ + "`n`n"
    }
}

## SHOW OUTPUT ##

Write-Output $Log