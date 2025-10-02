<#

Mangano IT - Get User Key for DeskDirector
Created by: Gabriel Nugent
Version: 1.0.1

This runbook is designed to be used in conjunction with other Azure Automation scripts.

#>

## SCRIPT VARIABLES ##

$ApiKey = .\KeyVault-GetSecret.ps1 -SecretName 'DeskDirector-API'
$UserId = .\KeyVault-GetSecret.ps1 -SecretName 'DeskDirector-UserId'

## SETUP REQUEST ##

$ApiArguments = @{
    Uri = "https://manganoit.deskdirector.com/api/v2/user/member/1317/userkey"
    Method = 'GET'
    ContentType = "application/json"
    Headers = @{
        Authorization = "DdApi $ApiKey"
    }
}

## ATTEMPT TO FETCH USER KEY ##

try {
    Write-Warning "Fetching user key for user $UserId..."
    $UserKey = Invoke-RestMethod @ApiArguments
    Write-Warning "User key fetched for user $UserId."
} catch {
    Write-Error "Unable to fetch user key for $UserId : $_"
}

## SEND OUTPUT TO SCRIPT ##

Write-Output $UserKey.userKey