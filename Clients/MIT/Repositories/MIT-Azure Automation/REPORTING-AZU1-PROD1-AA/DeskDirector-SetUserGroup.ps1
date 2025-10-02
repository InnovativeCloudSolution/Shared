<#

Mangano IT - DeskDirector - Set User Group
Created by: Alex Williams
Maintained by: Gabriel Nugent
Version: 1.2

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory=$true)]
    [string]$UserId,
    [Parameter(Mandatory=$true)]
    [string]$CompanyId,
    [Parameter(Mandatory=$true)]
    [string]$TeamId
)

## SCRIPT VARIABLES ##

$UserKey = .\DeskDirector-GetUserKey.ps1

## SETUP REQUEST ##

$ApiBody = @{
    role = 'guest'
}

$ApiArguments = @{
    Uri = "https://portal.manganoit.com.au/api/v3/tech/companies/$CompanyId/teams/$TeamId/users/$UserId"
    Method = 'PUT'
    ContentType = "application/json"
    Headers = @{
        Authorization = "DdAccessToken $UserKey"
    }
    Body = $ApiBody | ConvertTo-Json
}

## ATTEMPT TO MAKE REQUEST ##

try {
    $Log += "Attempting to add $UserId to $TeamId...`n"
    Invoke-RestMethod @ApiArguments | Out-Null
    $Log += "SUCCESS: Added $UserId to $TeamId."
    Write-Warning "SUCCESS: Added $UserId to $TeamId."
    $Result = $true
} catch {
    $Log += "ERROR: Unable to add $UserId to $TeamId.`nERROR DETAILS: " + $_
    Write-Error "Unable to add $UserId to $TeamId : $_"
}

## SEND OUTPUT TO FLOW ##

$Output = @{
    Result = $Result
    Log = $Log
}

Write-Output $Output | ConvertTo-Json