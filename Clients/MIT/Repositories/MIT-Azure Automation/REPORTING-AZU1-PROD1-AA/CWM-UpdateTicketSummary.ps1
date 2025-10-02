<#

Mangano IT - ConnectWise Manage - Update Ticket Summary
Created by: Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow or an Azure Automation script.

#>

param (
    [Parameter(Mandatory=$true)][int]$TicketId,
    [Parameter(Mandatory=$true)][string]$Summary,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

[string]$Log = ''
$ContentType = 'application/json'

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## SETUP API VARIABLES ##

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

## FORMAT SUMMARY ##

# If the summary is too long, shorten the summary
if ($Summary.Length -gt 100) {
    $Summary = $Summary.Substring(0, 100)
}

## UPDATE SUMMARY ##

# Build API body
$ApiBody = @(
    @{
        op = "replace"
        path = "/summary"
        value = $Summary
    }
)

# Build API arguments
$ApiArguments = @{
    Uri = "$CWMApiUrl/service/tickets/$TicketId"
    Method = 'PATCH'
    ContentType = $ContentType
    Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
    Body = ConvertTo-Json -InputObject $ApiBody -Depth 100
    UseBasicParsing = $true
}

try {
    $Log += "Changing #$TicketId's summary to $Summary...`n"
    Invoke-WebRequest @ApiArguments | Out-Null
    $Log += "SUCCESS: #$TicketId's summary has been changed to $Summary."
    Write-Warning "SUCCESS: #$TicketId's summary has been changed to $Summary."
    $Result = $true
} catch {
    $Log += "ERROR: Unable to change #$TicketId's summary to $Summary.`nERROR DETAILS: " + $_
    Write-Error "Unable to change #$TicketId's summary to $Summary : $_"
    $Result = $false
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Result = $Result
    Log = $Log
}

Write-Output $Output | ConvertTo-Json