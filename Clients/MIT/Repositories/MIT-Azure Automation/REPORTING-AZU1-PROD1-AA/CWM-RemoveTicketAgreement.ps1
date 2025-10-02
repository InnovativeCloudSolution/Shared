<#

Mangano IT - ConnectWise Manage - Remove Agreement from Ticket
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow or an Azure Automation script.

#>

param (
    [Parameter(Mandatory=$true)][int]$TicketId,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

# Track status of automation
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

# Setup body
$ApiBody = @(
    @{
        op = 'remove'
        path = 'agreement'
    }
)

$ApiArguments = @{
    Uri = "$CWMApiUrl/service/tickets/$TicketId"
    Method = 'PATCH'
    ContentType = $ContentType
    Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
    Body = ConvertTo-Json -InputObject $ApiBody -Depth 100
    UseBasicParsing = $true
}

## REMOVE AGREEMENT ##

try {
    $Log += "Updating #$TicketId...`n"
    Invoke-WebRequest @ApiArguments | Out-Null
    $Log += "SUCCESS: #$TicketId has been updated."
    Write-Warning "SUCCESS: #$TicketId has been updated."
    $Result = $true
} catch {
    $Log += "ERROR: Unable to update #$TicketId.`nERROR DETAILS: " + $_
    Write-Error "Unable to update #$TicketId : $_"
    $Result = $false
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Result = $Result
    Log = $Log
}

Write-Output $Output | ConvertTo-Json