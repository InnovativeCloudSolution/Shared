<#

Mangano IT - ConnectWise Manage - Fetch Ticket Details
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param (
    [Parameter(Mandatory=$true)][int]$TicketId,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$ContentType = 'application/json'
[string]$Log = ''

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## SETUP API VARIABLES ##

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

$ApiArguments = @{
    Uri = "$CWMApiUrl/service/tickets/$TicketId"
    Method = 'GET'
    ContentType = $ContentType
    Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
    UseBasicParsing = $true
}

## FETCH DETAILS FROM TICKET ##

try { 
    $Log += "Fetching ticket details for #$TicketId...`n"
    $Ticket = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
    $Log += "SUCCESS: Ticket details fetched for #$TicketId."
    Write-Warning "SUCCESS: Ticket details fetched for #$TicketId."
} catch { 
    $Log += "ERROR: Ticket tasks unable to be fetched for #$TicketId.`nERROR DETAILS: " + $_ + "`n"
    Write-Error "Ticket tasks unable to be fetched for #$TicketId : $_"
    $Ticket = $null
}

## SEND DETAILS TO FLOW ##

Write-Output $Ticket | ConvertTo-Json -Depth 100