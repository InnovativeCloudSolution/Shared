<#

Mangano IT - ConnectWise Manage - Find Contact from Ticket ID
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
$ApiResponse = $null

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## GET TICKET ##

$Ticket = .\CWM-FindTicketDetails.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets | ConvertFrom-Json
$ContactId = $Ticket.contact.id

## SETUP API VARIABLES ##

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

$ApiArguments = @{
    Uri = "$CWMApiUrl/company/contacts/$ContactId"
    Method = 'GET'
    ContentType = $ContentType
    Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
    UseBasicParsing = $true
}

## GET CONTACT DETAILS ##

try { $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json } catch {}

Write-Output $ApiResponse | ConvertTo-Json -Depth 100