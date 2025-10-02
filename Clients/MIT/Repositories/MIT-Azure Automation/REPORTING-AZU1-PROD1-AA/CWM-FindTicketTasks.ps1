<#

Mangano IT - ConnectWise Manage - Fetch Ticket Tasks
Created by: Gabriel Nugent
Version: 1.3.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param (
    [Parameter(Mandatory=$true)][int]$TicketId,
    [string]$TaskNotes,
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

$GetTasksArguments = @{
    Uri = "$CWMApiUrl/service/tickets/$TicketId/tasks?pageSize=1000"
    Method = 'GET'
    ContentType = $ContentType
    Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
    UseBasicParsing = $true
}

# If note is supplied, update Uri
if ($TaskNotes -ne '') {
    $GetTasksArguments.Uri += '&conditions=notes like "*' + $TaskNotes + '*"'
}

## FETCH TASKS FROM TICKET ##

try { 
    $Log += "Fetching ticket tasks for #$TicketId...`n"
    $Tasks = Invoke-WebRequest @GetTasksArguments | ConvertFrom-Json
    $Log += "SUCCESS: Ticket tasks fetched for #$TicketId."
    Write-Warning "SUCCESS: Ticket tasks fetched for #$TicketId."
} catch { 
    $Log += "ERROR: Ticket tasks unable to be fetched for #$TicketId.`nERROR DETAILS: " + $_
    Write-Error "Ticket tasks unable to be fetched for #$TicketId : $_"
    $Tasks = @()
}

## SEND DETAILS TO FLOW ##

Write-Output $Tasks | ConvertTo-Json -Depth 100