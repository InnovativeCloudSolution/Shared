<#

Mangano IT - ConnectWise Manage - Add Task to Ticket
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow or an Azure Automation script.

#>

param (
    [Parameter(Mandatory=$true)][int]$TicketId,
    [Parameter(Mandatory=$true)][string]$Notes,
    [string]$Resolution,
    [bool]$ClosedFlag = $false,
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

# Build request body
$ApiBody = @{
    notes = $Notes
    closedFlag = $ClosedFlag
}

# If resolution provided, add to body
if ($Resolution -ne '') { $ApiBody += @{ resolution = $Resolution } }

# Build API arguments
$ApiArguments = @{
    Uri = "$CWMApiUrl/service/tickets/$TicketId/tasks"
    Method = 'POST'
    ContentType = $ContentType
    Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
    Body = $ApiBody | ConvertTo-Json -Depth 100
    UseBasicParsing = $true
}

## ADD TASK TO TICKET ##
try {
    $Log += "Adding task '$Notes' to ticket #$TicketId...`n"
    Invoke-WebRequest @ApiArguments | Out-Null
    $Result = $true
    $Log += "SUCCESS: Task added to ticket #$TicketId."
    Write-Warning "SUCCESS: Task added to ticket #$TicketId."
} catch {
    $Result = $false
    $Log += "ERROR: Task not added to ticket #$TicketId.`nERROR DETAILS: " + $_
    Write-Error "Task not added to ticket #$TicketId : $_"
}

## SEND OUTPUT TO FLOW ##

$Output = @{
    Result = $Result
    Log = $Log
}

Write-Output $Output | ConvertTo-Json