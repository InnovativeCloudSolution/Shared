<#

Mangano IT - ConnectWise Manage - Update Specific Task
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with other Azure Automation scripts.

Takes a specific task ID instead of trying to find the task.

#>

param (
    [Parameter(Mandatory=$true)][int]$TicketId,
    [Parameter(Mandatory=$true)][int]$TaskId,
	[string]$Note,
    [string]$Resolution,
    [Parameter(Mandatory=$true)][bool]$ClosedStatus, # True = closed
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

## UPDATE TASK ##

$ApiBody = @(
    @{
        op = "replace"
        path = "/closedFlag"
        value = $ClosedStatus
    }
)

# If note provided, update note
if ($Note -ne '') {
    $ApiBody += @{
        op = "replace"
        path = "/notes"
        value = $Note
    }
}

# If resolution provided, update resolution
if ($Resolution -ne '') {
    $ApiBody += @{
        op = "replace"
        path = "/resolution"
        value = $Resolution
    }
}

$ApiArguments = @{
    Uri = "$CWMApiUrl/service/tickets/$TicketId/tasks/$TaskId"
    Method = 'PATCH'
    ContentType = $ContentType
    Body = ConvertTo-Json -InputObject $ApiBody
    Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
    UseBasicParsing = $true
}

# Update task
try {
    $Log += "Attempting to update task $TaskId...`n"
    $UpdatedTask = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
    $Log += "SUCCESS: Updated task $TaskId."
    Write-Warning "SUCCESS: Updated task $TaskId."
} catch {
    $UpdatedTask = $null
    $Log += "ERROR: Unable to update task $TaskId.`nERROR DETAILS: " + $_
    Write-Error "Unable to update task $TaskId : $_"
}

$Output = @{
    Task = $UpdatedTask
    Log = $Log
}

Write-Output $Output | ConvertTo-Json