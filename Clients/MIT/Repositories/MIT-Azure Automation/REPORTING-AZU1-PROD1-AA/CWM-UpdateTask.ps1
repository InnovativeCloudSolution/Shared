<#

Mangano IT - ConnectWise Manage - Update Task
Created by: Gabriel Nugent
Version: 1.3

This runbook is designed to be used in conjunction with other Azure Automation scripts.

Not to be used for tasks that have multiple with the same title!

#>

param (
    [Parameter(Mandatory=$true)][int]$TicketId,
    [int]$TaskId,
	[Parameter(Mandatory=$true)][string]$Note,
    [string]$Resolution,
    [Parameter(Mandatory=$true)][bool]$ClosedStatus, # True = closed
    [bool]$TaskWithDuplicates = $false, # For tasks where multiple have the same note
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

$GetTasksArguments = @{
    Uri = "$CWMApiUrl/service/tickets/$TicketId/tasks"
    Method = 'GET'
    ContentType = $ContentType
    Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
    UseBasicParsing = $true
}

# If ID provided, specify
if ($TaskId -ne 0) { $GetTasksArguments.Uri += "/$TaskId" }

## FETCH LIST OF TASKS ##

try {
    $Log += "Fetching tasks for #$TicketId...`n"
    $Tasks = Invoke-WebRequest @GetTasksArguments | ConvertFrom-Json
    $Log += "SUCCESS: Tasks fetched for #$TicketId.`n`n"
    Write-Warning "SUCCESS: Tasks fetched for #$TicketId."
} catch {
    $Tasks = $null
    $Log += "ERROR: Tasks not fetched for #$TicketId.`nERROR DETAILS: " + $_
    Write-Error "Tasks not fetched for #$TicketId : $_"
}

# If task title matches requested task, or all task details match given details, update task
foreach ($Task in $Tasks) {
    if (($Task.notes -eq $Note) -or ($Task.notes -eq $Note -and $Task.resolution -eq $Resolution -and $TaskWithDuplicates)`
    -or ($Task.id -eq $TaskId)) {
        $TaskId = $Task.id
        $Log += "INFO: Task located.`nTask ID: $TaskId`n`n"
        Write-Warning "Task ID: $TaskId"

        $UpdateTaskBody = @(
            @{
                op = "replace"
                path = "/closedFlag"
                value = $ClosedStatus
            }
        )

        # If resolution provided and task isn't a duplicate, update resolution
        if ($Resolution -ne '' -and !$TaskWithDuplicates) {
            $UpdateTaskBody += @{
                op = "replace"
                path = "/resolution"
                value = $Resolution
            }
        }

        $UpdateTaskArguments = @{
            Uri = "$CWMApiUrl/service/tickets/$TicketId/tasks/$TaskId"
            Method = 'PATCH'
            ContentType = $ContentType
            Body = ConvertTo-Json -InputObject $UpdateTaskBody
            Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
            UseBasicParsing = $true
        }

        try {
            $Log += "Attempting to update task $TaskId...`n"
            $UpdatedTask = Invoke-WebRequest @UpdateTaskArguments | ConvertFrom-Json
            $Log += "SUCCESS: Updated task $TaskId."
            Write-Warning "SUCCESS: Updated task $TaskId."
        } catch {
            $UpdatedTask = $null
            $Log += "ERROR: Unable to update task $TaskId.`nERROR DETAILS: " + $_
            Write-Error "Unable to update task $TaskId : $_"
        }
    }
}

$Output = @{
    Task = $UpdatedTask
    Log = $Log
}

Write-Output $Output | ConvertTo-Json