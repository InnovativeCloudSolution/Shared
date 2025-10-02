<#

Mangano IT - Exit User Flow - Create Tasks for Exit User Flow (Admin Account)
Created by: Gabriel Nugent
Version: 1.1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory=$true)][int]$TicketId,
    $ApiSecrets = $null,
    [string]$AadUserId,
    [string]$AadUserPrincipalName
)

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

# Initialize class for ticket task
class Task {
    [string]$Notes
    [string]$Resolution
    [boolean]$Exists
    [boolean]$Closed
}

## TASKS ##

$UserIdResolution = @{
    UserId = $AadUserId
    UserPrincipalName = $AadUserPrincipalName
} | ConvertTo-Json

# Define standard tasks
$Tasks = @(
    [Task]@{Notes="Revoke Sign In Sessions in AAD ($AadUserPrincipalName)";Resolution=$UserIdResolution;Exists=$false;Closed=$false},
    [Task]@{Notes="Disable MFA in AAD ($AadUserPrincipalName) (manual)"
    Resolution="This step must be completed manually.`n`nUser principal name: $AadUserPrincipalName"
    Exists=$false;Closed=$false},
    [Task]@{Notes="Remove from Cloud Groups ($AadUserPrincipalName)";Resolution=$UserIdResolution;Exists=$false;Closed=$false},
    [Task]@{Notes="Remove Azure AD Role Assignments ($AadUserPrincipalName) (manual)";Resolution="This step cannot be completed by automation due to permission issues. Please complete this manually.";Exists=$false;Closed=$false}
    [Task]@{Notes="Remove Azure Enterprise Applications ($AadUserPrincipalName)";Resolution=$UserIdResolution;Exists=$false;Closed=$false},
    [Task]@{Notes="Unassign Licenses ($AadUserPrincipalName)";Resolution=$UserIdResolution;Exists=$false;Closed=$false}
)

## FILTER THROUGH TASKS ##

try {
    $Log += "Getting tasks for #$TicketId...`n"
    $TicketTasks = .\CWM-FindTicketTasks.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets | ConvertFrom-Json
    $Log += "SUCCESS: Fetched tasks for #$TicketId.`n`n"
    Write-Warning "SUCCESS: Fetched tasks for #$TicketId."
} catch {
    $Log += "ERROR: Unable to fetch tasks for #$TicketId.`nERROR DETAILS: " + $_
    Write-Error "Unable to fetch tasks for #$TicketId : $_"
    $TicketTasks = $null
}

# Ignore tasks that already exist
:Outer foreach ($TicketTask in $TicketTasks) {
    foreach ($Task in $Tasks) {
        if ($Task.Notes -like $TicketTask.notes -and $Task.Resolution -like $TicketTask.resolution) {
            $TaskNotes = $Task.Notes
            $TaskClosed = ($Task.Closed).ToString()
            $Task.Exists = $true
            $Task.Closed = $TicketTask.closedFlag
            $Log += "`nINFO: Task '$TaskNotes' already exists. Closed status: $TaskClosed"
            Write-Warning "INFO: Task '$TaskNotes' already exists. Closed status: $TaskClosed"
            continue Outer
        }
    }
}

## ADD TICKET TASKS ##

foreach ($Task in $Tasks) {
    if (!$Task.Exists) {
        $TaskNotes = $Task.Notes
        $TaskResolution = $Task.Resolution
        $TaskParameters = @{
            TicketId = $TicketId
            Note = $TaskNotes
            Resolution = $TaskResolution
            ApiSecrets = $ApiSecrets
        }
        $NewTask = .\CWM-AddTicketTask.ps1 @TaskParameters | ConvertFrom-Json
        $Log += $NewTask.Log + "`n`n"
    }
}

## SEND DETAILS BACK TO FLOW ##

# Get all currently existing tasks on this ticket again
try {
    $Log += "`nFetching ticket tasks again for #$TicketId..."
    $TicketTasks = .\CWM-FindTicketTasks.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets | ConvertFrom-Json
    $Log += "`nSUCCESS: Ticket tasks fetched for #$TicketId."
} catch {
    $Log += "`nERROR: Ticket tasks unable to be fetched for #$TicketId.`nERROR DETAILS: " + $_
    Write-Error "Ticket tasks unable to be fetched for #$TicketId : $_"
}

$Output = @{
    Log = $Log
    TicketTasks = $TicketTasks
}

# Send back to Power Automate
Write-Output $Output | ConvertTo-Json