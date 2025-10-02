<#

Mangano IT - ConnectWise Manage - Update Task and Add Note for Flow
Created by: Gabriel Nugent
Version: 1.3

This runbook is designed to be used only in conjunction with Azure Automation runbooks.
It is to be called when a script (e.g. create new user) has finished.

#>

param (
    [Parameter(Mandatory=$true)][bool]$Result,
    [Parameter(Mandatory=$true)][int]$TicketId,
    [int]$TaskId,
    [Parameter(Mandatory=$true)][string]$TaskNotes,
    [string]$TaskResolution,
    [Parameter(Mandatory=$true)][string]$TicketNote_Success,
    [Parameter(Mandatory=$true)][string]$TicketNote_Failure,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

[string]$Log = ''

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## ADD NOTE ##

# Define ticket note text
$Text = "$TaskNotes [automated task]`n`n"
if ($Result) { 
    $Text += $TicketNote_Success
    $NoteArguments = @{
        TicketId = $TicketId
        Text = $Text
        ResolutionFlag = $true
        ApiSecrets = $ApiSecrets
    }
}
else {
    $Text += $TicketNote_Failure
    $NoteArguments = @{
        TicketId = $TicketId
        Text = $Text
        IssueFlag = $true
        InternalFlag = $true
        ApiSecrets = $ApiSecrets
    }
}

# Add note to ticket
$Note = .\CWM-AddTicketNote.ps1 @NoteArguments | ConvertFrom-Json
$Log += $Note.Log + "`n`n"

## UPDATE TASK ##

# Close task if successful
if ($TaskNotes -ne '') {
    $TaskArguments = @{
        TicketId = $TicketId
        Note = $TaskNotes
        ClosedStatus = $Result
        ApiSecrets = $ApiSecrets
    }

    # Add task ID if provided
    if ($TaskId -ne '') {
        $TaskArguments += @{
            TaskId = $TaskId
        }
    }

    # Add resolution if provided
    if ($TaskResolution -ne '') {
        $TaskArguments += @{
            Resolution = $TaskResolution
        }
    }

    $Task = .\CWM-UpdateTask.ps1 @TaskArguments | ConvertFrom-Json
    $Log += "`n`n" + $Task.Log
}

## WRITE LOG ##

$Output = @{
    Log = $Log
}

Write-Output $Output