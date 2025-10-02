<#

Mangano IT - New User Flow - Set DeskDirector Group for New User
Created by: Gabriel Nugent
Version: 1.0.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory=$true)][int]$TicketId,
    [Parameter(Mandatory=$true)][int]$ContactId,
    [Parameter(Mandatory=$true)][int]$CompanyId,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$TicketNoteDetails = ''
$TaskNotes = 'Set DeskDirector Group'
[string]$Log = ''

## GET CW MANAGE CREDENTIALS ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## FETCH TASKS FROM TICKET ##

$Tasks = .\CWM-FindTicketTasks.ps1 -TicketId $TicketId -TaskNotes $TaskNotes -ApiSecrets $ApiSecrets | ConvertFrom-Json

## CHECK LICENSE ASSIGNED STATUS ##

foreach ($Task in $Tasks) {
    if ($Task.notes -eq $TaskNotes) {
        $TaskId = $Task.id
        $TeamId = $Task.resolution
        $AddToGroupArguments = @{
            UserId = $ContactId
            CompanyId = $CompanyId
            TeamId = $TeamId
        }
        $Operation = .\DeskDirector-SetUserGroup.ps1 @AddToGroupArguments | ConvertFrom-Json
        $Log += "`n`n" + $Operation.Log

        # Close task if license assigned, re-open if not
        if ($Operation.Result) {
            .\CWM-UpdateSpecificTask.ps1 -TicketId $TicketId -TaskId $TaskId -ClosedStatus $true -ApiSecrets $ApiSecrets | Out-Null
        }
        else {
            .\CWM-UpdateSpecificTask.ps1 -TicketId $TicketId -TaskId $TaskId -ClosedStatus $false -ApiSecrets $ApiSecrets | Out-Null
            $TicketNoteDetails += "`n- $TeamId"
        }
    }
}

# Set note and status details depending on result
if ($TicketNoteDetails -ne '') {
    $Text = "$TaskNotes [automated task]`n`n"
    $Text += "The new user was not added to the following DeskDirector groups:$TicketNoteDetails"
    $GroupsAdded = $false
} else {
    $Text = "$TaskNotes [automated task]`n`n"
    $Text += "The new user's contact has been added to all required DeskDirector groups."
    $GroupsAdded = $true
}

$NoteArguments = @{
    TicketId = $TicketId
    Text = $Text
    ResolutionFlag = $true
    ApiSecrets = $ApiSecrets
}

.\CWM-AddTicketNote.ps1 @NoteArguments | Out-Null

## SEND DETAILS TO FLOW ##

$Output = @{
    GroupsAdded = $GroupsAdded
    Log = $Log
}

Write-Output $Output | ConvertTo-Json