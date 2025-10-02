<#

Mangano IT - New User Flow - Send SMS to Required Contacts
Created by: Gabriel Nugent
Version: 1.51

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory)][int]$TicketId,
    [string]$DisplayName,
    [string]$MessageToSendToUser,
    [string]$MessageToSendToManager,
    [string]$SenderName = 'Mangano IT',
    [string]$ScheduledSendDate,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

[string]$Log = ''
$TaskNotes = 'Send SMS'
$TicketNoteDetails = ''
$TicketNoteFailure = ''

## GET CW MANAGE CREDENTIALS ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## FETCH TASKS FROM TICKET ##

$Tasks = .\CWM-FindTicketTasks.ps1 -TicketId $TicketId -TaskNotes $TaskNotes -ApiSecrets $ApiSecrets | ConvertFrom-Json

## SEND SMS ##

# If task is not complete, attempt to send SMS
foreach ($Task in $Tasks) {
    $Arguments = @{}
    if (!$Task.closedFlag -and $Task.notes -eq $TaskNotes) {
        $Resolution = $Task.resolution | ConvertFrom-Json
        $MobileNumber = $Resolution.MobileNumber
        $Arguments = @{
            MobileNumber = $MobileNumber
            CountryCode = $Resolution.CountryCode
            SenderName = $SenderName
        }

        # Set message based on contact type
        if ($Resolution.ContactType -eq 'User') {
            $Arguments += @{
                Message = $MessageToSendToUser
            }
        }

        if ($Resolution.ContactType -eq 'Manager') {
            $Arguments += @{
                Message = $MessageToSendToManager
            }
        }

        # Add scheduled send date if provided
        if ($ScheduledSendDate -ne '') {
            $Arguments += @{
                ScheduledSendDate = $ScheduledSendDate
            }
        }

        $Operation = .\SMS-SendMessage.ps1 @Arguments | ConvertFrom-Json
        $Log += "`n`n" + $Operation.Log

        # Update task upon completion
        if ($Operation.Result) {
            # Add details for ticket note
            $TicketNoteDetails += "`n- $MobileNumber"

            # Close task
            $TaskId = $Task.id
            .\CWM-UpdateSpecificTask.ps1 -TicketId $TicketId -TaskId $TaskId -ClosedStatus $true -ApiSecrets $ApiSecrets | Out-Null
        } else {
            $TicketNoteFailure += "`n- $MobileNumber"
            if ($MobileNumber -eq '') {
                $TicketNoteFailure += "[The number provided was blank. Please confirm that the relevant contact has a mobile number set, or that the new user's mobile number has been provided properly.]"
            }
        }
    }
}

## ADD NOTE TO TICKET ##

# Add note for results
if ($TicketNoteDetails -ne '') {
    $Text = $TaskNotes + " [automated task]`n`n"
    $Text += "Passwords have been sent to the following numbers:$TicketNoteDetails"
}
if ($TicketNoteFailure -ne '') {
    $Text = $TaskNotes + " [automated task]`n`n"
    $Text += "An attempt was made to text the following numbers:$TicketNoteFailure"
}
.\CWM-AddTicketNote.ps1 -TicketId $TicketId -Text $Text -ResolutionFlag $true -ApiSecrets $ApiSecrets | Out-Null

## SEND DETAILS TO FLOW ##

$Output = @{
    Log = $Log
}

Write-Output $Output | ConvertTo-Json