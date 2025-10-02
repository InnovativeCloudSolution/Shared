<#

Mangano IT - New User Flow - Add Bullphish Addresses to Safe Senders List
Created by: Gabriel Nugent
Version: 1.0.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [int]$TicketId,
    [string]$EmailAddress,
    [string]$TenantUrl,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$TicketNoteDetails = ''
$FailedNoteDetails = ''
$TaskNotes = 'Add Bullphish Addresses to Safe Senders List'
[string]$Log = ''
$BullphishSenders = @("bp-service-support.com","bp-securityawareness.com","online-account.info", "myonlinesecuritysupport.com","service-noreply.info", "banking-alerts.info","bullphish.com","verifyaccount.help","suspected-fraud.info")

## GET CW MANAGE CREDENTIALS ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## FETCH TASKS FROM TICKET ##

$Tasks = .\CWM-FindTicketTasks.ps1 -TicketId $TicketId -TaskNotes $TaskNotes -ApiSecrets $ApiSecrets | ConvertFrom-Json

## FETCH TENANT SLUG ##

$TenantSlug = .\CWM-FindCompanySlug.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets

## ADD TO DISTRIBUTION GROUPS ##

# If task is not complete, attempt to assign access to distribution group
foreach ($Task in $Tasks) {
    if (!$Task.closedFlag -and $Task.notes -eq $TaskNotes) {
        $Arguments = @{
            TenantUrl = $TenantUrl
            TenantSlug = $TenantSlug
            EmailAddress = $EmailAddress
            Senders = $BullphishSenders
        }
        $Operation = .\EXO-AddSafeSenders.ps1 @Arguments | ConvertFrom-Json
        $Log += "`n`n" + $Operation.Log

        # Update task upon completion
        if ($Operation.Result) {
            # Add details for ticket note
            $TicketNoteDetails += $BullphishSenders

            # Close task
            $TaskId = $Task.id
            .\CWM-UpdateSpecificTask.ps1 -TicketId $TicketId -TaskId $TaskId -ClosedStatus $true -ApiSecrets $ApiSecrets | Out-Null
        } else { $FailedNoteDetails += $BullphishSenders }
    }
}

# Add note if operation completed
$Text = $TaskNotes + " [automated task]`n`n"

if ($TicketNoteDetails -ne '' -or $FailedNoteDetails -ne '') {
    if ($TicketNoteDetails -ne '') {
        $Text += "`n`nThe new user has had their safe senders list updated:`n$TicketNoteDetails"
    }
    if ($FailedNoteDetails -ne '') {
        $Text += "`n`nThe new user has NOT had their safe senders list updated:`n$FailedNoteDetails"
    }
    .\CWM-AddTicketNote.ps1 -TicketId $TicketId -Text $Text -ResolutionFlag $true -ApiSecrets $ApiSecrets | Out-Null
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Log = $Log
}

Write-Output $Output | ConvertTo-Json