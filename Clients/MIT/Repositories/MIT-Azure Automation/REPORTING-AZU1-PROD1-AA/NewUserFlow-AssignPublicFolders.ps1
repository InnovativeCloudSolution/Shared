<#

Mangano IT - New User Flow - Assign New User to Public Folders
Created by: Gabriel Nugent
Version: 1.4.1

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
$TaskNotes = 'Grant Access to Public Folder'
[string]$Log = ''

## GET CW MANAGE CREDENTIALS ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## FETCH TASKS FROM TICKET ##

$Tasks = .\CWM-FindTicketTasks.ps1 -TicketId $TicketId -TaskNotes $TaskNotes -ApiSecrets $ApiSecrets | ConvertFrom-Json

## FETCH TENANT SLUG ##

$TenantSlug = .\CWM-FindCompanySlug.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets

## ADD TO PUBLIC FOLDERS ##

# If task is not complete, attempt to assign access to public folder
foreach ($Task in $Tasks) {
    if (!$Task.closedFlag -and $Task.notes -eq $TaskNotes) {
        $Resolution = $Task.resolution | ConvertFrom-Json
        $Name = $Resolution.Name
        $AccessRights = $Resolution.AccessRights
        $Arguments = @{
            TenantUrl = $TenantUrl
            TenantSlug = $TenantSlug
            EmailAddress = $EmailAddress
            PublicFolder = $Name
            AccessRights = $AccessRights
        }
        $Operation = .\EXO-AssignPublicFolder.ps1 @Arguments | ConvertFrom-Json
        $Log += "`n`n" + $Operation.Log

        # Update task upon completion
        if ($Operation.Result) {
            # Add details for ticket note
            $TicketNoteDetails += "`n- $Name"

            # Close task
            $TaskId = $Task.id
            .\CWM-UpdateSpecificTask.ps1 -TicketId $TicketId -TaskId $TaskId -ClosedStatus $true -ApiSecrets $ApiSecrets | Out-Null
        } else { $FailedNoteDetails += "`n- $Name" }
    }
}

# Add note if the user was added to any public folders
$Text = $TaskNotes + " [automated task]`n`n"

if ($TicketNoteDetails -ne '' -or $FailedNoteDetails -ne '') {
    if ($TicketNoteDetails -ne '') {
        $Text += "`n`nThe new user has been assigned access to the following public folders:$TicketNoteDetails"
    }
    if ($FailedNoteDetails -ne '') {
        $Text += "`n`nThe new user has NOT been assigned access to the following public folders:$FailedNoteDetails"
    }
    .\CWM-AddTicketNote.ps1 -TicketId $TicketId -Text $Text -ResolutionFlag $true -ApiSecrets $ApiSecrets | Out-Null
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Log = $Log
}

Write-Output $Output | ConvertTo-Json