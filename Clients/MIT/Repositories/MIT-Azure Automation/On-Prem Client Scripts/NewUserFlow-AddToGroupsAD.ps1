<#

Mangano IT - New User Flow - Active Directory - Add New User to Groups (New User Flow)
Created by: Gabriel Nugent
Version: 1.1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [int]$TicketId,
    [string]$SamAccountName,
    [string]$AzureADServer,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$TicketNoteDetails = ''
$TaskNotes = 'Add to On-Premises Security Group'
[string]$Log = ''

## GET CW MANAGE CREDENTIALS ##

if ($null -eq $ApiSecrets) {
    $Log += "ERROR: ConnectWise Manage API secrets not provided. The script will not work."
    Write-Error "ConnectWise Manage API secrets not provided. The script will not work."
}

## FETCH TASKS FROM TICKET ##

$Tasks = .\CWM-FindTicketTasks.ps1 -TicketId $TicketId -TaskNotes $TaskNotes -ApiSecrets $ApiSecrets | ConvertFrom-Json

## ADD TO MICROSOFT 365 GROUPS ##

# If task is not complete, attempt to add to 365 group
foreach ($Task in $Tasks) {
    if (!$Task.closedFlag -and $Task.notes -eq $TaskNotes) {
        $Resolution = $Task.resolution | ConvertFrom-Json
        $Name = $Resolution.Name
        $Arguments = @{
            SamAccountName = $SamAccountName
            SecurityGroupName = $Name
        }
        $Operation = .\AD-AddToGroup.ps1 @Arguments | ConvertFrom-Json
        $Log += "`n`n" + $Operation.Log

        # Update task upon completion
        if ($Operation.Result) {
            # Add details for ticket note
            $TicketNoteDetails += "`n- $Name"

            # Close task
            $TaskId = $Task.id
            .\CWM-UpdateSpecificTask.ps1 -TicketId $TicketId -TaskId $TaskId -ClosedStatus $true -ApiSecrets $ApiSecrets | Out-Null
        }
    }
}

# Add note if the user was added to any distribution groups, and run AD sync
if ($TicketNoteDetails -ne '') {
    $Text = $TaskNotes + ": The new user has been assigned access to the following on-premises security groups:$TicketNoteDetails"
    .\CWM-AddTicketNote.ps1 -TicketId $TicketId -Text $Text -ResolutionFlag $true -ApiSecrets $ApiSecrets | Out-Null

    $SyncResult = .\AD-RunADSync.ps1 -AzureADServer $AzureADServer
    $Log += "`n`nINFO: ADSync $SyncResult"
    Write-Warning "AD sync result: $SyncResult"
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Log = $Log
}

Write-Output $Output | ConvertTo-Json