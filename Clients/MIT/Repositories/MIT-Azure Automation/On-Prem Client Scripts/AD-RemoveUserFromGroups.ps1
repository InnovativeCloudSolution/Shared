<#

Mangano IT - Active Directory - Remove AD User from Security Groups
Created by: Gabriel Nugent
Version: 1.61

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [int]$TicketId,
    [Parameter(Mandatory)][string]$SamAccountName,
    [string]$PrimaryGroupName = 'Domain Users',
    [bool]$ForceADSync,
    [string]$AzureADServer,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

[string]$Log = ''
$Result = $true
$TaskNotes = "Remove from On-Premises Security Groups"
$FilePath = "C:\Scripts\Logs\AD-RemoveUserFromGroups"
$FileName = "$GivenName $Surname-$Date.txt"
$RemovedGroups = ''
$FailedGroups = ''

## GET USER ##

$User = Get-ADUser -Filter "SamAccountName -eq '$SamAccountName'"

## GET ALL GROUPS WITH USER ##

try {
    $Log += "Getting all groups that $SamAccountName is a member of...`n"
    $Groups = Get-ADPrincipalGroupMembership -Identity $User
    $Log += "SUCCESS: Fetched all groups that $SamAccountName is a member of."
    Write-Warning "SUCCESS: Fetched all groups that $SamAccountName is a member of."
} catch {
    $Log += "ERROR: Unable to get all groups that $SamAccountName is a member of.`nERROR DETAILS: " + $_
    Write-Error "Unable to get all groups that $SamAccountName is a member of : $_"
}

## REMOVE USER FROM GROUPS ##

foreach ($Group in $Groups) {
    $GroupName = $Group.name
    if ($GroupName -ne $PrimaryGroupName) {
        $GroupGuid = $Group.objectGUID
        try {
            $Log += "`n`nAttempting to remove $SamAccountName from $GroupName...`n"
            Remove-ADGroupMember -Identity $GroupGuid -Members $User -Confirm:$false
            $Log += "SUCCESS: Removed $SamAccountName from $GroupName."
            Write-Warning "SUCCESS: Removed $SamAccountName from $GroupName."
            $RemovedGroups += "`n- $GroupName"
        } catch {
            $Log += "ERROR: Unable to remove $SamAccountName from $GroupName.`nERROR DETAILS: " + $_
            Write-Error "Unable to remove $SamAccountName from $GroupName : $_"
            $FailedGroups += "`n- $GroupName"
            $Result = $false
        }
    }
}

## RUN AZURE AD SYNC IF REQUIRED ##

if ($ForceADSync) {
    $SyncResult = .\AD-RunADSync.ps1 -AzureADServer $AzureADServer
    $Log += "`n`nINFO: AD Sync $SyncResult"
    Write-Warning "AD sync result: $SyncResult"
} else { $SyncResult = $false }

## UPDATE TICKET IF PROVIDED ##

if ($TicketId -ne 0) {
    # Define standard note text
    if ($RemovedGroups -eq '' -and $FailedGroups -eq '') {
        $NoteText += "There were no groups that $SamAccountName needed to be removed from.`n`n"
    } else {
        if ($RemovedGroups -ne '') {
            $NoteText += "$SamAccountName has been removed from the following groups:$RemovedGroups`n`n"
        }
        if ($FailedGroups -ne '') {
            $NoteText += "$SamAccountName has not been removed from the following groups:$FailedGroups`n`n"
        }
    }

    # Add details of AD sync if performed
    if ($ForceADSync) { $NoteText += "AD sync: $SyncResult`n`n" }

    # Add log location and update standard notes
    $NoteText += "Log file: $FilePath\$FileName"

    # Define arguments for ticket note and task
    $TaskNoteArguments = @{
        Result = $Result
        TicketId = $TicketId
        TaskNotes = $TaskNotes
        TicketNote_Success = $NoteText
        TicketNote_Failure = $NoteText
        ApiSecrets = $ApiSecrets
    }

    # Add missed groups to task if any
    if ($FailedGroups -ne '') {
        $TaskNoteArguments += @{
            TaskResolution = "Missed groups:$FailedGroups"
        }
    }
    
    # Add note and update task if successful
    $TaskAndNote = .\CWM-UpdateTaskAddNoteForFlow.ps1 @TaskNoteArguments
    $Log += $TaskAndNote.Log
}

## SEND DETAILS BACK TO FLOW ##

$Output = @{
    Result = $Result
    SyncResult = $SyncResult
    Log = $Log
    LogFile = "$FilePath\$FileName"
}

Write-Output $Output | ConvertTo-Json

# Makes folder for logs and outputs logs
.\CreateLogFile.ps1 -Log $Log -FilePath $FilePath -FileName $FileName | Out-Null