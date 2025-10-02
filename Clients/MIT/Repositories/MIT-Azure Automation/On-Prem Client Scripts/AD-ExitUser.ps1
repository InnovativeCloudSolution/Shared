<#

Mangano IT - Active Directory - Exit AD User
Created by: Gabriel Nugent
Version: 1.3

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [int]$TicketId,
    [Parameter(Mandatory)][string]$UserPrincipalName,
    [string]$DisabledUsersOU,
    [bool]$RemoveManager = $true,
    [bool]$HideFromGal = $true,
    [bool]$ForceADSync,
    [string]$AzureADServer,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

[string]$Log = ''
$Result = $false
$TaskNotes = "Disable User Account"
$MoveOuResult = 'N/A'

# Output the status variable content to a file
$Date = Get-Date -Format "dd-MM-yyyy HHmm"
$FilePath = "C:\Scripts\Logs\AD-ExitUser"
$FileName = "$GivenName $Surname-$Date.txt"

## GET USER ##

try {
    $Log += "Attempting to fetch user for $UserPrincipalName...`n"
    $User = Get-ADUser -Filter "UserPrincipalName -eq '$UserPrincipalName'"
    $Log += "SUCCESS: User has been fetched.`n`n"
    Write-Warning "SUCCESS: User has been fetched."
    $SamAccountName = $User.SamAccountName
} catch {
    $Log += "ERROR: Unable to fetch user.`nERROR DETAILS: " + $_
    Write-Error "Unable to fetch user : $_"
}

## DISABLE USER ACCOUNT ##

$DisableUserParams = @{
    Enabled = $false
    Manager = $null
    Confirm = $false
}

if ($HideFromGal) {
    $Log += "INFO: $SamAccountName will be hidden from the Global Address List.`n`n"
    $DisableUserParams += @{ Add = @{ msExchHideFromAddressLists = $true } }
}

if ($null -ne $User) {
    try {
        $Log += "Attempting to disable user account $SamAccountName...`n"
        $User | Set-ADUser @DisableUserParams | Out-Null
        $Log += "SUCCESS: User account $SamAccountName disabled.`n`n"
        Write-Warning "SUCCESS: User account $SamAccountName disabled."
        $Result = $true
    } catch {
        $Log += "ERROR: Unable to disable user account $SamAccountName.`n`nERROR DETAILS: " + $_
        Write-Error "Unable to disable user account $SamAccountName : $_"
    }
}

## MOVE TO DISABLED USERS OU ##

if ($DisabledUsersOU -ne '') {
    try {
        $Log += "Attempting to move $SamAccountName to $DisabledUsersOU...`n"
        $User | Move-ADObject -TargetPath $DisabledUsersOU
        $Log += "SUCCESS: Moved $SamAccountName to $DisabledUsersOU.`n`n"
        Write-Warning "SUCCESS: Moved $SamAccountName to $DisabledUsersOU."
        $MoveOuResult = 'True'
    } catch {
        $Log += "ERROR: Unable to move $SamAccountName to $DisabledUsersOU.`n`nERROR DETAILS: " + $_
        Write-Error "Unable to move $SamAccountName to $DisabledUsersOU : $_"
        $MoveOuResult = 'False'
    }
}

## RUN AZURE AD SYNC IF REQUIRED ##

if ($ForceADSync -and $Result) {
    $SyncResult = .\AD-RunADSync.ps1 -AzureADServer $AzureADServer
    $Log += "`n`nINFO: AD Sync $SyncResult"
    Write-Warning "AD sync result: $SyncResult"
} else { $SyncResult = $false }

## UPDATE TICKET IF PROVIDED ##

if ($TicketId -ne 0) {
    # Define arguments for ticket note and task
    $TaskNoteArguments = @{
        Result = $Result
        TicketId = $TicketId
        TaskNotes = $TaskNotes
        TicketNote_Success = "User $UserPrincipalName has been disabled in the company's Active Directory.`n`n"
        TicketNote_Failure = "ERROR: User $UserPrincipalName has NOT been disabled in the company's Active Directory.`n`n"
        ApiSecrets = $ApiSecrets
    }

    # Add details for new OU if moved
    if ($DisabledUsersOU -ne '') {
        $NoteText += "Moved into '$DisabledUsersOu': $MoveOuResult`n`n"
    }

    # Add details of AD sync if performed
    if ($ForceADSync) {
        $NoteText += "AD sync: $SyncResult`n`n"
    }

    # Add log location and update standard notes
    $NoteText += "Log file: $FilePath\$FileName"
    $TaskNoteArguments.TicketNote_Success += $NoteText
    $TaskNoteArguments.TicketNote_Failure += $NoteText
    
    # Add note and update task if successful
    $TaskAndNote = .\CWM-UpdateTaskAddNoteForFlow.ps1 @TaskNoteArguments
    $Log += $TaskAndNote.Log
}

## SEND DETAILS BACK TO FLOW ##

if ($null -eq $SamAccountName) { $SamAccountName = '' }

$Output = @{
    Result = $Result
    SamAccountName = $SamAccountName
    SyncResult = $SyncResult
    Log = $Log
    LogFile = "$FilePath\$FileName"
}

Write-Output $Output | ConvertTo-Json

# Makes folder for logs and outputs logs
.\CreateLogFile.ps1 -Log $Log -FilePath $FilePath -FileName $FileName | Out-Null