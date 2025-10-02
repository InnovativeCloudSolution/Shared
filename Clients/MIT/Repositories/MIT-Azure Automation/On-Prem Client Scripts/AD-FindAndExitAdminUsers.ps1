<#

Mangano IT - Active Directory - Check For and Disable Admin Account/s
Created by: Gabriel Nugent
Version: 1.8

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [int]$TicketId,
    [string]$StandardSamAccountName,
    [string]$AdminPrefix = 'Admin.',
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
$TaskNotes = "Check For and Disable Admin Account/s"
$AdminAccountsDisabled = ''
$AdminAccountsRemaining = ''
$AdminDetails = @()

# Output the status variable content to a file
$Date = Get-Date -Format "dd-MM-yyyy HHmm"
$FilePath = "C:\Scripts\Logs\AD-FindAndExitAdminUsers"
$FileName = "$StandardSamAccountName-$Date.txt"

## GET USER ##

try {
    $Log += "Attempting to fetch users that are admin accounts for $StandardSamAccountName...`n"
    if ($StandardSamAccountName.Length -gt 10) { $StandardSamAccountName = $StandardSamAccountName.Substring(0, 10) }
    $Users = Get-ADUser -Filter "SamAccountName -like '$AdminPrefix*' -and SamAccountName -like '*$StandardSamAccountName*'"
    $Log += "SUCCESS: User/s have been fetched.`n`n"
    Write-Warning "SUCCESS: User/s have been fetched."
} catch {
    $Log += "ERROR: Unable to fetch user.`nERROR DETAILS: " + $_
    Write-Error "Unable to fetch user : $_"
}

## DISABLE USER ACCOUNT/S ##

if ($null -ne $Users) {
    foreach ($User in $Users) {
        # Set up SAM account name
        $SamAccountName = $User.SamAccountName

        # Add UPN and SAM to list
        $AdminDetails += @{
            UserPrincipalName = $User.UserPrincipalName
            SamAccountName = $SamAccountName
        }

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
                $AdminAccountsDisabled += "`n- $SamAccountName"
            } catch {
                $Log += "ERROR: Unable to disable user account $SamAccountName.`n`nERROR DETAILS: " + $_
                Write-Error "Unable to disable user account $SamAccountName : $_"
                $AdminAccountsRemaining += "`n- $SamAccountName"
            }
        }
        
        ## MOVE TO DISABLED USERS OU ##
        
        if ($DisabledUsersOU -ne '') {
            try {
                $Log += "Attempting to move $SamAccountName to $DisabledUsersOU...`n"
                $User | Move-ADObject -TargetPath $DisabledUsersOU
                $Log += "SUCCESS: Moved $SamAccountName to $DisabledUsersOU.`n`n"
                Write-Warning "SUCCESS: Moved $SamAccountName to $DisabledUsersOU."
            } catch {
                $Log += "ERROR: Unable to move $SamAccountName to $DisabledUsersOU.`n`nERROR DETAILS: " + $_
                Write-Error "Unable to move $SamAccountName to $DisabledUsersOU : $_"
            }
        }
    }
} else {
    $Log += "INFO: No admin accounts located.`n`n"
    Write-Warning "INFO: No admin accounts located."
    $SamAccountName = ''
    $AdminDetails += @{ 
        UserPrincipalName = ''
        SamAccountName = ''
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
    # Set up standard note
    if ($AdminAccountsDisabled -eq '' -and $AdminAccountsRemaining -eq '') {
        $TicketNote = "Based on the automation's findings, this user has no admin accounts."
        $Result = $true
    } else {
        if ($AdminAccountsDisabled -ne '') {
            $TicketNote += "The following accounts have been disabled:$AdminAccountsDisabled"
            $Result = $true
        }
        if ($AdminAccountsRemaining -ne '') {
            if ($TicketNote -ne '') { $TicketNote += "`n`n" }
            $TicketNote += "The following accounts have NOT been disabled:$AdminAccountsDisabled"
            $Result = $false
        }
    }

    # Define arguments for ticket note and task
    $TaskNoteArguments = @{
        Result = $Result
        TicketId = $TicketId
        TaskNotes = $TaskNotes
        TicketNote_Success = $TicketNote
        TicketNote_Failure = $TicketNote
        ApiSecrets = $ApiSecrets
    }

    # Add details of AD sync if performed
    if ($ForceADSync) {
        $NoteText += "`n`nAD sync: $SyncResult`n`n"
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

$Output = @{
    Result = $Result
    AdminDetails = $AdminDetails
    SyncResult = $SyncResult
    Log = $Log
    LogFile = "$FilePath\$FileName"
}

Write-Output $Output | ConvertTo-Json -Depth 100

# Makes folder for logs and outputs logs
.\CreateLogFile.ps1 -Log $Log -FilePath $FilePath -FileName $FileName | Out-Null