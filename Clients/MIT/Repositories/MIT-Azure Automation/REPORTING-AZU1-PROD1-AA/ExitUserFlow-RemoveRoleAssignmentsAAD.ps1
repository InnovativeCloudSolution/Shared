<#

Mangano IT - Exit User Flow - Remove App Role Assignments for User (AAD)
Created by: Gabriel Nugent
Version: 1.6

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [int]$TicketId,
    [string]$BearerToken,
    [string]$TenantUrl,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$TaskNotes = 'Remove Azure Enterprise Applications'
[string]$Log = ''

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    $Log += "Bearer token not supplied. Getting bearer token for $TenantUrl...`n`n"
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

## GET CW MANAGE CREDENTIALS ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## FETCH TASKS FROM TICKET ##

$Tasks = .\CWM-FindTicketTasks.ps1 -TicketId $TicketId -TaskNotes $TaskNotes -ApiSecrets $ApiSecrets | ConvertFrom-Json

## ADD TO MICROSOFT 365 GROUPS ##

# If task is not complete, attempt to add to 365 group
foreach ($Task in $Tasks) {
    $Text = "$TaskNotes [automated task]`n`n"
    $TicketNoteDetails = ''
    $FailedNoteDetails = ''
    $Result = $true
    if ($Task.notes -like "$TaskNotes*" -and !$Task.closedFlag) {
        # Grab user ID from resolution
        $TaskResolution = $Task.resolution | ConvertFrom-Json
        $UserId = $TaskResolution.UserId
        $UserPrincipalName = $TaskResolution.UserPrincipalName

        # Gets all assigned groups
        $RolesParameters = @{
            BearerToken = $BearerToken
            UserId = $UserId
        }
        $Roles = .\AAD-GetAllAssignedRoles.ps1 @RolesParameters | ConvertFrom-Json

        # Remove from each group
        foreach ($Role in $Roles) {
            if ($Role.principalId -eq $UserId) {
                $Name = $Role.principalDisplayName
                $Arguments = @{
                    BearerToken = $BearerToken
                    UserId = $UserId
                    AppRoleId = $Role.id
                }

                $Operation = .\AAD-UnassignRole.ps1 @Arguments | ConvertFrom-Json
                $Log += "`n`n" + $Operation.Log

                # Update task upon completion
                if ($Operation.Result) {
                    # Add details for ticket note
                    $TicketNoteDetails += "`n- $Name"            
                } else { 
                    $Result = $false
                    $FailedNoteDetails += "`n- $Name" 
                }
            }
        }

        if ($Result) {
            # Close task
            $TaskId = $Task.id
            .\CWM-UpdateSpecificTask.ps1 -TicketId $TicketId -TaskId $TaskId -ClosedStatus $true -ApiSecrets $ApiSecrets | Out-Null
        }

        # Add note if the user was added to any distribution groups
        if ($TicketNoteDetails -ne '') {
            $Text += "The user ($UserPrincipalName) has been removed from the following Azure roles:$TicketNoteDetails"
        } else {
            $Text += "The user ($UserPrincipalName) has been not removed from any Azure roles."
        }
        if ($FailedNoteDetails -ne '') {
            $Text += "`n`nAutomation was unable to unassign the following roles:$FailedNoteDetails"
        }
        .\CWM-AddTicketNote.ps1 -TicketId $TicketId -Text $Text -ResolutionFlag $true -ApiSecrets $ApiSecrets | Out-Null
    }
}



## SEND DETAILS TO FLOW ##

$Output = @{
    Log = $Log
}

Write-Output $Output | ConvertTo-Json