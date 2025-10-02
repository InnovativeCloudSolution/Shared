<#

Mangano IT - Exit User Flow - Remove User from Groups (AAD)
Created by: Gabriel Nugent
Version: 1.8.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [int]$TicketId,
    [string]$BearerToken,
    [string]$TenantUrl,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$TaskNotes = 'Remove from Cloud Groups'
$GroupList = @()
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
    $TicketNoteDetails = ''
    $FailedNoteDetails = ''
    $Text = $TaskNotes + " [automated task]"
    $Result = $true
    if ($Task.notes -like "$TaskNotes*" -and !$Task.closedFlag) {
        # Grab user ID from resolution
        $TaskResolution = $Task.resolution | ConvertFrom-Json
        $UserId = $TaskResolution.UserId
        $UserPrincipalName = $TaskResolution.UserPrincipalName

        # Gets all assigned groups
        $GroupsParameters = @{
            BearerToken = $BearerToken
            UserId = $UserId
        }
        $Groups = .\AAD-GetAllAssignedGroups.ps1 @GroupsParameters | ConvertFrom-Json

        # Remove from each group
        foreach ($Group in $Groups) {
            # Only attempt to remove if the group is assigned
            if ($null -eq $Group.membershipRule -and $null -eq $Group.onPremisesSyncEnabled) {
                $Name = $Group.displayName
                $Id = $Group.id
                $Arguments = @{
                    BearerToken = $BearerToken
                    UserId = $UserId
                    SecurityGroupId = $Id
                }

                $Operation = .\AAD-RemoveFromGroup.ps1 @Arguments | ConvertFrom-Json
                $Log += "`n`n" + $Operation.Log

                # Update task upon completion
                if ($Operation.Result) {
                    # Add details for ticket note
                    $TicketNoteDetails += "`n- $Name ($Id)"         
                } else {
                    $Result = $false
                    $FailedNoteDetails += "`n- $Name ($Id)"
                }

                # Add group name to list without Id
                $GroupList += "$Name ($UserPrincipalName)"
            }
        }

        if ($Result) {
            # Close task
            $TaskId = $Task.id
            .\CWM-UpdateSpecificTask.ps1 -TicketId $TicketId -TaskId $TaskId -ClosedStatus $true -ApiSecrets $ApiSecrets | Out-Null
        }

        # Add note if the user was added to any distribution groups
        if ($TicketNoteDetails -eq '' -and $FailedNoteDetails -eq '' -and $Result) {
            $Text += "`n`nThe user ($UserPrincipalName) has not been removed from any Microsoft 365/AAD Security groups."
        } else {
            if ($TicketNoteDetails -ne '') {
                $Text += "`n`nThe user ($UserPrincipalName) has been removed from the following Microsoft 365/AAD Security groups:$TicketNoteDetails"
            }
            if ($FailedNoteDetails) {
                $Text += "`n`nThe user ($UserPrincipalName) has NOT been removed from the following Microsoft 365/AAD Security groups:$FailedNoteDetails`n`n"
                $Text += "`n`nThe automation has no way to filter distribution lists from regular 365 groups. If any of the groups on the above list are actually"
                $Text += " distribution lists, the Exchange section of the automation should handle them."
            }
        }

        .\CWM-AddTicketNote.ps1 -TicketId $TicketId -Text $Text -ResolutionFlag $true -ApiSecrets $ApiSecrets | Out-Null
    }
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Log = $Log
    Microsoft365Groups = $GroupList
}

Write-Output $Output | ConvertTo-Json