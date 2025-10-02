<#

Mangano IT - New User Flow - Add New User to Groups (AAD)
Created by: Gabriel Nugent
Version: 1.11.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [int]$TicketId,
    [string]$UserId,
    [string]$BearerToken,
    [string]$TenantUrl,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$TicketNoteDetails = ''
$FailedNoteDetails = ''
$TaskNotes = 'Add to Microsoft 365/AAD Security Group'
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
    if (!$Task.closedFlag -and $Task.notes -eq $TaskNotes) {
        $Resolution = $Task.resolution | ConvertFrom-Json
        $Name = $Resolution.Name
        $Arguments = @{
            BearerToken = $BearerToken
            UserId = $UserId
            SecurityGroupName = $Name
        }
        $Operation = .\AAD-AddToGroup.ps1 @Arguments | ConvertFrom-Json
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

# Add note if the user was added to any 365 groups
$Text = $TaskNotes + " [automated task]`n`n"

if ($TicketNoteDetails -ne '' -or $FailedNoteDetails -ne '') {
    if ($TicketNoteDetails -ne '') {
        $Text += "`n`nThe new user has been assigned access to the following Microsoft 365/AAD security groups:$TicketNoteDetails"
    }
    if ($FailedNoteDetails -ne '') {
        $Text += "`n`nThe new user has NOT been assigned access to the following Microsoft 365/AAD security groups:$FailedNoteDetails"
    }
    .\CWM-AddTicketNote.ps1 -TicketId $TicketId -Text $Text -ResolutionFlag $true -ApiSecrets $ApiSecrets | Out-Null
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Log = $Log
}

Write-Output $Output | ConvertTo-Json