<#

Mangano IT - Exit User Flow - Unassign Licenses (AAD)
Created by: Gabriel Nugent
Version: 1.3.6

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [int]$TicketId,
    [string]$BearerToken,
    [string]$TenantUrl,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$TaskNotes = 'Unassign Licenses'
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

$Tasks = .\CWM-FindTicketTasks.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets | ConvertFrom-Json

## ADD TO MICROSOFT 365 GROUPS ##

# If task is not complete, attempt to add to 365 group
foreach ($Task in $Tasks) {
    if ($Task.notes -like "$TaskNotes*" -and !$Task.closedFlag) {
        # Grab user ID from resolution
        $TaskResolution = $Task.resolution | ConvertFrom-Json
        $UserId = $TaskResolution.UserId
        $UserPrincipalName = $TaskResolution.UserPrincipalName

        # Log task info
        Write-Warning "Matching task located for $UserPrincipalName"

        # Unassigns licenses
        $LicensesParameters = @{
            BearerToken = $BearerToken
            UserId = $UserId
            UserPrincipalName = $UserPrincipalName
        }

        $Operation = .\AAD-UnassignLicenses.ps1 @LicensesParameters | ConvertFrom-Json
        $Log += $Operation.Log
        $LicenseList = $Operation.LicenseList

        # Define arguments for ticket note and task
        $TaskNoteArguments = @{
            Result = $Operation.Result
            TicketId = $TicketId
            TaskNotes = $Task.notes
            TaskId = $Task.id
            TicketNote_Success = "Automation has unassigned the following licenses from $UserPrincipalName : $LicenseList"
            TicketNote_Failure = "Automation was NOT able to unassign the following licenses from $UserPrincipalName : $LicenseList"
            ApiSecrets = $ApiSecrets
        }

        if ($LicenseList -eq '') {
            $TaskNoteArguments.TicketNote_Success = "$UserPrincipalName had no licenses that needed to be unassigned."
        }
        
        # Add note and update task if successful
        $TaskAndNote = .\CWM-UpdateTaskAddNoteForFlow.ps1 @TaskNoteArguments
        $Log += $TaskAndNote.Log
    }
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Log = $Log
}

Write-Output $Output | ConvertTo-Json