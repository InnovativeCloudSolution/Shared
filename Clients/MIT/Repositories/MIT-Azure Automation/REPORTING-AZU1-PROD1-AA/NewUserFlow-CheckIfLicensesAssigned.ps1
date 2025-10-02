<#

Mangano IT - New User Flow - Check if Licenses have been Assigned
Created by: Gabriel Nugent
Version: 1.2

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory=$true)][int]$TicketId,
    [Parameter(Mandatory=$true)][string]$UserId,
    [string]$TenantUrl,
    [string]$BearerToken,
    [string]$StatusToContinueFlow,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$TicketNoteDetails = ''
$TaskNotes = 'Confirm License is Assigned'
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

## CHECK LICENSE ASSIGNED STATUS ##

foreach ($Task in $Tasks) {
    if ($Task.notes -eq $TaskNotes) {
        $TaskId = $Task.id
        $Resolution = $Task.resolution | ConvertFrom-Json
        $Name = $Resolution.Name
        $SkuPartNumber = $Resolution.SkuPartNumber
        $CheckLicenseArguments = @{
            UserId = $UserId
            SkuPartNumber = $SkuPartNumber
            BearerToken = $BearerToken
        }
        $Operation = .\AAD-CheckIfLicenseAssigned.ps1 @CheckLicenseArguments | ConvertFrom-Json
        $Log += "`n`n" + $Operation.Log

        # Close task if license assigned, re-open if not
        if ($Operation.Result) {
            .\CWM-UpdateSpecificTask.ps1 -TicketId $TicketId -TaskId $TaskId -ClosedStatus $true -ApiSecrets $ApiSecrets | Out-Null
        }
        else {
            .\CWM-UpdateSpecificTask.ps1 -TicketId $TicketId -TaskId $TaskId -ClosedStatus $false -ApiSecrets $ApiSecrets | Out-Null
            $TicketNoteDetails += "`n- $Name"
        }
    }
}

# Set note and status details depending on result
if ($TicketNoteDetails -ne '') {
    $Text = $TaskNotes + " [automated task]`n`n"
    $Text += "The following licenses were not found assigned to the account:$TicketNoteDetails`n`n"
    $Text += "This may be due to a deficit of license(s) for group-based licensing to apply.`n`n"
    $Text += 'Once the licenses are available in the tenancy and are assigned to the account, please set the ticket to "' + $StatusToContinueFlow + '".'

    $StatusArguments = @{
        TicketId = $TicketId 
        StatusName = "Scheduling Required"
        CustomerRespondedFlag = $true 
        ApiSecrets = $ApiSecrets
    }
    $LicensesAssigned = $false
} else {
    $Text = "$TaskNotes [automated task]`n`n"
    $Text += "A check has been performed against the new user, and has confirmed that the required licenses have been assigned. "
    $Text += "The automation will now continue."

    $StatusArguments = @{
        TicketId = $TicketId 
        StatusName = "Automation in Progress"
        CustomerRespondedFlag = $true 
        ApiSecrets = $ApiSecrets
    }
    $LicensesAssigned = $true
}
$NoteArguments = @{
    TicketId = $TicketId
    Text = $Text
    ResolutionFlag = $true
    ApiSecrets = $ApiSecrets
}

.\CWM-AddTicketNote.ps1 @NoteArguments | Out-Null
.\CWM-UpdateTicketStatus.ps1 @StatusArguments | Out-Null

## SEND DETAILS TO FLOW ##

$Output = @{
    LicensesAssigned = $LicensesAssigned
    Log = $Log
}

Write-Output $Output | ConvertTo-Json