<#

Mangano IT - ConnectWise Manage - Deactivate Contact
Created by: Gabriel Nugent
Version: 1.3

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param (
	[Parameter(Mandatory=$true)][string]$CompanyName,
    [Parameter(Mandatory=$true)][string]$EmailAddress,
    [int]$ContactId,
    $ApiSecrets = $null,
    [int]$TicketId
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

$TaskNotes = "Make Contact Inactive in ConnectWise Manage"
$ContentType = 'application/json'

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## SETUP API VARIABLES ##

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

## FIND CONTACT IF NEEDED ##

if ($ContactId -eq 0) {
    $FindContactArguments = @{
        EmailAddress = $EmailAddress
        CompanyName = $CompanyName
        ApiSecrets = $ApiSecrets
    }

    $Contact = .\CWM-FindContactByEmail.ps1 @FindContactArguments | ConvertFrom-Json

    if ($null -ne $Contact.id) {
        $ContactId = $Contact.id
        $Log += "SUCCESS: Contact located. Contact ID: $ContactId`n"
        Write-Warning "SUCCESS: Contact located. Contact ID: $ContactId"
    } else { $ContactId = 0 }
}

## DEACTIVATE CONTACT ##

if ($ContactId -ne 0) {
    $ApiBody = @(
        @{
            op = 'replace'
            path = '/inactiveFlag'
            value = $true
        }
    )

    $ApiArguments = @{
        Uri = "$CWMApiUrl/company/contacts/$ContactId"
        Method = 'PATCH'
        Body = ConvertTo-Json -InputObject $ApiBody -Depth 100
        ContentType = $ContentType
        Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
        UseBasicParsing = $true
    }

    try {
        $Log += "Marking $ContactId as inactive...`n"
        Write-Warning "Marking $ContactId as inactive..."
        Invoke-WebRequest @ApiArguments | Out-Null
        $Log += "SUCCESS: $ContactId has been marked as inactive."
        Write-Warning "SUCCESS: $ContactId has been marked as inactive."
        $Result = $true
    } catch {
        $Log += "ERROR: Unable to mark $ContactId as inactive.`nERROR DETAILS: " + $_
        Write-Error "Unable to mark $ContactId as inactive : $_"
        $Result = $false
    }
} else {
    $Result = $false
}

## UPDATE TICKET IF PROVIDED ##

if ($TicketId -ne 0) {
    # Define arguments for ticket note and task
    $TaskNoteArguments = @{
        Result = $Result
        TicketId = $TicketId
        TaskNotes = $TaskNotes
        TicketNote_Success = "The contact for $EmailAddress ($ContactId) has been marked as inactive."
        TicketNote_Failure = "ERROR: The contact for $EmailAddress has NOT been marked as inactive."
        ApiSecrets = $ApiSecrets
    }

    if ($ContactId -eq 0) {
        $TaskNoteArguments.TicketNote_Failure += "`n`nNo contact ID was found for the provided email address."
    }
    
    # Add note and update task if successful
    $TaskAndNote = .\CWM-UpdateTaskAddNoteForFlow.ps1 @TaskNoteArguments
    $Log += $TaskAndNote.Log
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Result = $Result
    Log = $Log
}

Write-Output $Output | ConvertTo-Json