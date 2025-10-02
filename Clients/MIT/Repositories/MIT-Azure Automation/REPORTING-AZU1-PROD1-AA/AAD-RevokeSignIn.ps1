<#

Mangano IT - Azure Active Directory - Revoke Sign In Sessions for AAD User
Created by: Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$UserId,
    [string]$UserPrincipalName,
    [string]$BearerToken,
    [string]$TenantUrl,
    [int]$TicketId,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

$TaskNotes = 'Revoke Sign In Sessions in AAD'

## CHECKS AND BALANCES FOR USER/GROUP DETAILS ##

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    $Log += "Bearer token not supplied. Getting bearer token for $TenantUrl...`n`n"
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

# Get user ID if not provided
if ($UserId -eq '') {
    $User = .\AAD-CheckUserExists.ps1 -UserPrincipalName $UserPrincipalName -BearerToken $BearerToken | ConvertFrom-Json
    $UserId = $User.UserId
    $Log += $User.Log + "`n`n"
}

## REVOKE SIGN IN ##

$RevokeSignInArguments = @{
    Uri = "https://graph.microsoft.com/v1.0/users/$UserId/revokeSignInSessions"
    Method = 'POST'
    Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
    UseBasicParsing = $true
}

try {
    $Log += "Revoking sign in sessions for $UserId...`n"
    Invoke-WebRequest @RevokeSignInArguments | Out-Null
    $Log += "SUCCESS: Revoked sign in sessions for $UserId.`n`n"
    Write-Warning "SUCCESS: Revoked sign in sessions for $UserId."
    $Result = $true
} catch {
    $Log += "ERROR: Unable to revoke sign in sessions for $UserId.`nERROR DETAILS: "
    Write-Error "Unable to revoke sign in sessions for $UserId : $_"
    $Result = $false
}

## UPDATE TICKET IF PROVIDED ##

if ($TicketId -ne 0) {
    # Define arguments for ticket note and task
    $TaskNoteArguments = @{
        Result = $Result
        TicketId = $TicketId
        TaskNotes = $TaskNotes
        TicketNote_Success = "$UserId has had their sign in sessions revoked."
        TicketNote_Failure = "ERROR: $UserId has NOT had their sign in sessions revoked."
        ApiSecrets = $ApiSecrets
    }
    
    # Add note and update task if successful
    $TaskAndNote = .\CWM-UpdateTaskAddNoteForFlow.ps1 @TaskNoteArguments
    $Log += $TaskAndNote.Log
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Result = $Result
    UserId = $UserId
    Log = $Log
}

Write-Output $Output | ConvertTo-Json -Depth 100