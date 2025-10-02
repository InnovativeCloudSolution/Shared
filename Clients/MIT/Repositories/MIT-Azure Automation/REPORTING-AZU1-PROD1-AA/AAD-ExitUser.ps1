<#

Mangano IT - Azure Active Directory - Exit AAD User
Created by: Gabriel Nugent
Version: 1.1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [int]$TicketId,
    [string]$UserId,
    [string]$UserPrincipalName,
    [string]$BearerToken,
    [string]$TenantUrl,
    [bool]$RemoveManager = $true,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

[string]$Log = ''
$Result = $false
$TaskNotes = "Disable User Account"

## CHECKS AND BALANCES FOR USER/GROUP DETAILS ##

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    $Log += "Bearer token not supplied. Getting bearer token for $TenantUrl...`n`n"
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

# Get user ID if not provided
if ($UserId -eq '' -and $UserPrincipalName -ne '') {
    $User = .\AAD-CheckUserExists.ps1 -UserPrincipalName $UserPrincipalName -BearerToken $BearerToken -SkipExtraAttempts $true | ConvertFrom-Json
    $UserId = $User.UserId
    $Log += $User.Log + "`n`n"
} elseif ($UserPrincipalName -eq '' -and $UserId -ne '') {
    $User = .\AAD-CheckUserExists.ps1 -UserId $UserId -BearerToken $BearerToken -SkipExtraAttempts $true | ConvertFrom-Json
    $UserPrincipalName = $User.UserPrincipalName
    $Log += $User.Log + "`n`n"
}

## DISABLE USER ACCOUNT ##

if ($UserId -ne '') {
    # Build arguments for API request
    $ApiArguments = @{
        Uri = "https://graph.microsoft.com/v1.0/users/$UserId"
        Method = 'PATCH'
        Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
        Body = @{ accountEnabled = $false } | ConvertTo-Json
        UseBasicParsing = $true
    }

    try {
        $Log += "Attempting to disable user account $UserPrincipalName...`n"
        Invoke-WebRequest @ApiArguments | Out-Null
        $Log += "SUCCESS: User account $UserPrincipalName disabled.`n`n"
        Write-Warning "SUCCESS: User account $UserPrincipalName disabled."
        $Result = $true
    } catch {
        $Log += "ERROR: Unable to disable user account $UserPrincipalName.`n`nERROR DETAILS: " + $_
        Write-Error "Unable to disable user account $UserPrincipalName : $_"
        $Result = $false
    }
}

## REMOVE MANAGER ##

if ($RemoveManager) {
    # Build arguments for API request
    $RemoveManagerArguments = @{
        Uri = "https://graph.microsoft.com/v1.0/users/$UserId/manager/" + '$ref'
        Method = 'DELETE'
        Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
        UseBasicParsing = $true
    }

    try {
        $Log += "Attempting to remove the user's manager...`n"
        Invoke-WebRequest @RemoveManagerArguments | Out-Null
        $Log += "SUCCESS: User account $UserPrincipalName has had their manager removed.`n`n"
        Write-Warning "SUCCESS: User account $UserPrincipalName has had their manager removed."
        $Result = $true
    } catch {
        $Log += "ERROR: Unable to remove manager from $UserPrincipalName.`n`nERROR DETAILS: " + $_
        Write-Error "Unable to remove manager from $UserPrincipalName : $_"
        $Result = $false
    }
}

## UPDATE TICKET IF PROVIDED ##

if ($TicketId -ne 0) {
    # Define arguments for ticket note and task
    $TaskNoteArguments = @{
        Result = $Result
        TicketId = $TicketId
        TaskNotes = $TaskNotes
        TicketNote_Success = "User $UserPrincipalName has been disabled in Azure Active Directory."
        TicketNote_Failure = "ERROR: User $UserPrincipalName has NOT been disabled in Azure Active Directory."
        ApiSecrets = $ApiSecrets
    }
    
    # Add note and update task if successful
    $TaskAndNote = .\CWM-UpdateTaskAddNoteForFlow.ps1 @TaskNoteArguments
    $Log += $TaskAndNote.Log
}

## SEND DETAILS BACK TO FLOW ##

$Output = [ordered]@{
    Result = $Result
    UserId = $UserId
    UserPrincipalName = $UserPrincipalName
    Log = $Log
}

Write-Output $Output | ConvertTo-Json