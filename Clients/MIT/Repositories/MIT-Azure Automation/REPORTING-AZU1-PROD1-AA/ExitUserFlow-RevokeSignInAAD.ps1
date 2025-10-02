<#

Mangano IT - Exit User Flow - Revoke Sign In (AAD)
Created by: Gabriel Nugent
Version: 1.1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [int]$TicketId,
    [string]$BearerToken,
    [string]$TenantUrl,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$TaskNotes = 'Revoke Sign In Sessions in AAD'
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

## REVOKE SIGN IN ##

# If task is not complete, attempt to revoke sign in
foreach ($Task in $Tasks) {
    $Text = "$TaskNotes [automated task]`n`n"
    if ($Task.notes -like "$TaskNotes*" -and !$Task.closedFlag) {
        # Grab user ID and UPN from resolution
        $TaskResolution = $Task.resolution | ConvertFrom-Json
        $UserId = $TaskResolution.UserId
        $UserPrincipalName = $TaskResolution.UserPrincipalName

        # Log task info
        Write-Warning "Matching task located for $UserPrincipalName"

        # Revokes sign in
        $RevokeParameters = @{
            BearerToken = $BearerToken
            UserId = $UserId
        }
        
        $Operation = .\AAD-RevokeSignIn.ps1 @RevokeParameters | ConvertFrom-Json

        if ($Operation.Result) {
            # Close task
            $TaskId = $Task.id
            .\CWM-UpdateSpecificTask.ps1 -TicketId $TicketId -TaskId $TaskId -ClosedStatus $true -ApiSecrets $ApiSecrets | Out-Null
            $Text += "The user ($UserPrincipalName) has had their sign in sessions revoked in AAD."
        } else {
            $Text += "The user ($UserPrincipalName) has not had their sign in sessions revoked in AAD."
        }

        .\CWM-AddTicketNote.ps1 -TicketId $TicketId -Text $Text -ResolutionFlag $true -ApiSecrets $ApiSecrets | Out-Null
    }
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Log = $Log
}

Write-Output $Output | ConvertTo-Json