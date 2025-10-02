<#

Mangano IT - Exit User Flow - Check For and Disable Admin Account/s (AAD)
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [int]$TicketId,
    [string]$BearerToken,
    [string]$TenantUrl,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$TaskNotes = 'Check For and Disable Admin Account/s'
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
    $Text = $TaskNotes + " [automated task]"
    $Result = $true
    if ($Task.notes -like "$TaskNotes*" -and !$Task.closedFlag) {
        # Grab user details from resolution
        $TaskResolution = $Task.resolution | ConvertFrom-Json
        $UserPrincipalName = $TaskResolution.UserPrincipalName
        $AdminPrefix = $TaskResolution.$AdminPrefix

        # Attempt to locate user
        $AdminParameters = @{
            UserPrincipalName = "$AdminPrefix$UserPrincipalName"
            BearerToken = $BearerToken
            SkipExtraAttempts = $true
        }
        $AdminUser = .\AAD-CheckUserExists.ps1 @AdminParameters | ConvertFrom-Json

        # Mark as successful if no user found, or run exit and add tasks if found
        if (!$AdminUser.Result) {
            $Text += "`n`nBased on the standard prefix ($AdminPrefix), the user ($UserPrincipalName) does not have any admin accounts to remove."
            $Result = $true
        } else {
            $AdminUserPrincipalName = $AdminUser.UserPrincipalName
            $AdminUserId = $AdminUser.UserId

            # Create tasks for admin user's exit process
            $AdminTasksParameters = @{
                TicketId = $TicketId
                AadUserId = $AdminUserId
                AadUserPrincipalName = $AdminUserPrincipalName
                ApiSecrets = $ApiSecrets
            }

            .\ExitUserFlow-CreateAdminTasks.ps1 @AdminTasksParameters | Out-Null
            
            # Exit admin user
            if ($AdminUser.AccountEnabled) {
                $ExitUserParameters = @{
                    TicketId = $TicketId
                    UserId = $AdminUserId
                    UserPrincipalName = $AdminUserPrincipalName
                    BearerToken = $BearerToken
                    TenantUrl = $TenantUrl
                    ApiSecrets = $ApiSecrets
                }
                $AdminExit = .\AAD-ExitUser.ps1 @ExitUserParameters | ConvertFrom-Json
                $Result = $AdminExit.Result
            } else {
                $Text += "`n`n$AdminUserPrincipalName has already been disabled."
                $Result = $true
            }

            $Text += "`n`nTasks have been added to the ticket for the exit of $AdminUserPrincipalName."
        }

        # Close task if successful
        if ($Result) {
            $TaskId = $Task.id
            .\CWM-UpdateSpecificTask.ps1 -TicketId $TicketId -TaskId $TaskId -ClosedStatus $true -ApiSecrets $ApiSecrets | Out-Null
        }

        .\CWM-AddTicketNote.ps1 -TicketId $TicketId -Text $Text -ResolutionFlag $true -ApiSecrets $ApiSecrets | Out-Null
    }
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Result = $Result
    Log = $Log
}

Write-Output $Output | ConvertTo-Json