<#

Mangano IT - New User Flow - Set Usage Location
Created by: Gabriel Nugent
Version: 1.61

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory=$true)][int]$TicketId,
    [Parameter(Mandatory=$true)][string]$UserPrincipalName,
    [string]$TenantUrl,
    [string]$BearerToken,
    [string]$StatusToContinueFlow,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$TaskNotes = 'Set Usage Location'
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

## SET USAGE LOCATION ##

foreach ($Task in $Tasks) {
    $Operation = $null
    $Attempts = 0
    if (!$Task.closedFlag -and $Task.notes -eq $TaskNotes) {
        $Resolution = $Task.resolution
        $Arguments = @{
            BearerToken = $BearerToken
            UserPrincipalName = $UserPrincipalName
            UsageLocation = $Resolution
        }
        while ($Operation.Result -ne $true -and $Attempts -lt 10) {
            $Operation = .\AAD-SetUsageLocation.ps1 @Arguments | ConvertFrom-Json
            $Log += "`n`n" + $Operation.Log
            if ($Operation.Result -ne $true) {
                $Attempts += 1
                Start-Sleep -Seconds 60
            }
        }

        # Define arguments for ticket note and task
        $TaskNoteArguments = @{
            Result = $Operation.Result
            TicketId = $TicketId
            TaskNotes = $TaskNotes
            TicketNote_Success = "The usage location for $UserPrincipalName was set to $Resolution."
            TicketNote_Failure = "ERROR! The usage location for $UserPrincipalName was not set to $Resolution.`n`nWithout a usage location set, assigning a license will not work. Please set the usage location manually."
            ApiSecrets = $ApiSecrets
        }
        
        # Add note and update task if successful
        $TaskAndNote = .\CWM-UpdateTaskAddNoteForFlow.ps1 @TaskNoteArguments
        $Log += $TaskAndNote.Log
    }
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Result = $Operation.Result
    Log = $Log
}

Write-Output $Output | ConvertTo-Json