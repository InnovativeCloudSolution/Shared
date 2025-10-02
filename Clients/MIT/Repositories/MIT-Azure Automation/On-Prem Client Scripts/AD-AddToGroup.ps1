<#

Mangano IT - Active Directory - Add User to Group (AD)
Created by: Gabriel Nugent
Version: 1.2

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$SamAccountName = '',
    [string]$SecurityGroupName = '',
    [bool]$ForceADSync = $false,
    [string]$AzureADServer
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

# Output the status variable content to a file
$Date = Get-Date -Format "dd-MM-yyyy HHmm"
$FilePath = "C:\Scripts\Logs\AD-AddToGroup"
$FileName = "$SamAccountName - $SecurityGroupName-$Date.txt"

$Result = $false

## ADD TO GROUP ##

try {
    $Log += "Adding $SamAccountName to $SecurityGroupName...`n"
    Add-ADGroupMember -Identity $SecurityGroupName -Members $SamAccountName | Out-Null
    $Log += "SUCCESS: $SamAccountName has been added to $SecurityGroupName."
    Write-Warning "SUCCESS: $SamAccountName has been added to $SecurityGroupName."
    $Result = $true
} catch {
    $Log += "ERROR: Unable to add $SamAccountName to $SecurityGroupName.`nERROR DETAILS: " + $_
    Write-Error "Unable to add $SamAccountName to $SecurityGroupName : $_"
}

## RUN AD SYNC ##

if ($ForceADSync -and $Result) {
    $SyncResult = .\AD-RunADSync.ps1 -AzureADServer $AzureADServer
    $Log += "`n`nINFO: ADSync $SyncResult"
    Write-Warning "AD sync result: $SyncResult"
}

## SEND DETAILS BACK TO FLOW ##

$Output = @{
    Result = $Result
    SyncResult = $SyncResult
    Log = $Log
    LogFile = "$FilePath\$FileName"
}

Write-Output $Output | ConvertTo-Json

# Makes folder for logs and outputs logs
.\CreateLogFile.ps1 -Log $Log -FilePath $FilePath -FileName $FileName | Out-Null