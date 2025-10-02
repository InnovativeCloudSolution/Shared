<#

Mangano IT - ConnectWise Manage - Add Note to Ticket
Created by: Gabriel Nugent
Version: 1.2

This runbook is designed to be used in conjunction with a Power Automate flow or an Azure Automation script.

#>

param (
    [Parameter(Mandatory=$true)][int]$TicketId,
    [Parameter(Mandatory=$true)][string]$Text,
    [bool]$DiscussionFlag = $false,
    [bool]$InternalFlag = $false,
    [bool]$ResolutionFlag = $false,
    [bool]$IssueFlag = $false,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

[string]$Log = ''
$ContentType = 'application/json'
$ContactId = 15655 # Automation Bot contact

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## SETUP API VARIABLES ##

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

# Set internal/external based on selected flags
if ($DiscussionFlag -or $ResolutionFlag) { $ExternalFlag = $true }
else { $ExternalFlag = $false }

# Build API body
$ApiBody = @{
    text = $Text
    detailDescriptionFlag = $DiscussionFlag
    internalAnalysisFlag = $InternalFlag
    resolutionFlag = $ResolutionFlag
    issueFlag = $IssueFlag
    internalFlag = $InternalFlag
    externalFlag = $ExternalFlag
    contact = @{
        id = $ContactId
    }
}

# Build API arguments
$ApiArguments = @{
    Uri = "$CWMApiUrl/service/tickets/$TicketId/notes"
    Method = 'POST'
    ContentType = $ContentType
    Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
    Body = $ApiBody | ConvertTo-Json -Depth 100
    UseBasicParsing = $true
}

## ADD NOTE TO TICKET ##

try {
    $Log += "Adding note to ticket #$TicketId...`n"
    Invoke-WebRequest @ApiArguments | Out-Null
    $Result = $true
    $Log += "SUCCESS: Note added to ticket #$TicketId."
    Write-Warning "SUCCESS: Note added to ticket #$TicketId."
} catch {
    $Result = $false
    $Log += "ERROR: Note not added to ticket #$TicketId.`nERROR DETAILS: " + $_
    Write-Error "Note not added to ticket #$TicketId : $_"
}

## SEND OUTPUT TO FLOW ##

$Output = @{
    Result = $Result
    Log = $Log
}

Write-Output $Output | ConvertTo-Json