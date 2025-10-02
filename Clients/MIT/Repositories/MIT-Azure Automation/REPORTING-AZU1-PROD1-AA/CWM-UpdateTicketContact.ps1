<#

Mangano IT - ConnectWise Manage - Update Ticket Contact
Created by: Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow or an Azure Automation script.

#>

param (
    [Parameter(Mandatory=$true)][int]$TicketId,
    [Parameter(Mandatory=$true)][int]$ContactId,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

[string]$Log = ''
$ContentType = 'application/json'

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## SETUP API VARIABLES ##

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

## GET REQUESTED TICKET STATUS ##

# Grabs the ticket details to get the board
$GetTicketArguments = @{
    Uri = "$CWMApiUrl/service/tickets/$TicketId"
    Method = 'GET'
    ContentType = $ContentType
    Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
    UseBasicParsing = $true
}

try {
    $Log += "Attempting to pull ticket details for #$TicketId...`n"
    $Ticket = Invoke-WebRequest @GetTicketArguments | ConvertFrom-Json
    $Log += "SUCCESS: Ticket details pulled for #$TicketId.`n`n"
    Write-Warning "SUCCESS: Ticket details pulled for #$TicketId."
} catch {
    $Ticket = $null
    $Log += "ERROR: Unable to pull ticket details for #$TicketId.`nERROR DETAILS: " + $_
    Write-Error "Unable to pull ticket details for #$TicketId : $_"
    $Result = $false
}

# If ticket isn't empty, get update the ticket contact
if ($null -ne $Ticket) {
    $ApiBody = @(
        @{
            op = 'replace'
            path = '/contact/id'
            value = $ContactId
        }
    )

    $ApiArguments = @{
        Uri = "$CWMApiUrl/service/tickets/$TicketId"
        Method = 'PATCH'
        Body = ConvertTo-Json -InputObject $ApiBody -Depth 100
        ContentType = $ContentType
        Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
        UseBasicParsing = $true
    }

    try {
        $Log += "Updating #$TicketId's to $ContactId...`n"
        Invoke-WebRequest @ApiArguments | Out-Null
        $Log += "SUCCESS: #$TicketId's contact has been set to $ContactId."
        Write-Warning "SUCCESS: #$TicketId's contact has been set to $ContactId."
        $Result = $true
    } catch {
        $Log += "ERROR: Unable to set #$TicketId's contact to $ContactId.`nERROR DETAILS: " + $_
        Write-Error "Unable to set #$TicketId's contact to $ContactId : $_"
        $Result = $false
    }
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Result = $Result
    Log = $Log
}

Write-Output $Output | ConvertTo-Json