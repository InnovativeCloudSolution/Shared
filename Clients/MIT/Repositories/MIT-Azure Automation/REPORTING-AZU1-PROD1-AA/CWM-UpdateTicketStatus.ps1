<#

Mangano IT - ConnectWise Manage - Update Ticket Status
Created by: Gabriel Nugent
Version: 1.3

This runbook is designed to be used in conjunction with a Power Automate flow or an Azure Automation script.

#>

param (
    [Parameter(Mandatory=$true)][int]$TicketId,
    [Parameter(Mandatory=$true)][string]$StatusName,
    [bool]$CustomerRespondedFlag = $false,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

[string]$Log = ''
$ContentType = 'application/json'
$StatusId = 0

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
    Write-Warning "Attempting to pull ticket details for #$TicketId..."
    $Ticket = Invoke-WebRequest @GetTicketArguments | ConvertFrom-Json
    $Log += "SUCCESS: Ticket details pulled for #$TicketId.`n`n"
    Write-Warning "SUCCESS: Ticket details pulled for #$TicketId."
} catch {
    $Ticket = $null
    $Log += "ERROR: Unable to pull ticket details for #$TicketId.`nERROR DETAILS: " + $_
    Write-Error "Unable to pull ticket details for #$TicketId : $_"
    $Result = $false
}

# If ticket isn't empty, get statuses for board
if ($null -ne $Ticket) {
    $BoardId = $Ticket.board.id
    $GetStatusesArguments = @{
        Uri = "$CWMApiUrl/service/boards/$BoardId/statuses?conditions=name like '" + $StatusName + "'"
        Method = 'GET'
        ContentType = $ContentType
        Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
        UseBasicParsing = $true
    }

    try { 
        $Log += "Attempting to pull statuses for board $BoardId...`n"
        Write-Warning "Attempting to pull statuses for board $BoardId..."
        $Statuses = Invoke-WebRequest @GetStatusesArguments | ConvertFrom-Json
        $Log += "SUCCESS: Statuses pulled for board $BoardId.`n"
        Write-Warning "SUCCESS: Statuses pulled for board $BoardId."
    } catch {
        $Statuses = $null
        $Log += "ERROR: Statuses not pulled for board $BoardId.`nERROR DETAILS: " + $_
        Write-Error "Statuses not pulled for board $BoardId : $_"
        $Result = $false
    }

    # If statuses aren't empty, find status ID
    if ($null -ne $Statuses) {
        foreach ($Status in $Statuses) {
            if ($Status.name -like $StatusName) {
                $StatusId = $Status.id
                $Log += "SUCCESS: $StatusName located. Status ID: $StatusId.`n`n"
                Write-Warning "SUCCESS: $StatusName located. Status ID: $StatusId."
            }
        }

        ## UPDATE TICKET STATUS ##

        # Build API body
        $ApiBody = @(
            @{
                op = "replace"
                path = "/status/id"
                value = $StatusId
            },
            @{
                op = "replace"
                path = "/customerUpdatedFlag"
                value = $CustomerRespondedFlag
            }
        )

        # Build API arguments
        $ApiArguments = @{
            Uri = "$CWMApiUrl/service/tickets/$TicketId"
            Method = 'PATCH'
            ContentType = $ContentType
            Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
            Body = ConvertTo-Json -InputObject $ApiBody -Depth 100
            UseBasicParsing = $true
        }

        try {
            $Log += "Updating #$TicketId to status $StatusId...`n"
            Invoke-WebRequest @ApiArguments | Out-Null
            $Log += "SUCCESS: #$TicketId has been set to status $StatusId."
            Write-Warning "SUCCESS: #$TicketId has been set to status $StatusId."
            $Result = $true
        } catch {
            $Log += "ERROR: Unable to update #$TicketId to status $StatusId.`nERROR DETAILS: " + $_
            Write-Error "Unable to update #$TicketId to status $StatusId : $_"
            $Result = $false
        }
    }
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Result = $Result
    Log = $Log
}

Write-Output $Output | ConvertTo-Json