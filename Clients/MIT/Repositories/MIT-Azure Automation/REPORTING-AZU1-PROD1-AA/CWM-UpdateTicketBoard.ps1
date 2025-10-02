<#

Mangano IT - ConnectWise Manage - Update Ticket Board
Created by: Gabriel Nugent
Version: 1.2

This runbook is designed to be used in conjunction with a Power Automate flow or an Azure Automation script.

#>

param (
    [Parameter(Mandatory=$true)][int]$TicketId,
    [Parameter(Mandatory=$true)][string]$BoardName,
    [string]$TypeName,
    [string]$SubtypeName,
    [string]$ItemName,
    [string]$StatusName,
    [string]$LevelName,
    [bool]$CustomerRespondedFlag = $false,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

$ContentType = 'application/json'
$BoardId = 0
$TypeId = 0
$SubtypeId = 0
$ItemId = 0
$StatusId = 0

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## SETUP API VARIABLES ##

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

## GET REQUESTED BOARD DETAILS ##

# Grabs the ticket details to get the board
$GetBoardsArguments = @{
    Uri = "$CWMApiUrl/service/boards?conditions=name like '" + $BoardName + "'"
    Method = 'GET'
    ContentType = $ContentType
    Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
    UseBasicParsing = $true
}

try {
    $Log += "Pulling all boards..`n"
    $Boards = Invoke-WebRequest @GetBoardsArguments | ConvertFrom-Json
    $Log += "SUCCESS: All boards pulled.`n`n"
    Write-Warning "SUCCESS: All boards pulled."
} catch {
    $Log += "ERROR: Unable to pull boards.`nERROR DETAILS: " + $_
    Write-Error "Unable to pull boards : $_"
    $Boards = $null
    $Result = $false
}

# If boards isn't empty, find right board
if ($null -ne $Boards) {
    foreach ($Board in $Boards) {
        if ($Board.name -eq $BoardName) {
            $BoardId = $Board.id
            $Log += "INFO: Board located. ID: $BoardId.`n`n"
            Write-Warning "INFO: Board located. ID: $BoardId."
            break
        } else { $Log += "INFO: " + $Board.name + " does not match $BoardName.`n" }
    }
}

if ($BoardId -ne 0) {
    # Locate type
    $GetTypesArguments = @{
        Uri = "$CWMApiUrl/service/boards/$BoardId/types?conditions=name like '" + $TypeName + "'"
        Method = 'GET'
        ContentType = $ContentType
        Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
        UseBasicParsing = $true
    }

    try { 
        $Log += "Attempting to pull types for board $BoardId...`n"
        $Types = Invoke-WebRequest @GetTypesArguments | ConvertFrom-Json
        $Log += "SUCCESS: Types pulled for board $BoardId.`n"
    } catch {
        $Types = $null
        $Log += "ERROR: Types not pulled for board $BoardId.`nERROR DETAILS: " + $_
        Write-Error "Types not pulled for board $BoardId : $_"
        $Result = $false
        $Types = $null
    }

    if ($null -ne $Types) {
        foreach ($Type in $Types) {
            if ($Type.name -eq $TypeName) {
                $TypeId = $Type.id
                $Log += "SUCCESS: $TypeName located. Type ID: $TypeId.`n"
                Write-Warning "SUCCESS: $TypeName located. Type ID: $TypeId."
                break
            }
        }
    }

    # Locate subtype
    $GetSubtypesArguments = @{
        Uri = "$CWMApiUrl/service/boards/$BoardId/subtypes?conditions=name like '" + $SubtypeName + "'"
        Method = 'GET'
        ContentType = $ContentType
        Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
        UseBasicParsing = $true
    }

    try { 
        $Log += "Attempting to pull subtypes for board $BoardId...`n"
        $Subtypes = Invoke-WebRequest @GetSubtypesArguments | ConvertFrom-Json
        $Log += "SUCCESS: Subtypes pulled for board $BoardId.`n"
    } catch {
        $Subtypes = $null
        $Log += "ERROR: Subtypes not pulled for board $BoardId.`nERROR DETAILS: " + $_
        Write-Error "Subtypes not pulled for board $BoardId : $_"
        $Result = $false
        $Subtypes = $null
    }

    if ($null -ne $Subtypes) {
        foreach ($Subtype in $Subtypes) {
            if ($Subtype.name -eq $SubtypeName) {
                $SubtypeId = $Subtype.id
                foreach ($TypeAssociationId in $Subtype.typeAssociationIds) {
                    if ($TypeAssociationId -eq $TypeId) {
                        $Log += "SUCCESS: $SubtypeName located. Subtype ID: $SubtypeId.`n"
                        Write-Warning "SUCCESS: $SubtypeName located. Subtype ID: $SubtypeId."
                        break
                    }
                }
            }
        }
    }

    # Locate item
    $GetItemsArguments = @{
        Uri = "$CWMApiUrl/service/boards/$BoardId/items?conditions=name like '" + $ItemName + "'"
        Method = 'GET'
        ContentType = $ContentType
        Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
        UseBasicParsing = $true
    }

    try { 
        $Log += "Attempting to pull items for board $BoardId...`n"
        $Items = Invoke-WebRequest @GetItemsArguments | ConvertFrom-Json
        $Log += "SUCCESS: Items pulled for board $BoardId.`n"
    } catch {
        $Items = $null
        $Log += "ERROR: Items not pulled for board $BoardId.`nERROR DETAILS: " + $_
        Write-Error "Items not pulled for board $BoardId : $_"
        $Result = $false
        $Items = $null
    }

    if ($null -ne $Items) {
        foreach ($Item in $Items) {
            if ($Item.name -eq $ItemName) {
                $ItemId = $Item.id
                $Log += "SUCCESS: $ItemName located. Item ID: $ItemId.`n"
                Write-Warning "SUCCESS: $ItemName located. Item ID: $ItemId."
                break
            }
        }
    }

    # Locate status
    $GetStatusesArguments = @{
        Uri = "$CWMApiUrl/service/boards/$BoardId/statuses?conditions=name like '" + $StatusName + "'"
        Method = 'GET'
        ContentType = $ContentType
        Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
        UseBasicParsing = $true
    }

    try { 
        $Log += "Attempting to pull statuses for board $BoardId...`n"
        $Statuses = Invoke-WebRequest @GetStatusesArguments | ConvertFrom-Json
        $Log += "SUCCESS: Statuses pulled for board $BoardId.`n"
    } catch {
        $Statuses = $null
        $Log += "ERROR: Statuses not pulled for board $BoardId.`nERROR DETAILS: " + $_
        Write-Error "Statuses not pulled for board $BoardId : $_"
        $Result = $false
        $Statuses = $null
    }

    # If statuses aren't empty, find status ID
    if ($null -ne $Statuses) {
        foreach ($Status in $Statuses) {
            if ($Status.name -eq $StatusName) {
                $StatusId = $Status.id
                $Log += "SUCCESS: $StatusName located. Status ID: $StatusId.`n`n"
                Write-Warning "SUCCESS: $StatusName located. Status ID: $StatusId."
                break
            }
        }

        ## UPDATE TICKET STATUS ##

        # Build API body
        $ApiBody = @(
            @{
                op = "replace"
                path = "/board/id"
                value = $BoardId
            },
            @{
                op = "replace"
                path = "/customerUpdatedFlag"
                value = $CustomerRespondedFlag
            }
        )

        # Add IDs if not null
        if ($TypeId -ne 0) {
            $ApiBody += @{
                op = "replace"
                path = "/type/id"
                value = $TypeId
            }
        }

        if ($SubtypeId -ne 0) {
            $ApiBody += @{
                op = "replace"
                path = "/subType/id"
                value = $SubtypeId
            }
        }

        if ($ItemId -ne 0) {
            $ApiBody += @{
                op = "replace"
                path = "/item/id"
                value = $ItemId
            }
        }

        if ($StatusId -ne 0) {
            $ApiBody += @{
                op = "replace"
                path = "/status/id"
                value = $StatusId
            }
        }

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
            $Log += "Updating #$TicketId...`n"
            Invoke-WebRequest @ApiArguments | Out-Null
            $Log += "SUCCESS: #$TicketId has been updated."
            Write-Warning "SUCCESS: #$TicketId has been updated."
            $Result = $true
        } catch {
            $Log += "ERROR: Unable to update #$TicketId.`nERROR DETAILS: " + $_
            Write-Error "Unable to update #$TicketId : $_"
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