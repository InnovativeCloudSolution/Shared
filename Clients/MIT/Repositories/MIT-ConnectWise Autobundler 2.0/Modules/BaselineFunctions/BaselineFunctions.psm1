Import-Module 'ConnectWiseManageAPI'

$ClientRespondedStatusName = 'Client Responded'

## FUNCTIONS FOR INTERNAL MODULE USE ##
function Remove-TicketField {
    param (
        [Parameter(Mandatory)][int]$TicketId,
        [Parameter(Mandatory)][string]$Path
    )

    <#
    .DESCRIPTION
    Removes a field from a ticket.
    #>
    
    # Build arguments for updated field
    $UpdateTicketParam = @{
        ID = $TicketId
        Operation = 'remove'
        Path = $Path
        Value = ''
    }

    # Update field on ticket
    try {
        Update-CWMTicket @UpdateTicketParam | Out-Null
        Write-Warning "[#$($TicketId)] Remove-TicketField: #$($TicketId)'s $Path has been removed."
    } catch {
        Write-Error "[#$($TicketId)] Remove-TicketField: Unable to remove #$($TicketId)'s $Path : $($_)"
    }
}

function Update-TicketField {
    param(
        [Parameter(Mandatory)][int]$TicketId,
        [Parameter(Mandatory)][string]$Path,
        $Value,
        [int]$ValueId,
        [string]$ValueName,
        $Operation = 'replace'
    )

    <#
    .DESCRIPTION
    Updates a field on a ticket.
    #>

    # Build arguments for updated field
    $UpdateTicketParam = @{
        ID = $TicketId
        Operation = $Operation
        Path = $Path
    }

    # Make sure summary is less than 100 characters
    if ($Path -eq 'summary' -and $Value.length -gt 100) {
        $Value = $Value.Substring(0, 99)
    }

    # Replace value with ID version if provided
    if ($ValueId -ne 0 -and $null -ne $ValueId) {
        $UpdateTicketParam += @{Value = @{id = $ValueId}}
        if ($ValueName -ne '' -and $null -ne $ValueId) {
            $Value = "$ValueName ($ValueId)"
        } else {
            $Value = $ValueId
        }
    } else {
        $UpdateTicketParam += @{Value = $Value}
    }

    # Update field on ticket
    try {
        Update-CWMTicket @UpdateTicketParam | Out-Null
        if ($Path -eq 'customFields') {
            Write-Warning "[#$($TicketId)] Update-TicketField: #$($TicketId)'s custom fields have been updated."
        } else {
            Write-Warning "[#$($TicketId)] Update-TicketField: #$($TicketId)'s $Path has been set to $Value."
        }
    } catch {
        Write-Error "[#$($TicketId)] Update-TicketField: Unable to set #$($TicketId)'s $Path to $Value : $($_)"
    }
}

function Update-TicketBoard {
    param (
        [Parameter(Mandatory)][int]$TicketId,
        [Parameter(Mandatory)][int]$BoardId,
        [string]$TypeName,
        [string]$SubtypeName,
        [string]$ItemName
    )

    <#
    .DESCRIPTION
    Updates a ticket's board.
    #>

    # Fetch ticket details pre-change
    $Ticket = Get-CWMTicket -ticketId $TicketId

    # Fill type, subtype, and item names if not provided
    if (($TypeName -eq '') -and ($null -ne $Ticket.type.name)) { $TypeName = $Ticket.type.name }
    if (($SubtypeName -eq '') -and ($null -ne $Ticket.subType.name)) { $SubtypeName = $Ticket.subType.name }
    if (($ItemName -eq '') -and ($null -ne $Ticket.item.name)) { $ItemName = $Ticket.item.name }

    # Get board to confirm validity
    try {
        $Board = Get-CWMServiceBoard -id $BoardId
        Write-Host "[#$($TicketId)] Update-TicketBoard: Fetched board $($Board.name) ($($Board.id))."
    }
    catch {
        Write-Error "[#$($TicketId)] Update-TicketBoard: Unable to fetch board $BoardId : $($_)"
    }

    if ($null -ne $Board) {
        # Get type based on ticket or specified name
        if ($TypeName -ne '') {
            $Type = Get-CWMBoardType -parentId $BoardId -condition "name = '$TypeName'"
        }

        if ($null -ne $Type.id) {
            # Get subtype alone if item not provided
            if (($ItemName -eq '') -and ($SubtypeName -ne '')) {
                $Subtype = Get-CWMBoardSubtype -parentId $BoardId -condition "name = '$SubtypeName'"
            }
            # Get subtype and item together
            elseif (($SubtypeName -ne '') -and ($ItemName -ne '')) {
                $TypeSubTypeItemAssociation = Get-CWMBoardTypeSubTypeItemAssociation -parentId $BoardId `
                -condition "type/id = $($Type.id) AND subtype/name = '$SubtypeName' AND item/name = '$ItemName'"
                $Subtype = $TypeSubTypeItemAssociation.subType
                $Item = $TypeSubTypeItemAssociation.item
            }
        }
        
        # Update board, type, subtype, and item on ticket
        Update-TicketField -TicketId $TicketId -Path 'board' -ValueId $Board.id
        if ($null -ne $Type) { Update-TicketField -TicketId $TicketId -Path 'type' -ValueId $Type.id }
        if ($null -ne $Subtype) { Update-TicketField -TicketId $TicketId -Path 'subType' -ValueId $Subtype.id }
        if ($null -ne $Item) { Update-TicketField -TicketId $TicketId -Path 'item' -ValueId $Item.id }
    }
}

function Update-TicketLevel {
    param (
        [Parameter(Mandatory)][int]$TicketId,
        [Parameter(Mandatory)][string]$Level
    )

    <#
    .DESCRIPTION
    Updates the Level custom field on a ticket.

    .NOTES
    This function pulls all custom fields, finds the Level custom field, and updates it.
    CWM requires that custom fields are provided as a complete object with every field accounted for.
    #>

    # Create custom fields array
    $CustomFields = @()

    # Create custom object to store custom fields

    # Get ticket details
    $Ticket = Get-CWMTicket -id $TicketId

    # Continue if ticket exists
    if ($null -ne $Ticket) {
        # Create object to push level info
        foreach ($CustomField in $Ticket.customFields) {
            # Build new custom field
            $NewCustomField = [ordered]@{
                id = $CustomField.id
                caption = $CustomField.caption
                type = $CustomField.type
                entryMethod = $CustomField.entryMethod
                numberOfDecimals = $CustomField.numberOfDecimals
            }

            # Add level info if relevant, or existing value if not
            if ($CustomField.caption -eq 'Level') {
                $NewCustomField.Add('value', $Level)
            } elseif ($null -ne $CustomField.value) {
                $NewCustomField.Add('value', $CustomField.value)
            }

            # Add to array
            $CustomFields += $NewCustomField
        }

        # Update all custom fields
        Update-TicketField -TicketId $TicketId -Path 'customFields' -Operation 'replace' -Value $CustomFields
    }
}

function Add-TicketTask {
    param (
        [Parameter(Mandatory)][int]$TicketId,
        [Parameter(Mandatory)][string]$Notes,
        [string]$Resolution
    )

    <#
    .DESCRIPTION
    Adds a task to a given ticket, with an optional Resolution field.
    #>

    # Build task params
    $TaskParams = @{
        parentId = $TicketId
        notes = $Notes
    }

    if ($Resolution -ne '') { $TaskParams += @{ resolution = $Resolution } }
    
    try {
        $Task = New-CWMTicketTask @TaskParams
        Write-Warning "[#$($TicketId)] Add-TicketTask: Added task $($Task.id) to #$($TicketId)."
    }
    catch {
        Write-Error "[#$($TicketId)] Add-TicketTask: Unable to add task $($Task.id) to #$($TicketId) : $($_)"
    }
}

function Add-ChildTickets {
    param (
        [Parameter(Mandatory)][int]$TicketId,
        [Parameter(Mandatory)][array]$ChildTicketIds
    )

    <#
    .DESCRIPTION
    Bundles child ticket/s into a parent ticket.

    .NOTES
    This script uses the ConnectWise Manage API directly, as the PowerShell module does not include a function for ticket bundling.
    #>

    # Initialise final list of child tickets
    $FinalChildTicketIds = @()

    # Check to make sure the target ticket exists
    $ParentTicket = Get-CWMTicket -id $TicketId

    # Check to make sure target ticket is open
    if ($null -ne $ParentTicket) {
        if ($ParentTicket.closedFlag -eq $true) {
            $30DaysAgo = (Get-Date).AddDays(-30)
            if ([DateTime]$ParentTicket[0].closedDate -ge $30DaysAgo) {
                # Reopen closed ticket before bundling
                Write-Host "[#$($TicketId)] Add-ChildTickets: Parent ticket #$($ParentTicket.id) was closed within the last 30 days, and will be reopened if possible."

                # Find Re-Opened status
                $Status = Get-CWMBoardStatus -parentId $ParentTicket.board.id -condition "Name = 'Re-Opened'"

                # Find default status if Re-Opened doesn't exist
                if ($null -eq $Status) {
                    $Status = Get-CWMBoardStatus -parentId $ParentTicket.board.id -condition "defaultFlag = true"
                }

                # Re-open ticket if possible - otherwise, stop bundling
                if ($null -ne $Status) {
                    Update-TicketField -TicketId $TicketId -Path 'status' -ValueId $Status.id
                } else {
                    throw "$($ParentTicket.board.name) does not have an appropriate status. #$($TicketId) will not be re-opened, and child ticket/s will not be bundled."
                }
            }
        }
    } else {
        throw "[#$($TicketId)] Add-ChildTickets: #$TicketId could not be located. Child ticket/s will not be bundled."
    }

    # Scan to make sure ticket isn't being bundled into itself
    foreach ($ChildTicketId in $ChildTicketIds) {
        if ($TicketId -ne $ChildTicketId) {
            $FinalChildTicketIds += $ChildTicketId
        } else {
            Write-Warning "[#$($TicketId)] Add-ChildTickets: A request was made to bundle this ticket into itself. Ignoring request."
        }
    }

    # Run if there are any child tickets left
    if ($FinalChildTicketIds -ne @()) {
        # Build API request parameters
        $ApiBody = @{
            childTicketIds = $FinalChildTicketIds
        }

        $AuthorisationKey = "$($env:APIKEY_CWMANAGE_COMPANYID)+$($env:APIKEY_CWMANAGE_PUBLICKEY):$($env:APIKEY_CWMANAGE_PRIVATEKEY)"
        $EncodedKey = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($AuthorisationKey))
        $ApiAuthorisation = "Basic $($EncodedKey)"

        $ApiParameters = @{
            Uri = "$($env:APIKEY_CWMANAGE_URL)/v4_6_release/apis/3.0/service/tickets/$TicketId/attachChildren"
            Method = 'POST'
            Body = ConvertTo-Json -InputObject $ApiBody -Depth 100
            ContentType = 'application/json'
            Headers = @{
                'clientId' = $env:APIKEY_CWMANAGE_CLIENTID
                'Authorization' = $ApiAuthorisation
            }
            UseBasicParsing = $true
        }

        # Attach child ticket/s to parent ticket
        try {
            Invoke-WebRequest @ApiParameters | Out-Null
            Write-Warning "[#$($TicketId)] Add-ChildTickets: Child ticket/s have been attached to #$($TicketId)."
        }
        catch {
            Write-Error "[#$($TicketId)] Add-ChildTickets: Child tickets have NOT been attached to #$($TicketId) : $($_)"
        }

        # Update details on child tickets
        foreach ($FinalChildTicketId in $FinalChildTicketIds) {
            $Ticket = Get-CWMTicket -ticketId $FinalChildTicketId
            Update-ChildTicket -Ticket $Ticket
        }

        # Mark parent ticket as Customer Updated
        Update-TicketField -TicketId $TicketId -Path 'customerUpdatedFlag' -Value $true -Operation 'replace'

        # Move ticket to Client Responded
        $ClientRespondedStatus = Get-CWMBoardStatus -serviceBoardId $ParentTicket.board.id -condition "name = '$ClientRespondedStatusName'"

        # Update child ticket's status if located
        if ($null -ne $ClientRespondedStatus.id) {
            Write-Host "[#$($TicketId)] Add-ChildTickets: Status '$($ClientRespondedStatus.name)' ($($ClientRespondedStatus.id)) has been fetched."
            Update-TicketField -TicketId $Ticket.id -Path 'status' -ValueId $ClientRespondedStatus.id -ValueName $ClientRespondedStatus.name
        } else {
            Write-Host "[#$($TicketId)] Add-ChildTickets: Board $($ParentTicket.board.name) does not have a status named '$ClientRespondedStatusName'."
        }
    }
}

function Connect-CWManage {
    param (
        [Parameter(Mandatory)][int]$TicketId
    )

    <#
    .DESCRIPTION
    
    Pulls ConnectWise Manage credentials from the Azure Function app properties, then connects to the CWM API.
    #>

    $ConnectionDetails = @{
        Server = $env:APIKEY_CWMANAGE_URL
        Company = $env:APIKEY_CWMANAGE_COMPANYID
        PubKey = $env:APIKEY_CWMANAGE_PUBLICKEY
        PrivateKey = $env:APIKEY_CWMANAGE_PRIVATEKEY
        ClientID = $env:APIKEY_CWMANAGE_CLIENTID
    }

    try {
        Connect-CWM @ConnectionDetails | Out-Null
        Write-Host "[#$($TicketId)] Connect-CWManage: Connection to ConnectWise Manage instance has been established."
    } catch {
        Write-Error "[#$($TicketId)] Connect-CWManage: Unable to connect to ConnectWise Manage instance : $($_)"
    }
}