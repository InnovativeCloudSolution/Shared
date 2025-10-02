function ConnectCWM {
    param (
        [string]$AzKeyVaultName,
        [string]$CWMClientIdName,
        [string]$CWMPublicKeyName,
        [string]$CWMPrivateKeyName,
        [string]$CWMCompanyIdName,
        [string]$CWMUrlName
    )

    $Secrets = Get-MSGraph-Secrets -AzKeyVaultName $AzKeyVaultName -SecretNames @($CWMClientIdName, $CWMPublicKeyName, $CWMPrivateKeyName, $CWMCompanyIdName, $CWMUrlName)

    $CWMClientId = $Secrets[$CWMClientIdName]
    $CWMPublicKey = $Secrets[$CWMPublicKeyName]
    $CWMPrivateKey = $Secrets[$CWMPrivateKeyName]
    $CWMCompanyId = $Secrets[$CWMCompanyIdName]
    $CWMUrl = $Secrets[$CWMUrlName]

    Connect-CWM -Server $CWMUrl -Company $CWMCompanyId -pubkey $CWMPublicKey -privatekey $CWMPrivateKey -clientId $CWMClientId
}

function ConnectCWMTest {
    param ()

    $CWMUrl = 'https://api-aus.myconnectwise.net/v4_6_release/apis/3.0'
    $CWMCompanyId = 'mit'
    $CWMPublicKey = "vY1EdpIW4IBCg94L"
    $CWMPrivateKey = "mzAzc4CgopmhfLP7"
    $CWMClientId = "1208536d-40b8-4fc0-8bf3-b4955dd9d3b7"

    Connect-CWM -Server $CWMUrl -Company $CWMCompanyId -pubkey $CWMPublicKey -privatekey $CWMPrivateKey -clientId $CWMClientId
}

function Get-CWM-Company {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ClientIdentifier
    )

    $Company = Get-CWMCompany -condition "identifier='$ClientIdentifier'"
    return $Company
}

function Get-CWM-Contact {
    param (
        [string]$ClientIdentifier,
        [string]$ContactName
    )

    $Company = Get-CWM-Company -ClientIdentifier $ClientIdentifier
    $Names = $ContactName -split ' '
    if ($Names.Count -eq 1) {
        $FirstName = '"' + $Names[0] + '"'
        $ContactCondition = "company/identifier='$($Company.identifier)' AND firstName=$FirstName"
        $Contact = Get-CWMCompanyContact -Condition $ContactCondition
        return $Contact
    }
    if ($Names.Count -gt 2) {
        $FirstName = '"' + $Names[0] + '"'
        $LastName = '"' + ($Names[1..($Names.Count - 1)] -join ' ') + '"'
    }
    else {
        $FirstName = '"' + $Names[0] + '"'
        $LastName = '"' + $Names[1] + '"'
    }
    $ContactCondition = "company/identifier='$($Company.identifier)' AND firstName=$FirstName AND lastName=$LastName"
    $Contact = Get-CWMCompanyContact -Condition $ContactCondition
    if ($null -eq $Contact -and $Names.Count -gt 2) {
        $FirstName = '"' + ($Names[0..1] -join ' ') + '"'
        $LastName = '"' + ($Names[2..($Names.Count - 1)] -join ' ') + '"'
        $ContactCondition = "company/identifier='$($Company.identifier)' AND firstName=$FirstName AND lastName=$LastName"
        $Contact = Get-CWMCompanyContact -Condition $ContactCondition
    }
    return $Contact
}

function Get-CWM-Board {
    param (
        [string]$ServiceBoardName
    )

    $ServiceBoardCondition = "name='$ServiceBoardName'"
    $ServiceBoard = Get-CWMServiceBoard -Condition $ServiceBoardCondition

    return $ServiceBoard
}

function Get-CWM-BoardStatus {
    param (
        [string]$ServiceBoardName,
        [string]$ServiceBoardStatusName
    )

    $ServiceBoard = Get-CWM-Board -ServiceBoardName $ServiceBoardName
    $ServiceBoardStatusCondition = "name='$ServiceBoardStatusName'"
    $BoardStatus = Get-CWMBoardStatus -parentId $ServiceBoard.id -condition $ServiceBoardStatusCondition

    return $BoardStatus
}

function Get-CWM-BoardType {
    param (
        [string]$ServiceBoardName,
        [string]$ServiceBoardTypeName
    )

    $ServiceBoard = Get-CWM-Board -ServiceBoardName $ServiceBoardName
    $ServiceBoardTypeCondition = "name='$ServiceBoardTypeName'"
    $BoardType = Get-CWMBoardType -parentId $ServiceBoard.id -condition $ServiceBoardTypeCondition
    return $BoardType
}

function Get-CWM-BoardSubType {
    param (
        [string]$ServiceBoardName,
        [string]$ServiceBoardSubTypeName
    )

    $ServiceBoard = Get-CWM-Board -ServiceBoardName $ServiceBoardName
    $ServiceBoardSubTypeCondition = "name='$ServiceBoardSubTypeName'"
    $BoardSubType = Get-CWMBoardSubType -parentId $ServiceBoard.id -condition $ServiceBoardSubTypeCondition
    return $BoardSubType
}

function Get-CWM-BoardItem {
    param (
        [string]$ServiceBoardName,
        [string]$ServiceBoardTypeName,
        [string]$ServiceBoardSubTypeName,
        [string]$ServiceBoardItemName
    )

    $ServiceBoard = Get-CWM-Board -ServiceBoard $ServiceBoardName
    $ServiceBoardItemCondition = "type/name='$ServiceBoardTypeName' AND subType/name='$ServiceBoardSubTypeName' AND item/name='$ServiceBoardItemName'"
    $BoardItem = Get-CWMBoardTypeSubTypeItemAssociation -parentId $ServiceBoard.id -condition $ServiceBoardItemCondition

    return $BoardItem
}

function Get-CWM-Priority {
    param (
        [string]$PriorityName
    )

    $PriorityCondition = "name='$PriorityName'"
    $Priority = Get-CWMPriority -condition $PriorityCondition

    return $Priority
}

function Get-CWM-CompanySites {
    param (
        [Parameter(Mandatory = $true)]
        [int]$ParentId
    )

    $Sites = Get-CWMCompanySite -parentId $ParentId
    return $Sites
}

function Get-CWM-TimeEntries {
    param (
        [string]$CompanyIdentifier,
        [string]$WorkTypeName,
        [string]$TimeRange
    )

    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("clientId", "$CWMClientId")
    $headers.Add("Authorization", "$CWMAuthentication")

    $currentDate = Get-Date
    switch ($TimeRange) {
        'day' { $StartDate = $currentDate.AddDays(-1) }
        'week' { $StartDate = $currentDate.AddDays(-7) }
        'month' { $StartDate = $currentDate.AddMonths(-1) }
        'year' { $StartDate = $currentDate.AddYears(-1) }
        default { throw "Invalid TimeRange specified. Use 'day', 'week', 'month', or 'year'." }
    }

    $formattedDate = $StartDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
    $TimeEntriesCondition = "company/identifier='$CompanyIdentifier' AND workType/name='$WorkTypeName' and lastUpdated >= [$formattedDate]"

    $TimeEntries = Get-CWMTimeEntry -condition $TimeEntriesCondition -all
    return $TimeEntries
}

function Get-CWM-Ticket {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TicketID
    )

    $Ticket = Get-CWMTicket -id $TicketID
    return $Ticket
}

function New-CWM-Ticket {
    param (
        [Parameter(Mandatory = $true)]
        [psobject]$CWMCompany,
        [Parameter(Mandatory = $true)]
        [psobject]$CWMContact,
        [Parameter(Mandatory = $true)]
        [psobject]$CWMBoard,
        [Parameter(Mandatory = $true)]
        [string]$CWMTicketSummary,
        [Parameter(Mandatory = $true)]
        [string]$CWMTicketNoteInitial,
        [psobject]$CWMBoardStatus = $null,
        [psobject]$CWMBoardType = $null,
        [psobject]$CWMBoardSubType = $null,
        [psobject]$CWMBoardItem = $null,
        [psobject]$CWMPriority = $null,
        [string]$CWMTicketNoteInternal = '',
        [string]$CWMTicketNoteResolution = ''
    )

    try {
        if ($CWMTicketSummary.Length -gt 99) {
            $CWMTicketSummary = $CWMTicketSummary.Substring(0, 99)
        }

        $NewTicketParameters = @{
            summary            = $CWMTicketSummary
            company            = @{id = $CWMCompany.id }
            contact            = @{id = $CWMContact.id }
            board              = @{id = $CWMBoard.id }
            initialDescription = $CWMTicketNoteInitial
        }

        if ($CWMBoardStatus -ne $null) { 
            $NewTicketParameters.Add('status', @{name = $CWMBoardStatus.name }) 
        }

        if ($CWMBoardType -ne $null) { 
            $NewTicketParameters.Add('type', @{name = $CWMBoardType.name })
            if ($CWMBoardSubType -ne $null) {
                $NewTicketParameters.Add('subType', @{name = $CWMBoardSubType.name })
                if ($CWMBoardItem -ne $null) { 
                    $NewTicketParameters.Add('item', @{name = $CWMBoardItem.item.name }) 
                }
            }
        }

        if ($CWMPriority -ne $null) { 
            $NewTicketParameters.Add('priority', @{name = $CWMPriority.name }) 
        }

        if ($CWMTicketNoteInternal -ne '') { 
            $NewTicketParameters.Add('initialInternalAnalysis', $CWMTicketNoteInternal) 
        }
        if ($CWMTicketNoteResolution -ne '') { 
            $NewTicketParameters.Add('initialResolution', $CWMTicketNoteResolution) 
        }

        $NewTicket = New-CWMTicket @NewTicketParameters

        return $NewTicket
    }
    catch {
        Write-Error -Message "Failed to create new ticket: $_"
        return $null
    }
}

function New-CWM-TicketNote {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TicketID,
        [string]$TicketNote,
        [bool]$DescriptionFlag = $false,
        [bool]$InternalFlag = $false,
        [bool]$ResolutionFlag = $false
    )

    if (-not $DetailDescriptionFlag -and -not $InternalFlag -and -not $ResolutionFlag) {
        throw "At least one of 'DetailDescriptionFlag', 'InternalFlag', or 'ResolutionFlag' must be set to 'true'."
    }

    $Ticket = New-CWMTicketNote -ticketId $TicketID -text $TicketNote -detailDescriptionFlag $DescriptionFlag -internalAnalysisFlag $InternalFlag -resolutionFlag $ResolutionFlag

    return $Ticket
}

function Get-CWM-TicketTask {
    param (
        [Parameter(Mandatory)][int]$TicketId
    )
    
    return Get-CWMTicketTask -parentId $TicketId -all
}

function New-CWM-TicketTask {
    param (
        [Parameter(Mandatory)][int]$TicketId,
        [Parameter(Mandatory)][string]$Notes,
        [string]$Resolution
    )

    $TaskParams = @{
        parentId = $TicketId
        notes    = $Notes
    }

    if ($Resolution -ne '') { $TaskParams += @{ resolution = $Resolution } }
    
    return New-CWMTicketTask @TaskParams    
}