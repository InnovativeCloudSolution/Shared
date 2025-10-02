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
function Get-CWM-Company {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ClientIdentifier
    )
    $Company = Get-CWMCompany -condition "identifier='$ClientIdentifier'"
    return $Company
}
function Get-CWM-CompanySites {
    param (
        [Parameter(Mandatory = $false)]
        [int]$SiteId,
        [Parameter(Mandatory = $false)]
        [string]$SiteName,
        [Parameter(Mandatory = $true)]
        [int]$CompanyId
    )
    
    try {
        if ($SiteId) {
            $Sites = Get-CWMCompanySite -parentId $CompanyId -id $SiteId
            return $Sites
        }
        elseif ($SiteName) {
            $condition = "name='$SiteName'"
            $Sites = Get-CWMCompanySite -parentId $CompanyId -condition $condition -all
            return $Sites
        }
        else {
            $Sites = Get-CWMCompanySite -parentId $CompanyId
            return $Sites
        }
    }
    catch {
        Write-Error "Error retrieving company sites: $_"
        return $null
    }
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
function Get-CWM-Ticket {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TicketID
    )
    $Ticket = Get-CWMTicket -id $TicketID
    return $Ticket
}
function Get-CWM-Tickets {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Condition
    )
    
    try {
        $tickets = Get-CWMTicket -condition $Condition
        return $tickets
    }
    catch {
        Write-Error -Message "Failed to search tickets with condition '$Condition': $_"
        return $null
    }
}
function Get-CWM-TicketTask {
    param (
        [Parameter(Mandatory)][int]$TicketId
    )
    return Get-CWMTicketTask -parentId $TicketId -all
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
        if ($CWMTicketSummary.Length -gt 99) { $CWMTicketSummary = $CWMTicketSummary.Substring(0, 99) }
        $NewTicketParameters = @{
            summary            = $CWMTicketSummary
            company            = @{id = $CWMCompany.id }
            contact            = @{id = $CWMContact.id }
            board              = @{id = $CWMBoard.id }
            initialDescription = $CWMTicketNoteInitial
        }
        if ($CWMBoardStatus -ne $null) { $NewTicketParameters.Add('status', @{name = $CWMBoardStatus.name }) }
        if ($CWMBoardType -ne $null) {
            $NewTicketParameters.Add('type', @{name = $CWMBoardType.name })
            if ($CWMBoardSubType -ne $null) {
                $NewTicketParameters.Add('subType', @{name = $CWMBoardSubType.name })
                if ($CWMBoardItem -ne $null) { $NewTicketParameters.Add('item', @{name = $CWMBoardItem.item.name }) }
            }
        }
        if ($CWMPriority -ne $null) { $NewTicketParameters.Add('priority', @{name = $CWMPriority.name }) }
        if ($CWMTicketNoteInternal -ne '') { $NewTicketParameters.Add('initialInternalAnalysis', $CWMTicketNoteInternal) }
        if ($CWMTicketNoteResolution -ne '') { $NewTicketParameters.Add('initialResolution', $CWMTicketNoteResolution) }
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
function New-CWM-Configuration {
    param (
        [Parameter(Mandatory = $true)]
        [psobject]$Company,
        [Parameter(Mandatory = $true)]
        [string]$ConfigurationType,
        [Parameter(Mandatory = $false)]
        [string]$ConfigurationStatus,
        [Parameter(Mandatory = $false)]
        [string]$Name,
        [Parameter(Mandatory = $false)]
        [string]$Description,
        [Parameter(Mandatory = $false)]
        [hashtable]$AdditionalFields
    )
    try {
        Write-Output "[INFO]: Creating Configuration Item for Company: $($Company.name)"
        $configParams = @{
            name    = if ($Name) { $Name } else { "New Configuration" }
            type    = @{ name = $ConfigurationType }
            company = @{ id = $Company.id }
        }
        if ($ConfigurationStatus) { $configParams.status = @{ name = $ConfigurationStatus } }
        if ($AdditionalFields) {
            foreach ($field in $AdditionalFields.Keys) { $configParams[$field] = $AdditionalFields[$field] }
        }
        $configItem = New-CWMCompanyConfiguration @configParams
        if ($configItem) {
            Write-Output "[INFO]: Configuration Item created successfully with ID: $($configItem.id)"
            return $configItem
        }
        else {
            Write-Output "[ERROR]: Failed to create Configuration Item"
            return $null
        }
    }
    catch {
        Write-Output "[ERROR]: Exception occurred while creating Configuration Item: $_"
        return $null
    }
}
function Get-CWM-Configuration {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [string]$TypeName,
        [Parameter(Mandatory = $false)]
        [string]$Company
    )
    try {
        $conditions = "name='$Name' AND type/name='$TypeName'"
        
        if ($Company) {
            $conditions += " AND company/identifier='$Company'"
        }
        
        $queryParams = @{
            'conditions' = $conditions
        }
        $configItems = Get-CWMCompanyConfiguration -condition $queryParams.conditions
        return $configItems
    }
    catch {
        Write-Output "[ERROR]: Exception occurred while retrieving Configuration Items: $_"
        return $null
    }
}
function Update-CWM-Configuration {
    param (
        [Parameter(Mandatory = $true)]
        [int]$ConfigurationId,
        [Parameter(Mandatory = $true)]
        [string]$FieldPath,
        [Parameter(Mandatory = $true)]
        [object]$FieldValue,
        [Parameter(Mandatory = $false)]
        [string]$Operation = "replace"
    )
    try {
        Write-Output "[INFO]: Updating Configuration Item $ConfigurationId field '$FieldPath' to: $FieldValue"
        $updatedConfig = Update-CWMCompanyConfiguration -id $ConfigurationId -Operation $Operation -Path "$FieldPath" -Value $FieldValue
        if ($updatedConfig) {
            Write-Output "[INFO]: Successfully updated Configuration Item $ConfigurationId field '$FieldPath' to: $FieldValue"
            return $updatedConfig
        }
        else {
            Write-Output "[ERROR]: Failed to update Configuration Item $ConfigurationId field '$FieldPath'"
            return $null
        }
    }
    catch {
        Write-Output "[ERROR]: Exception occurred while updating Configuration Item field: $_"
        return $null
    }
}
function Update-CWM-ConfigurationQuestions {
    param (
        [Parameter(Mandatory = $true)]
        [int]$ConfigurationId,
        [Parameter(Mandatory = $true)]
        [hashtable]$PayloadData
    )
    try {
        Write-Output "[INFO]: Updating Configuration Item $ConfigurationId with Questions from payload"
        $configItem = Get-CWMCompanyConfiguration -id $ConfigurationId
        if (-not $configItem) {
            Write-Output "[ERROR]: Failed to retrieve Configuration Item $ConfigurationId"
            return $null
        }
        $questionsToUpdate = @()
        foreach ($key in $PayloadData.Keys) {
            $value = $PayloadData[$key]
            foreach ($question in $configItem.questions) {
                if ($question.question -eq $key) {
                    $questionsToUpdate += @{questionid = $question.questionid; answer = $value }
                    break
                }
            }
        }
        if ($questionsToUpdate.Count -gt 0) {
            $configItem = Update-CWMCompanyConfiguration -id $ConfigurationId -Operation "replace" -Path "questions" -Value $questionsToUpdate
        }
        return $configItem
    }
    catch {
        Write-Output "[ERROR]: Exception occurred while updating Configuration Item Questions: $_"
        return $null
    }
}
function New-CWM-TicketConfiguration {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TicketID,
        [Parameter(Mandatory = $true)]
        [string]$deviceIdentifier
    )
    try {
        Write-Output "[INFO]: Linking Configuration Item $deviceIdentifier to Ticket $TicketID"
        $ticketConfig = New-CWMTicketConfiguration -parentId $TicketID -id $deviceIdentifier
        if ($ticketConfig) {
            Write-Output "[INFO]: Successfully linked Configuration Item $deviceIdentifier to Ticket $TicketID"
            return $ticketConfig
        }
        else {
            Write-Output "[ERROR]: Failed to link Configuration Item $deviceIdentifier to Ticket $TicketID"
            return $null
        }
    }
    catch {
        Write-Output "[ERROR]: Exception occurred while linking Configuration Item to Ticket: $_"
        return $null
    }
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
function Update-CWM-TicketCustomFields {
    param (
        [Parameter(Mandatory = $true)]
        [int]$TicketID,
        [Parameter(Mandatory = $true)]
        [hashtable]$PayloadData
    )
    try {
        Write-Output "[INFO]: Updating Ticket $TicketID with Custom Fields from payload"
        $Ticket = Get-CWM-Ticket -TicketID $TicketID
        if (-not $Ticket) {
            Write-Output "[ERROR]: Failed to retrieve Ticket $TicketID"
            return $null
        }
        foreach ($key in $PayloadData.Keys) {
            $value = $PayloadData[$key]
            foreach ($customField in $Ticket.customFields) {
                if ($customField.caption -eq $key) {
                    $payload = @{id = $customField.id; caption = $key; value = $value }
                    $Ticket = Update-CWMTicket -id $TicketID -Operation "replace" -Path "customFields" -Value @($payload)
                }
            }
        }
        return $Ticket
    }
    catch {
        Write-Output "[ERROR]: Exception occurred while updating Ticket Custom Fields: $_"
        return $null
    }  
}