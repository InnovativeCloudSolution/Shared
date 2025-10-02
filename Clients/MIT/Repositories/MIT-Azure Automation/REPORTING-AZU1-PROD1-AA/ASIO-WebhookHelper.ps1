param (
    [Parameter(Mandatory = $true)]
    [object]$WebhookData
)

# ================================================================================
# DEPENDENCIES AND CONFIGURATION
# ================================================================================

. .\Function-ALL-CWM.ps1
. .\Function-ALL-MSGraph.ps1

#==============================================================================
# WEBHOOK DATA VALIDATION FUNCTIONS
#==============================================================================

function Get-WebhookData {
    param (
        [Parameter(Mandatory = $true)]
        [object]$WebhookData
    )

    if (-not $WebhookData) {
        throw "Invalid Webhook Data: WebhookData parameter is null or empty."
    }
    
    if (-not $WebhookData.PSObject.Properties.Name -contains 'RequestHeader') {
        throw "Invalid Webhook Data: Missing RequestHeader property."
    }
    
    if (-not $WebhookData.PSObject.Properties.Name -contains 'WebhookName') {
        throw "Invalid Webhook Data: Missing WebhookName property."
    }
    
    if (-not $WebhookData.PSObject.Properties.Name -contains 'RequestBody') {
        throw "Invalid Webhook Data: Missing RequestBody property."
    }
    
    try {
        $parsedBody = ConvertFrom-Json $WebhookData.RequestBody
        if (-not $parsedBody) {
            throw "Invalid Webhook Data: RequestBody could not be parsed as JSON."
        }
    }
    catch {
        throw "Invalid Webhook Data: RequestBody is not valid JSON. Error: $_"
    }
    
    return @{
        RequestHeader = $WebhookData.RequestHeader
        WebhookName   = $WebhookData.WebhookName
        RequestBody   = $parsedBody
    }
}

function ConvertTo-Hashtable {
    param (
        [Parameter(Mandatory = $true)]
        [object]$InputObject
    )
    
    if ($null -eq $InputObject) {
        return @{}
    }
    
    if ($InputObject -is [hashtable]) {
        return $InputObject
    }
    
    if ($InputObject -is [System.Collections.IDictionary]) {
        $hashtable = @{}
        foreach ($key in $InputObject.Keys) {
            $hashtable[$key] = $InputObject[$key]
        }
        return $hashtable
    }
    
    if ($InputObject -is [PSCustomObject]) {
        $hashtable = @{}
        $InputObject.PSObject.Properties | ForEach-Object {
            $hashtable[$_.Name] = $_.Value
        }
        return $hashtable
    }
    
    if ($InputObject -is [array]) {
        $hashtable = @{}
        for ($i = 0; $i -lt $InputObject.Length; $i++) {
            $hashtable[$i.ToString()] = $InputObject[$i]
        }
        return $hashtable
    }
    
    return @{ Value = $InputObject }
}

#==============================================================================
# GENERIC PROCESSING FUNCTIONS
#==============================================================================

function Start-Onboarding {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Payload,
        [Parameter(Mandatory = $true)]
        [psobject]$Ticket,
        [Parameter(Mandatory = $true)]
        [psobject]$Company
    )
    
    Write-Output "[INFO]: Processing onboarding request"
    
    $configResult = New-CWM-Configuration -Company $Company -ConfigurationType "Automation - Submission" -Name "$($Ticket.id)"
    if (-not $configResult) {
        Write-Error "[ERROR]: Configuration Item creation failed"
        return
    }
    
    Update-CWM-Configuration -ConfigurationId $configResult.id -FieldPath "status" -FieldValue @{ name = "Active" }
    
    New-CWM-TicketConfiguration -TicketID $Ticket.id -deviceIdentifier $configResult.id
    Write-Output "[INFO]: Configuration Item created and linked to Ticket $($Ticket.id)"
    
    $Payload = ConvertTo-Hashtable -InputObject $Payload
    if ($Payload -is [array]) {
        $Payload = $Payload[0]
    }
    if ($Payload -isnot [hashtable] -and $Payload -isnot [System.Collections.IDictionary]) {
        $Payload = ConvertTo-Hashtable -InputObject $Payload
    }
    
    try {
        if ($Payload -is [array]) { $Payload = $Payload[0] }
        if ($Payload -isnot [hashtable] -and $Payload -isnot [System.Collections.IDictionary]) {
            $convertedPayload = ConvertTo-Hashtable -InputObject $Payload
            if ($convertedPayload -is [array] -and $convertedPayload.Count -gt 0) {
                $Payload = $convertedPayload[0]
            } else {
                $Payload = $convertedPayload
            }
        }
        $Payload = Update-PersonaFields -Payload $Payload -PersonaValues "DEFAULT-ONBOARDING" -Company $Company -ConfigurationId $configResult.id
    }
    catch {
        Write-Output "[WARNING]: Error in Update-PersonaFields for DEFAULT-ONBOARDING: $($_.Exception.Message)"
    }
    
    if (-not [string]::IsNullOrEmpty($Payload.organisation_persona)) {
        if ($Payload -is [array]) {
            $Payload = $Payload[0]
        }
        if ($Payload -isnot [hashtable] -and $Payload -isnot [System.Collections.IDictionary]) {
            $Payload = ConvertTo-Hashtable -InputObject $Payload
        }
        try {
            if ($Payload -is [array]) { $Payload = $Payload[0] }
            if ($Payload -isnot [hashtable] -and $Payload -isnot [System.Collections.IDictionary]) {
                $convertedPayload = ConvertTo-Hashtable -InputObject $Payload
                if ($convertedPayload -is [array] -and $convertedPayload.Count -gt 0) {
                    $Payload = $convertedPayload[0]
                } else {
                    $Payload = $convertedPayload
                }
            }
            $Payload = Update-PersonaFields -Payload $Payload -PersonaValues $Payload.organisation_persona -Company $Company -ConfigurationId $configResult.id
        }
        catch {
            Write-Output "[WARNING]: Error in Update-PersonaFields for organisation_persona: $($_.Exception.Message)"
        }
    }
    
    if (-not [string]::IsNullOrEmpty($Payload.organisation_site)) {
        $Payload = Set-SiteDetails -Payload $Payload -Company $Company
    }
    
    $Payload = Set-Fields -Fields $Payload
    Update-CWM-ConfigurationQuestions -ConfigurationId $configResult.id -PayloadData $Payload
    Write-Output "[INFO]: Onboarding completed - Configuration Item: $($configResult.id)"

    Start-Automation -Ticket $Ticket -RunAutomation "GLOBAL - User Onboarding"
    Write-Output "[INFO]: Automation job started - Ticket: $($Ticket.id)"
}

function Start-Offboarding {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Payload,
        [Parameter(Mandatory = $true)]
        [psobject]$Ticket,
        [Parameter(Mandatory = $true)]
        [psobject]$Company
    )
    
    Write-Output "[INFO]: Processing offboarding request"
    
    $configResult = New-CWM-Configuration -Company $Company -ConfigurationType "Automation - Submission" -Name "$($Ticket.id)"
    if (-not $configResult) {
        Write-Error "[ERROR]: Configuration Item creation failed"
        return
    }
    
    Update-CWM-Configuration -ConfigurationId $configResult.id -FieldPath "status" -FieldValue @{ name = "Active" }
    
    New-CWM-TicketConfiguration -TicketID $Ticket.id -deviceIdentifier $configResult.id
    Write-Output "[INFO]: Configuration Item created and linked to Ticket $($Ticket.id)"
    
    $Payload = ConvertTo-Hashtable -InputObject $Payload
    if ($Payload -is [array]) {
        $Payload = $Payload[0]
    }
    if ($Payload -isnot [hashtable] -and $Payload -isnot [System.Collections.IDictionary]) {
        $Payload = ConvertTo-Hashtable -InputObject $Payload
    }
    
    try {
        if ($Payload -is [array]) { $Payload = $Payload[0] }
        if ($Payload -isnot [hashtable] -and $Payload -isnot [System.Collections.IDictionary]) {
            $convertedPayload = ConvertTo-Hashtable -InputObject $Payload
            if ($convertedPayload -is [array] -and $convertedPayload.Count -gt 0) {
                $Payload = $convertedPayload[0]
            } else {
                $Payload = $convertedPayload
            }
        }
        $Payload = Update-PersonaFields -Payload $Payload -PersonaValues "DEFAULT-OFFBOARDING" -Company $Company -ConfigurationId $configResult.id
    }
    catch {
        Write-Output "[WARNING]: Error in Update-PersonaFields for DEFAULT-OFFBOARDING: $($_.Exception.Message)"
    }
    
    $Payload = Set-Fields -Fields $Payload
    Update-CWM-ConfigurationQuestions -ConfigurationId $configResult.id -PayloadData $Payload
    Write-Output "[INFO]: Offboarding completed - Configuration Item: $($configResult.id)"

    Start-Automation -Ticket $Ticket -RunAutomation "GLOBAL - User Offboarding"
    Write-Output "[INFO]: Automation job started - Ticket: $($Ticket.id)"
}

function Update-Onboarded {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Payload,
        [Parameter(Mandatory = $true)]
        [psobject]$Ticket,
        [Parameter(Mandatory = $true)]
        [psobject]$Company
    )
    
    Write-Output "[INFO]: Processing onboarded request - updating mobile only"

    $configItems = Get-CWM-Configuration -Name "$($Ticket.id)" -TypeName "Automation - Submission" -Company $Company.identifier
    if (-not $configItems -or $configItems.Count -eq 0) {
        Write-Error "[ERROR]: No configuration item found with name $($Ticket.id)"
        return
    }

    $configItem = $configItems[0]
    Write-Output "[INFO]: Found configuration item $($configItem.id) for ticket $($Ticket.id)"

    $mobilePayload = @{
        user_mobile = $Payload.user_mobile
    }
    
    $mobilePayload = Set-Fields -Fields $mobilePayload
    Update-CWM-ConfigurationQuestions -ConfigurationId $configItem.id -PayloadData $mobilePayload
    
    Write-Output "[INFO]: Mobile update completed - Configuration Item: $($configItem.id)"
}

function Start-Generic {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Payload,
        [Parameter(Mandatory = $true)]
        [psobject]$Ticket,
        [Parameter(Mandatory = $true)]
        [psobject]$Company
    )
    
    Write-Output "[INFO]: Processing generic payload as-is"
    
    $configResult = New-CWM-Configuration -Company $Company -ConfigurationType "Automation - Submission" -Name "$($Ticket.id)"
    if (-not $configResult) {
        Write-Error "[ERROR]: Configuration Item creation failed"
        return
    }
    
    Update-CWM-Configuration -ConfigurationId $configResult.id -FieldPath "status" -FieldValue @{ name = "Active" }
    
    $Payload = ConvertTo-Hashtable -InputObject $Payload
    if ($Payload -is [array]) {
        $Payload = $Payload[0]
    }
    if ($Payload -isnot [hashtable] -and $Payload -isnot [System.Collections.IDictionary]) {
        $Payload = ConvertTo-Hashtable -InputObject $Payload
    }
    
    $Payload = Set-Fields -Fields $Payload
    Update-CWM-ConfigurationQuestions -ConfigurationId $configResult.id -PayloadData $Payload
    
    Write-Output "[INFO]: Generic processing completed - Configuration Item: $($configResult.id)"
}

function Convert-Fields {
    param (
        [Parameter(Mandatory = $true)]
        [array]$sections,
        [Parameter(Mandatory = $true)]
        [string]$ticket,
        [Parameter(Mandatory = $true)]
        [string]$formName
    )
    
    if (-not $sections -or $sections.Count -eq 0) {
        throw "Sections parameter cannot be null or empty"
    }
    
    if ([string]::IsNullOrEmpty($formName)) {
        throw "FormName parameter cannot be null or empty"
    }
    
    $result = [ordered]@{
        status_source                = "Desk Director"
        status_request_type          = ""
        status_result                = ""
        status_preapproved           = "No"
        status_approved              = "No"
        status_sspr                  = "No"
        status_user_created          = "No"
        status_license_assigned      = "No"
        status_offboarding_scheduled = "No"
        status_cs_task               = "No"
        status_placeholder03         = ""
        task_operation01             = ""
        task_operation02             = ""
        task_operation03             = ""
        task_operation04             = ""
        task_operation05             = ""
        task_operation06             = ""
        task_operation07             = ""
        task_operation08             = ""
        task_operation09             = ""
        task_operation10             = ""
        task_cs_operation01          = ""
        task_cs_operation02          = ""
        task_cs_operation03          = ""
        task_cs_operation04          = ""
        task_cs_operation05          = ""
        task_cs_operation06          = ""
        task_cs_operation07          = ""
        task_cs_operation08          = ""
        task_cs_operation09          = ""
        task_cs_operation10          = ""
        organisation_environment     = ""
        organisation_persona         = ""
        organisation_site            = ""
        exchange_microsoft_domain    = ""
        exchange_email_domain        = ""
        user_firstname               = ""
        user_lastname                = ""
        user_middlename              = ""
        user_displayname             = ""
        user_upn                     = ""
        user_mailnickname            = ""
        user_username                = ""
        user_fullname                = ""
        user_password                = ""
        user_member_type             = "Member"
        user_password_reset_required = "Yes"
        user_startdate               = ""
        user_enddate                 = ""
        user_company                 = ""
        user_department              = ""
        user_title                   = ""
        user_manager_name            = ""
        user_manager_upn             = ""
        user_primary_smtp            = ""
        user_externalemailaddress    = ""
        user_office                  = ""
        user_streetaddress           = ""
        user_city                    = ""
        user_state                   = ""
        user_postalcode              = ""
        user_country                 = ""
        user_employee_id             = ""
        user_employee_type           = ""
        user_mobile                  = ""
        user_business                = ""
        user_home                    = ""
        user_fax                     = ""
        user_ou                      = ""
        user_home_drive              = ""
        user_home_driveletter        = ""
        user_extensionattribute1     = ""
        user_extensionattribute2     = ""
        user_extensionattribute3     = ""
        user_extensionattribute4     = ""
        user_extensionattribute5     = ""
        user_extensionattribute6     = ""
        user_extensionattribute7     = ""
        user_extensionattribute8     = ""
        user_extensionattribute9     = ""
        user_extensionattribute10    = ""
        user_extensionattribute11    = ""
        user_extensionattribute12    = ""
        user_extensionattribute13    = ""
        user_extensionattribute14    = ""
        user_extensionattribute15    = ""
        group_security01             = ""
        group_security02             = ""
        group_security03             = ""
        group_security04             = ""
        group_security05             = ""
        group_security06             = ""
        group_security07             = ""
        group_security08             = ""
        group_exchange01             = ""
        group_exchange02             = ""
        group_exchange03             = ""
        group_exchange04             = ""
        group_exchange05             = ""
        asset01_required             = "No"
        asset01_source               = ""
        asset01_vendor               = ""
        asset01_tag                  = ""
        asset02_required             = "No"
        asset02_source               = ""
        asset02_vendor               = ""
        asset02_tag                  = ""
        asset03_required             = "No"
        asset03_source               = ""
        asset03_vendor               = ""
        asset03_tag                  = ""
        asset04_required             = "No"
        asset04_source               = ""
        asset04_vendor               = ""
        asset04_tag                  = ""
        asset05_required             = "No"
        asset05_source               = ""
        asset05_vendor               = ""
        asset05_tag                  = ""
    }

    $managerHandled = $false
    $groupFields = @()
    1..8 | ForEach-Object { $groupFields += ('group_security{0:D2}' -f $_) }
    1..5 | ForEach-Object { $groupFields += ('group_exchange{0:D2}' -f $_) }

    foreach ($section in $sections) {
        foreach ($field in $section.fields) {
            $identifier = $field.identifier
            if (-not $identifier -or $identifier.Length -le 7) { continue }

            if (-not $managerHandled -and $identifier -eq "user_manager" -and $field.PSObject.Properties.Name -contains 'choices') {
                foreach ($choice in $field.choices) {
                    if ($choice.selected -eq $true) {
                        $raw = $choice.name -replace '&lt;', '<' -replace '&gt;', '>'
                        if ($raw -match "^(?<name>.+?) <(?<email>[^>]+)>$") {
                            $result.user_manager_name = $Matches["name"]
                            $result.user_manager_upn = $Matches["email"]
                            $managerHandled = $true
                            break
                        }
                    }
                }
                if ($managerHandled) { continue }
            }

            if ($identifier -match '^status_(preapproved|sspr)$') {
                $result[$identifier] = "Yes"
                continue
            }

            if ($formName -match 'Onboarding') {
                $fn = $result.user_firstname
                $ln = $result.user_lastname
                if (-not [string]::IsNullOrEmpty($fn) -and -not [string]::IsNullOrEmpty($ln)) {
                    $result.user_displayname = "$fn $ln"
                    $result.user_fullname = "$fn $ln"
                }
            }

            if ($formName -notmatch 'Onboarding') {
                if ((($identifier -eq "user_offboardeduser") -or ($identifier -eq "user_identifier")) -and $field.PSObject.Properties.Name -contains 'choices') {
                    $first = $field.choices | Select-Object -First 1
                    if ($first -and $first.name -match "^(?<name>.+?) <(?<email>[^>]+)>$") {
                        $result.user_fullname = $Matches["name"]
                        $result.user_upn = $Matches["email"]
                        $result.user_primary_smtp = $Matches["email"]
                        continue
                    }
                }
            }

            if ($field.PSObject.Properties.Name -contains 'choices' -and $groupFields -contains $identifier) {
                $out = @()
                foreach ($choice in $field.choices) {
                    if ($choice.name) {
                        $raw = $choice.name -replace '&amp;', '&' -replace '&lt;', '<' -replace '&gt;', '>'
                        $isManual = $false
                        if ($raw -match '^(?<base>.+?) \[MANUAL\]$') {
                            $isManual = $true
                            $raw = $Matches["base"]
                        }
                        if ($raw -match '^(?<pre>[^\[]*?)\s*\[(?<inner>[^\]]+)\]$') {
                            $pre = ($Matches["pre"]).Trim()
                            $inner = $Matches["inner"]
                            $innerSplit = $inner -split '\|', 2
                            $main = $innerSplit[0]
                            $email = if ($innerSplit.Count -gt 1) { $innerSplit[1] } else { "" }
                            $mainParts = $main -split ':', 2
                            $partA = $mainParts[0]
                            $partB = if ($mainParts.Count -gt 1) { $mainParts[1] } else { "" }
                            $permList = @("Member", "Owner", "Send On Behalf", "Send As", "Full Access", "Read Permission")
                            $perm = if ($permList -contains $partB) { $partB } else { "" }
                            $sourceTag = if ($isManual) { "MANUAL" } else { "" }
                            $out += "$pre`:$partA`:$perm`:$sourceTag|$email"
                        }
                    }
                }
                if ($out.Count -gt 0) { $result[$identifier] = ($out -join ',') }
                continue
            }

            if ($field.PSObject.Properties.Name -contains 'choices') {
                $out = @()
                foreach ($choice in $field.choices) {
                    if ($choice.name) {
                        $decoded = $choice.name -replace '&amp;', '&' -replace '&lt;', '<' -replace '&gt;', '>'
                        if ($decoded -match '<(?<email>[^>]+)>') {
                            $out += $Matches["email"]
                        }
                        else {
                            $out += $decoded
                        }
                    }
                }
                if ($out.Count -gt 0) { $result[$identifier] = ($out -join ',') }
                continue
            }

            if ($field.PSObject.Properties.Name -contains 'value') {
                $cleaned = $field.value
                if ($cleaned -match '\[(?<inner>[^\]]+)\]') { $cleaned = $Matches["inner"] }
                elseif ($cleaned -match '<(?<email>[^>]+)>') { $cleaned = $Matches["email"] }
                if (-not [string]::IsNullOrEmpty($cleaned)) { $result[$identifier] = $cleaned }
                continue
            }

            if ($identifier -in @("password_recipient01", "password_recipient02")) {
                if ($field.value -and $field.value -match '\d+') {
                    $result.task_operation03 = $Matches[0]
                }
                continue
            }

            if (-not [string]::IsNullOrEmpty($field.value)) {
                $result[$identifier] = $field.value
            }
        }
    }

    return $result
}

function Set-SiteDetails {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Payload,
        [Parameter(Mandatory = $true)]
        [psobject]$Company
    )
    
    if (-not $Payload -or $Payload.Count -eq 0) {
        throw "Payload parameter cannot be null or empty"
    }
    
    if (-not $Company) {
        throw "Company parameter cannot be null or empty"
    }
    
    if ([string]::IsNullOrEmpty($Payload.organisation_site)) {
        return $Payload
    }
    
    try {
        $sites = Get-CWM-CompanySites -SiteName $Payload.organisation_site -CompanyId $Company.id
        if (-not $sites -or $sites.Count -eq 0) {
            return $Payload
        }
        $site = $sites[0]
        $friendlyNameField = $site.customFields | Where-Object { $_.caption -eq "Friendly Site Name" } | Select-Object -First 1
        $friendlyName = $null
        if ($friendlyNameField) {
            if ($friendlyNameField.PSObject.Properties['value']) {
                $friendlyName = $friendlyNameField.value
            } elseif ($friendlyNameField.PSObject.Properties['fieldValue']) {
                $friendlyName = $friendlyNameField.fieldValue
            }
        }
        if ($friendlyName) { 
            $Payload["user_office"] = $friendlyName 
        }
        elseif ($site.name) { 
            $Payload["user_office"] = $site.name 
        }
        if ($site.addressLine2) {
            $Payload["user_streetaddress"] = "$($site.addressLine1) $($site.addressLine2)"
        }
        else {
            $Payload["user_streetaddress"] = $site.addressLine1
        }
        if ($site.city) { $Payload["user_city"] = $site.city }
        if ($site.stateReference.name) { $Payload["user_state"] = $site.stateReference.name }
        if ($site.zip) { $Payload["user_postalcode"] = $site.zip }
        if ($site.country) { $Payload["user_country"] = $site.country.name }
    }
    catch {
        Write-Error "[ERROR]: Failed to process site details for site: $($Payload.organisation_site). Error: $_"
    }
    return $Payload
}

function Set-Fields {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Fields
    )
    
    if (-not $Fields) { return @{} }
    if ($Fields.Count -eq 0) { return @{} }
    
    $filteredFields = @{}
    foreach ($key in $Fields.Keys) {
        $value = $Fields[$key]
        if (-not [string]::IsNullOrEmpty($value)) {
            $filteredFields[$key] = $value
        }
    }

    $orderedKeys = $filteredFields.Keys | Sort-Object
    $orderedFields = [ordered]@{}
    foreach ($key in $orderedKeys) { $orderedFields[$key] = $filteredFields[$key] }
    return $orderedFields
}

#==============================================================================
# UTILITY FUNCTIONS
#==============================================================================

function Convert-ArrayToHashtable {
    param (
        [Parameter(Mandatory = $true)]
        [object]$InputObject,
        [Parameter(Mandatory = $false)]
        [string]$Context = "Unknown"
    )
    
    if ($InputObject -is [hashtable]) { return $InputObject }
    if ($InputObject -isnot [array]) { return @{} }
    $convertedResult = @{}
    foreach ($item in $InputObject) {
        if ($item -is [hashtable]) {
            foreach ($key in $item.Keys) { $convertedResult[$key] = $item[$key] }
        }
    }
    return $convertedResult
}

function Start-Automation {
    param (
        [Parameter(Mandatory = $true)]
        [psobject]$TicketID,
        [Parameter(Mandatory = $true)]
        [string]$RunAutomation
    )
    try {
        Write-Output "[INFO]: Starting automation job $($RunAutomation) on Ticket $($Ticket.id)"
        $payload = Set-Fields -Fields @{"Run Automation" = $RunAutomation}
        Update-CWM-TicketCustomFields -TicketID $Ticket.id -PayloadData $payload
        return $Ticket
    }
    catch {
        Write-Output "[ERROR]: Exception occurred while running automation job: $_"
        return $null
    }
}

#==============================================================================
# PERSONA FIELD PROCESSING FUNCTIONS
#==============================================================================

function Format-UserString {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FormatString,
        [Parameter(Mandatory = $false)]
        [string]$FirstName,
        [Parameter(Mandatory = $false)]
        [string]$LastName
    )
    
    if (-not $FirstName -or -not $LastName) { return $FormatString }
    
    $result = $FormatString
    
    $patterns = @(
        @{ Pattern = "Firstname\.Lastname"; Replacement = "$FirstName.$LastName"; CaseSensitive = $false },
        @{ Pattern = "firstname\.lastname"; Replacement = "$FirstName.$LastName"; CaseSensitive = $true },
        @{ Pattern = "Firstname_Lastname"; Replacement = "$FirstName`_$LastName"; CaseSensitive = $false },
        @{ Pattern = "firstname_lastname"; Replacement = "$FirstName`_$LastName"; CaseSensitive = $true },
        @{ Pattern = "Firstname-Lastname"; Replacement = "$FirstName-$LastName"; CaseSensitive = $false },
        @{ Pattern = "firstname-lastname"; Replacement = "$FirstName-$LastName"; CaseSensitive = $true },
        @{ Pattern = "F_Lastname"; Replacement = "$($FirstName.Substring(0, 1).ToUpper())_$LastName"; CaseSensitive = $false },
        @{ Pattern = "F-Lastname"; Replacement = "$($FirstName.Substring(0, 1).ToUpper())-$LastName"; CaseSensitive = $false },
        @{ Pattern = "F\.Lastname"; Replacement = "$($FirstName.Substring(0, 1).ToUpper()).$LastName"; CaseSensitive = $false },
        @{ Pattern = "FLastname"; Replacement = "$($FirstName.Substring(0, 1).ToUpper())$LastName"; CaseSensitive = $false },
        @{ Pattern = "flastname"; Replacement = "$($FirstName.Substring(0, 1).ToLower())$LastName"; CaseSensitive = $true },
        @{ Pattern = "first_name"; Replacement = $FirstName; CaseSensitive = $true },
        @{ Pattern = "last_name"; Replacement = $LastName; CaseSensitive = $true },
        @{ Pattern = "first-name"; Replacement = $FirstName; CaseSensitive = $true },
        @{ Pattern = "last-name"; Replacement = $LastName; CaseSensitive = $true },
        @{ Pattern = "firstname"; Replacement = $FirstName; CaseSensitive = $false },
        @{ Pattern = "lastname"; Replacement = $LastName; CaseSensitive = $false },
        @{ Pattern = "FirstName"; Replacement = $FirstName; CaseSensitive = $true },
        @{ Pattern = "LastName"; Replacement = $LastName; CaseSensitive = $true },
        @{ Pattern = "FIRSTNAME"; Replacement = $FirstName; CaseSensitive = $true },
        @{ Pattern = "LASTNAME"; Replacement = $LastName; CaseSensitive = $true },
        @{ Pattern = "FLast"; Replacement = "$($FirstName.Substring(0, 1).ToUpper())$($LastName.Substring(0, 4))"; CaseSensitive = $false },
        @{ Pattern = "F\.L"; Replacement = "$($FirstName.Substring(0, 1).ToUpper()).$($LastName.Substring(0, 1).ToUpper())"; CaseSensitive = $false },
        @{ Pattern = "f\.l"; Replacement = "$($FirstName.Substring(0, 1).ToLower()).$($LastName.Substring(0, 1).ToLower())"; CaseSensitive = $true },
        @{ Pattern = "F_L"; Replacement = "$($FirstName.Substring(0, 1).ToUpper())`_$($LastName.Substring(0, 1).ToUpper())"; CaseSensitive = $false },
        @{ Pattern = "f_l"; Replacement = "$($FirstName.Substring(0, 1).ToLower())`_$($LastName.Substring(0, 1).ToLower())"; CaseSensitive = $true },
        @{ Pattern = "F-L"; Replacement = "$($FirstName.Substring(0, 1).ToUpper())-$($LastName.Substring(0, 1).ToUpper())"; CaseSensitive = $false },
        @{ Pattern = "f-l"; Replacement = "$($FirstName.Substring(0, 1).ToLower())-$($LastName.Substring(0, 1).ToLower())"; CaseSensitive = $true },
        @{ Pattern = "\bF\b"; Replacement = $FirstName.Substring(0, 1).ToUpper(); CaseSensitive = $false },
        @{ Pattern = "\bL\b"; Replacement = $LastName.Substring(0, 1).ToUpper(); CaseSensitive = $false },
        @{ Pattern = "\bf\b"; Replacement = $FirstName.Substring(0, 1).ToLower(); CaseSensitive = $true },
        @{ Pattern = "\bl\b"; Replacement = $LastName.Substring(0, 1).ToLower(); CaseSensitive = $true }
    )
    
    foreach ($pattern in $patterns) {
        if ($pattern.CaseSensitive) {
            $result = $result -replace $pattern.Pattern, $pattern.Replacement
        } else {
            $result = $result -replace $pattern.Pattern, $pattern.Replacement
        }
    }

    $result = $result -replace "`_", "_"
    
    return $result
}

function Update-PersonaFields {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Payload,
        [Parameter(Mandatory = $true)]
        [string]$PersonaValues,
        [Parameter(Mandatory = $false)]
        [psobject]$Company,
        [Parameter(Mandatory = $false)]
        [int]$ConfigurationId
    )
    
    if ([string]::IsNullOrEmpty($PersonaValues)) { return $Payload }
    if (-not $Payload -or $Payload.Count -eq 0) { throw "Payload parameter cannot be null or empty" }
    
    # Ensure Payload is a proper hashtable
    if ($Payload -is [array]) { $Payload = $Payload[0] }
    if ($Payload -isnot [hashtable] -and $Payload -isnot [System.Collections.IDictionary]) {
        $convertedPayload = ConvertTo-Hashtable -InputObject $Payload
        if ($convertedPayload -is [array] -and $convertedPayload.Count -gt 0) {
            $Payload = $convertedPayload[0]
        } else {
            $Payload = $convertedPayload
        }
    }

    if ($Company -and $Company.types -and $Company.types.Count -gt 0) {
        $organisationEnvironment = ""
        foreach ($type in $Company.types) {
            $companyType = $type.name
            if ($companyType -like "*Hybrid*") {
                $organisationEnvironment = "Hybrid"
                break
            }
            elseif ($companyType -like "*Cloud*") {
                $organisationEnvironment = "Cloud"
                break
            }
        }
        
        if (-not [string]::IsNullOrEmpty($organisationEnvironment)) {
            $Payload["organisation_environment"] = $organisationEnvironment
            Write-Output "[DEBUG]: Mapped Company types to organisation_environment: $organisationEnvironment"
        }
        else {
            Write-Output "[DEBUG]: No matching environment type found in Company.types: $($Company.types.name -join ', ')"
        }
    }
    
    $submissionQuestions = @()
    if ($ConfigurationId) {
        $submissionCI = Get-CWMCompanyConfiguration -id $ConfigurationId
        if ($submissionCI -and $submissionCI.questions) {
            $submissionQuestions = $submissionCI.questions | ForEach-Object { $_.question }
        }
    }
    
    $personas = $PersonaValues -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    foreach ($persona in $personas) {
        $persona = $persona.Trim()
        if ([string]::IsNullOrEmpty($persona)) { continue }
        try {
            if ($Company -and $Company.identifier) {
                Write-Output "[DEBUG]: Using Company identifier: $($Company.identifier) for persona: $persona"
                $configItems = Get-CWM-Configuration -Name $persona -TypeName "Automation - Persona" -Company $Company.identifier
            } else {
                Write-Output "[WARNING]: No Company or Company.identifier provided for persona: $persona - SKIPPING to prevent false positives"
                continue
            }
            if (-not $configItems) { continue }
        }
        catch {
            Write-Error "[ERROR]: Failed to retrieve configuration for persona: $persona. Error: $_"
            continue
        }
        foreach ($configItem in $configItems) {
            foreach ($question in $configItem.questions) {
                $identifier = $question.question
                $answerValue = $question.answer
                if ([string]::IsNullOrEmpty($answerValue)) { continue }
                if ($identifier -eq "mn_format") {
                    $formattedValue = Format-UserString -FormatString $answerValue -FirstName $Payload.user_firstname -LastName $Payload.user_lastname
                    $Payload.user_mailnickname = $formattedValue
                }
                elseif ($identifier -eq "upn_format") {
                    $formattedValue = Format-UserString -FormatString $answerValue -FirstName $Payload.user_firstname -LastName $Payload.user_lastname
                    if (-not [string]::IsNullOrEmpty($Payload.exchange_microsoft_domain)) {
                        $Payload.user_upn = "$formattedValue@$($Payload.exchange_microsoft_domain)"
                    }
                    else {
                        $Payload.user_upn = $formattedValue
                    }
                    if (-not [string]::IsNullOrEmpty($Payload.exchange_email_domain)) {
                        $Payload.user_primary_smtp = "$formattedValue@$($Payload.exchange_email_domain)"
                    }
                    else {
                        $Payload.user_primary_smtp = $Payload.user_upn
                    }
                }
                elseif ($identifier -eq "un_format") {
                    $formattedValue = Format-UserString -FormatString $answerValue -FirstName $Payload.user_firstname -LastName $Payload.user_lastname
                    $Payload.user_username = $formattedValue
                }
                else {
                    if ($submissionQuestions -contains $identifier) {
                        $currentValue = $null
                        if ($Payload.ContainsKey($identifier)) { 
                            $currentValue = $Payload[$identifier] 
                        }
                        
                        if ([string]::IsNullOrWhiteSpace($currentValue) -or $currentValue -eq "No") {
                            $Payload[$identifier] = $answerValue
                        }
                        else {
                            $existingValues = ([string]$currentValue -split '\|') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                            $newValues = ([string]$answerValue -split '\|') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                            $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
                            $allValues = @()
                            foreach ($v in $existingValues) { if ($seen.Add($v)) { $allValues += $v } }
                            foreach ($v in $newValues) { if ($seen.Add($v)) { $allValues += $v } }
                            $Payload[$identifier] = ($allValues -join ',')
                        }
                    }
                }
            }
        }
    }
    return $Payload
}

#==============================================================================
# MAIN PROCESSING FUNCTION
#==============================================================================

function ProcessMainFunction {
    Write-Output "[INFO]: Processing Webhook"

    try {
        $WebhookData = Get-WebhookData -WebhookData $WebhookData
        Write-Output "[INFO]: Webhook data parsed successfully"
    }
    catch {
        Write-Error "[ERROR]: Failed to parse WebhookData. Error: $_"
        exit 1
    }

    Write-Output "[INFO]: The Webhook Name: $($WebhookData.WebhookName)"
    Write-Output "[INFO]: The Webhook Header: $($WebhookData.RequestHeader)"
    Write-Output "[INFO]: Webhook Processed. Starting the Main Process"

    try {
        $DecodedRequestBody = $WebhookData.RequestBody
        Write-Output "[INFO]: Webhook RequestBody successfully extracted"
    }
    catch {
        Write-Error "[ERROR]: Failed to parse WebhookData.RequestBody. Error: $_"
        exit 1
    }

    if (-not $DecodedRequestBody -or $DecodedRequestBody.Count -eq 0) {
        Write-Error "[ERROR]: Webhook RequestBody is empty or null"
        exit 1
    }
    
    # Connect to ConnectWise Manage FIRST
    Write-Output "[INFO]: Getting Secrets from AzureVault."
    $AzKeyVaultName = Get-AutomationVariable -Name 'AzKeyVaultName'

    Write-Output "[INFO]: Connecting to ConnectWise Manage."
    ConnectCWM -AzKeyVaultName $AzKeyVaultName `
        -CWMClientIdName 'MIT-CWAApi-ClientId' `
        -CWMPublicKeyName 'MIT-CWMApi-PubKey' `
        -CWMPrivateKeyName 'MIT-CWMApi-PrivateKey' `
        -CWMCompanyIdName 'MIT-CWMApi-CompanyId' `
        -CWMUrlName 'MIT-CWMApi-Server'
    
    if ($DecodedRequestBody.PSObject.Properties.Name -contains 'status_source' -and 
        $DecodedRequestBody.status_source -eq 'Employment Hero') {
        $ticketId = $DecodedRequestBody.cwpsa_ticket
    } else {
        $ticketId = $DecodedRequestBody[0].ticket.entityId
    }
    
    if ([string]::IsNullOrEmpty($ticketId)) {
        Write-Error "[ERROR]: Missing ticket ID in payload"
        return
    }
    
    $Ticket = Get-CWM-Ticket -TicketID $ticketId
    if (-not $Ticket) {
        Write-Output "[ERROR]: Failed to retrieve Ticket $ticketId"
        return
    }
    
    $Company = Get-CWM-Company -ClientIdentifier $Ticket.company.identifier
    if (-not $Company) {
        Write-Output "[ERROR]: Failed to retrieve Company for Ticket $ticketId"
        return
    }
    
    Write-Output "[DEBUG]: DecodedRequestBody type: $($DecodedRequestBody.GetType().FullName)"
    Write-Output "[DEBUG]: DecodedRequestBody count: $($DecodedRequestBody.Count)"
    Write-Output "[DEBUG]: DecodedRequestBody properties: $($DecodedRequestBody.PSObject.Properties.Name -join ', ')"
    
    if ($DecodedRequestBody.PSObject.Properties.Name -contains 'status_source' -and $DecodedRequestBody.status_source -eq 'Employment Hero') {
        Write-Output "[INFO]: Employment Hero payload detected"
        Write-Output "[DEBUG]: About to convert Employment Hero payload to hashtable"
        Write-Output "[DEBUG]: DecodedRequestBody is null: $(-not $DecodedRequestBody)"
        Write-Output "[DEBUG]: DecodedRequestBody content: $($DecodedRequestBody | ConvertTo-Json -Depth 2)"
        
        if (-not $DecodedRequestBody) {
            Write-Error "[ERROR]: DecodedRequestBody is null for Employment Hero"
            return
        }
        
        try {
            $hashtablePayload = ConvertTo-Hashtable -InputObject $DecodedRequestBody
            Write-Output "[DEBUG]: ConvertTo-Hashtable succeeded"
        }
        catch {
            Write-Error "[ERROR]: ConvertTo-Hashtable failed: $_"
            return
        }
        
        if (-not $hashtablePayload) {
            Write-Error "[ERROR]: Failed to convert Employment Hero payload to hashtable"
            return
        }
        
        switch ($DecodedRequestBody.status_request_type) {
            "user_onboarding" {
                Start-Onboarding -Payload $hashtablePayload -Ticket $Ticket -Company $Company
            }
            "user_offboarding" {
                Start-Offboarding -Payload $hashtablePayload -Ticket $Ticket -Company $Company
            }
            "user_onboarded" {
                Update-Onboarded -Payload $hashtablePayload -Ticket $Ticket -Company $Company
            }
            default {
                Write-Error "[ERROR]: Unknown Employment Hero request type: $($DecodedRequestBody.status_request_type)"
                return
            }
        }
    } else {
        Write-Output "[INFO]: Desk Director webhook detected"
        $formName = $DecodedRequestBody[0].form.name
        $ProcessedPayload = Convert-Fields -sections $DecodedRequestBody[0].form.sections -ticket $ticketId -formName $formName
        
        if (-not $ProcessedPayload) {
            Write-Error "[ERROR]: Failed to convert Desk Director form data"
            return
        }
        
        $hashtablePayload = ConvertTo-Hashtable -InputObject $ProcessedPayload
        if (-not $hashtablePayload) {
            Write-Error "[ERROR]: Failed to convert DD payload to hashtable"
            return
        }
        
        if ($formName -match 'Onboarding') {
            Start-Onboarding -Payload $hashtablePayload -Ticket $Ticket -Company $Company
        }
        elseif ($formName -match 'Offboarding') {
            Start-Offboarding -Payload $hashtablePayload -Ticket $Ticket -Company $Company
        }
        elseif ($formName -match 'Guest User Onboarding') {
            Start-Onboarding -Payload $hashtablePayload -Ticket $Ticket -Company $Company
        }
        else {
            Start-Generic -Payload $hashtablePayload -Ticket $Ticket -Company $Company
        }
    }
}

#==============================================================================
# SCRIPT EXECUTION
#==============================================================================

ProcessMainFunction
