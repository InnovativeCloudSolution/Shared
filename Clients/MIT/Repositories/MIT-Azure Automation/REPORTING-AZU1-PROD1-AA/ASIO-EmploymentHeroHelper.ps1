param (
    [Parameter(Mandatory = $false)]
    [object]$WebhookData
)

# ================================================================================
# DEPENDENCIES AND CONFIGURATION
# ================================================================================

. .\Function-ALL-EH.ps1
. .\Function-ALL-CWM.ps1
. .\Function-ALL-MSGraph.ps1

$PersonaLocations = @(
    "BRIBIE ISLAND",
    "CALOUNDRA",
    "EASTERN HEIGHTS",
    "KALLANGUR",
    "MANGO HILL",
    "MANGO HILL CARE SUITES",
    "REDBANK PLAINS",
    "SINNAMON PARK",
    "SUPPORT OFFICE",
    "WATERFORD WEST"
)

$PersonaRoles = @(
    "ADMINISTRATION",
    "BUS DRIVER",
    "CARE COORDINATOR",
    "CARE MANAGER",
    "CARE SERVICES SCHEDULER",
    "CHEF",
    "CLEANER (CARE)",
    "CLEANER (GSC)",
    "COMMUNITY MANAGEMENT",
    "COOK",
    "ENROLLED NURSE",
    "FOOD SERVICES ASSISTANT",
    "HOSPITALITY",
    "LIFESTYLE",
    "MAINTENANCE",
    "PERSONAL CARE WORKER",
    "REGISTERED NURSE",
    "SUPPORT OFFICE"
)

# ================================================================================
# CORE UTILITY FUNCTIONS
# ================================================================================

function Get-WebhookData {
    param (
        [Parameter(Mandatory = $true)]
        [object]$WebhookData
    )

    if (-not $WebhookData) {
        throw "Invalid Webhook Data."
    }
    
    Write-Message "Raw WebhookData type: $($WebhookData.GetType().FullName)"
    Write-Message "Raw RequestBody: $($WebhookData.RequestBody)"
    Write-Message "RequestBody type: $($WebhookData.RequestBody.GetType().FullName)"
    
    $ParsedRequestBody = $null
    if ($WebhookData.RequestBody -is [string]) {
        try {
            $ParsedRequestBody = $WebhookData.RequestBody | ConvertFrom-Json
            Write-Message "Successfully parsed RequestBody from JSON string"
        }
        catch {
            Write-Message "Failed to parse RequestBody as JSON: $_" "ERROR"
            $ParsedRequestBody = $WebhookData.RequestBody
        }
    }
    else {
        $ParsedRequestBody = $WebhookData.RequestBody
        Write-Message "RequestBody is already an object, using directly"
    }
    
    return @{
        RequestHeader = $WebhookData.RequestHeader
        WebhookName   = $WebhookData.WebhookName
        RequestBody   = $ParsedRequestBody
        EmployeeData  = $ParsedRequestBody.data
        Event         = $ParsedRequestBody.event
    }
}

function Write-Message {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "$timestamp [$Level] $Message"
}

function Write-ScriptError {
    param (
        [string]$Message
    )
    Write-Message $Message "ERROR"
    throw $Message
}

function Get-Secrets {
    param (
        [string]$AzKeyVaultName
    )

    try {
        Connect-AzAccount -Identity | Out-Null
        return @{
            CWManageUrl      = (Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWMApi-Server' -AsPlainText) + "v4_6_release/apis/3.0"
            CWMCompanyId     = (Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWMApi-CompanyId' -AsPlainText)
            CWMPublicKey     = (Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWMApi-PubKey' -AsPlainText)
            CWMPrivateKey    = (Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWMApi-PrivateKey' -AsPlainText)
            CWMClientId      = (Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWAApi-ClientId' -AsPlainText)
            EHclient_Id      = (Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'SEA-EH-ClientID' -AsPlainText)
            EHclient_secret  = (Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'SEA-EH-ClientSecret' -AsPlainText)
            EHcode           = (Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'SEA-EH-Code' -AsPlainText)
            EHrefresh_token  = (Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'SEA-EH-RefreshToken' -AsPlainText)
            EHOrganizationId = (Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'SEA-EH-OrganizationId' -AsPlainText)
            EHRedirectUri    = (Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'SEA-EH-RedirectUri' -AsPlainText)
        }
    }
    catch {
        Write-Error -Message $_.Exception.Message
        throw $_.Exception
    }
}

function Get-PreCheck {
    param (
        [Parameter(Mandatory = $true)]
        [object]$ParsedWebhookData
    )

    try {
        if ($null -eq $ParsedWebhookData) {
            Write-Message "ParsedWebhookData is null" "ERROR"
            return $null
        }
        
        if ($null -eq $ParsedWebhookData.EmployeeData) {
            Write-Message "No employee data found in webhook data" "ERROR"
            return $null
        }
        
        $EmployeeData = $ParsedWebhookData.EmployeeData
        $EmployeeId = $EmployeeData.id
        
        if ([string]::IsNullOrEmpty($EmployeeId)) {
            Write-Message "Employee ID is null or empty in webhook data" "ERROR"
            return $null
        }
        
        $EventType = $ParsedWebhookData.Event
        
        if ([string]::IsNullOrEmpty($EventType)) {
            Write-Message "Event is null or empty in webhook data" "ERROR"
            return $null
        }

        return @{
            Event        = $EventType
            EmployeeData = $EmployeeData
            EmployeeId   = $EmployeeId
        }
    }
    catch {
        Write-Message "Error in Get-PreCheck: $_" "ERROR"
        return $null
    }
}

# ================================================================================
# DATE AND VALIDATION FUNCTIONS
# ================================================================================

function Test-DateValid {
    param (
        [string]$dateString,
        [ref]$parsedDate
    )
    $formats = @("dd/MM/yyyy hh:mm tt", "dd/MM/yyyy hh:mmtt", "yyyy-MM-ddTHH:mm:sszzz")
    foreach ($format in $formats) {
        try {
            $parsedDate.Value = [DateTime]::ParseExact($dateString, $format, $null)
            return $true
        }
        catch {
            continue
        }
    }
    return $false
}

function Set-EmployeeDate {
    param (
        [string]$inputDate,
        [string]$type
    )
    $dateTime = [DateTime]::Parse($inputDate)
    if ($type -eq "start") {
        return $dateTime.ToString("ddd dd MMM, yyyy")
    }
    elseif ($type -eq "end") {
        return $dateTime.ToString("ddd dd MMM, yyyy 'at' h:mm tt")
    }
}

function Set-EmployeeDateForPayload {
    param (
        [string]$inputDate
    )
    $dateTime = [DateTime]::Parse($inputDate)
    return $dateTime.ToString("yyyy-MM-ddTHH:mm")
}

# ================================================================================
# EMPLOYEE DATA PROCESSING FUNCTIONS
# ================================================================================

function Set-Persona {
    param (
        $EmployeeTeamNames,
        $PersonaLocations,
        $PersonaRoles
    ) 

    $MatchedPersonaRoles = @()
    $MatchedPersonaLocations = @()

    foreach ($EmployeeTeamName in $EmployeeTeamNames) {
        if ($PersonaRoles -contains $EmployeeTeamName) {
            $MatchedPersonaRoles += $EmployeeTeamName
        }

        if ($PersonaLocations -contains $EmployeeTeamName) {
            $MatchedPersonaLocations += $EmployeeTeamName
        }
    }

    if ($MatchedPersonaRoles.Count -eq 0 -or $MatchedPersonaLocations.Count -eq 0) {
        Write-Message "No valid roles or locations found in team names" "WARN"
        return $null
    }

    $PersonaCombinations = @()
    foreach ($Role in $MatchedPersonaRoles) {
        foreach ($Location in $MatchedPersonaLocations) {
            $PersonaCombinations += "$Role-$Location"
        }
    }

    $AllPersonas = $PersonaCombinations -join ', '
    
    return $AllPersonas
}

function Set-Department {
    param (
        $Persona,
        $PersonaRoles
    )
    $Role = $null
    $Department = "Seasons"

    $IndividualPersonas = $Persona -split ','
    
    foreach ($IndividualPersona in $IndividualPersonas) {
        $IndividualPersona = $IndividualPersona.Trim() # Remove any whitespace
        
        foreach ($PersonaRole in $PersonaRoles) {
            if ($IndividualPersona.IndexOf($PersonaRole, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $Role = $PersonaRole
                if ($Role -eq "BUS DRIVER" -or $Role -eq "CHEF" -or $Role -eq "COOK" -or $Role -eq "CLEANER (CARE)" -or $Role -eq "CLEANER (GSC)" -or $Role -eq "FOOD SERVICES ASSISTANT" -or $Role -eq "HOSPITALITY" -or $Role -eq "MAINTENANCE" -or $Role -eq "PERSONAL CARE WORKER") {
                    $Department = "F3 Seasons"
                }
                break
            }
        }

        if ($Role) {
            break
        }
    }
    
    return $Department
}

function Set-Role {
    param (
        $Persona,
        $PersonaRoles
    )
    $Role = $null

    $IndividualPersonas = $Persona -split ','
    
    foreach ($IndividualPersona in $IndividualPersonas) {
        $IndividualPersona = $IndividualPersona.Trim()
        
        foreach ($PersonaRole in $PersonaRoles) {
            if ($IndividualPersona.IndexOf($PersonaRole, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $Role = $PersonaRole
                break
            }
        }

        if ($Role) {
            break
        }
    }
    
    return $Role
}

function Get-Location {
    param (
        $EmployeeLocation,
        $CWMLocations
    )
    $MatchedLocation = ""
    foreach ($EmployeeTeamName in $EmployeeTeamNames) {
        foreach ($CWMLocation in $CWMLocations) {
            if ($CWMLocation.IndexOf($EmployeeTeamName, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $MatchedLocation = $CWMLocation
            }
        }
    }
    return $MatchedLocation
}

# ================================================================================
# DEVICE REQUIREMENT FUNCTIONS
# ================================================================================

function Set-ComputerRequirement {
    param (
        $Role,
        $PersonaRoles
    )
    $ComputerRequirement = "No"
    $ComputerIdentifier = $null
    
    if ($PersonaRoles -contains $Role) {
        if ($Role -in @("BUS DRIVER", "CLEANER (CARE)", "CLEANER (GSC)", "HOSPITALITY", "PERSONAL CARE WORKER", "COOK", "ENROLLED NURSE", "REGISTERED NURSE")) {
            $ComputerRequirement = "No"
        }
        else {
            $ComputerRequirement = "Yes"
            if ($Role -in @("CHEF", "LIFESTYLE", "MAINTENANCE")) {
                $ComputerIdentifier = "Check with Britt if the user will share a laptop, use a hot desk, or get a new laptop"
            }
            else {
                $ComputerIdentifier = "Check with Britt if the user will get a new computer or reuse an existing one"
            }
        }
    }
    return @($ComputerRequirement, $ComputerIdentifier)
}

function Set-MobileRequirement {
    param (
        $Role,
        $PersonaRoles
    )
    $MobileRequirement = "No"
    $MobileIdentifier = $null
    
    if ($PersonaRoles -contains $Role) {
        if ($Role -in @("ADMINISTRATION", "BUS DRIVER", "CLEANER (CARE)", "CLEANER (GSC)", "PERSONAL CARE WORKER", "REGISTERED NURSE", "ENROLLED NURSE", "HOSPITALITY", "COOK")) {
            $MobileRequirement = "No"
        }
        else {
            $MobileRequirement = "Yes"
            $MobileIdentifier = "Check with Britt if the user will get a new mobile or reuse an existing one"
        }
    }
    return @($MobileRequirement, $MobileIdentifier)
}

function Set-AdditionalNotes {
    param (
        $Role
    )
    $AdditionalNotes = $null
    
    if ($Role -eq "SUPPORT OFFICE") {
        $AdditionalNotes = "Part of SUPPORT OFFICE, check with Brittany Smith what AAD Groups are needed`n"
        $AdditionalNotes += "Part of SUPPORT OFFICE, check with Brittany Smith what shared mailboxes are needed"
    }
    return $AdditionalNotes
}

# ================================================================================
# TICKET CREATION FUNCTIONS
# ================================================================================

function Set-TicketSummaryNewUser {
    param (
        [string]$FirstName,
        [string]$LastName,
        [string]$StartDate
    )

    try {
        $SummaryData = "User Onboarding: $FirstName $LastName, starting on $StartDate"
        return $SummaryData
    }
    catch {
        Write-Error -Message "Failed to set SEA New User ticket summary: $_"
        return $null
    }
}

function Set-TicketSummaryExitUser {
    param (
        [string]$FirstName,
        [string]$LastName,
        [string]$EndDate
    )

    try {
        $SummaryData = "User Offboarding: $FirstName $LastName, departing on $EndDate"
        return $SummaryData
    }
    catch {
        Write-Error -Message "Failed to set SEA Exit User ticket summary: $_"
        return $null
    }
}

function Set-TicketNoteInitialNewUser {
    param (
        [string]$FirstName,
        [string]$LastName,
        [string]$StartDate,
        [string]$MobilePersonal,
        [string]$Company,
        [string]$JobTitle,
        [string]$ManagerTicket,
        [string]$Department,
        [string]$Location,
        [string]$Persona,
        [string]$ComputerRequirement,
        [string]$ComputerIdentifier,
        [string]$MobileRequirement,
        [string]$MobileIdentifier,
        [string]$AdditionalNotes
    )

    try {
        $NoteData = "### Pre-Approved Submission`n"
        $NoteData += "This submission is pre-approved. This submission will be completed without prompting for approval.`n`n"
        $NoteData += "### User Onboarding`n"
        $NoteData += "The new user will be set up using the information exactly as it is provided here. If you are unsure about any information provided here, please call us on 07 3151 9000 to discuss.`n`n"
        $NoteData += "### Response Time`n"
        $NoteData += "It may take us up to four business hours to respond to your request, and it can take a number of days for some aspects of a new user request to be fulfilled, the lead time for new user requests is 5 business days, please provide as much notice as possible so we can action your request and meet your requirements.`n`n"
        $NoteData += "If your new user is starting within the next two business days, please call us on 07 3151 9000 to draw our attention to the new user's start date.`n`n"
        $NoteData += "### Given Name(s)`n"
        $NoteData += "$FirstName`n`n"
        $NoteData += "### Last Name`n"
        $NoteData += "$LastName`n`n"
        $NoteData += "### Start Date`n"
        $NoteData += "$StartDate`n`n"

        if (-not [string]::IsNullOrEmpty($MobilePersonal)) {
            $NoteData += "### User's Mobile Phone Number`n"
            $NoteData += "$MobilePersonal`n`n"
        }

        if (-not [string]::IsNullOrEmpty($Company)) {
            $NoteData += "### Company`n"
            $NoteData += "$Company`n`n"
        }

        $NoteData += "### Position Title`n"
        $NoteData += "$JobTitle`n`n"
        $NoteData += "### Department`n"
        $NoteData += "$Department`n`n"

        if (-not [string]::IsNullOrEmpty($ManagerTicket)) {
            $NoteData += "### Manager`n"
            $NoteData += "$ManagerTicket`n`n"
        }

        $NoteData += "### Persona`n"
        $NoteData += "$Persona`n`n"
        $NoteData += "### Location`n"
        $NoteData += "$Location`n`n"
        $NoteData += "### Computer - Requirement`n"
        $NoteData += "$ComputerRequirement`n`n"

        if ($ComputerRequirement -eq "Yes") {
            $NoteData += "### Computer - Identifier`n"
            $NoteData += "$ComputerIdentifier`n`n"
        }

        $NoteData += "### Mobile - Requirement`n"
        $NoteData += "$MobileRequirement`n`n"

        if ($MobileRequirement -eq "Yes") {
            $NoteData += "### Mobile - Identifier`n"
            $NoteData += "$MobileIdentifier`n`n"
        }

        if (-not [string]::IsNullOrEmpty($AdditionalNotes)) {
            $NoteData += "### Additional Notes`n"
            $NoteData += "$AdditionalNotes`n`n"
        }

        return $NoteData
    }
    catch {
        Write-Error -Message "Failed to set SEA New User initial ticket note: $_"
        return $null
    }
}

function Set-TicketNoteInitialExitUser {
    param (
        [string]$EmployeeTicket,
        [string]$EndDate
    )

    try {
        $NoteData = "### Pre-Approved Submission`n"
        $NoteData += "This submission is pre-approved. This submission will be completed without prompting for approval.`n"
        $NoteData += "### Response Time`n"
        $NoteData += "It may take us up to four business hours to respond to your request and it may take several days to fulfill requests.`n`n"
        $NoteData += "If your user is exiting the organisation within the next two business days please complete the form in full and then call us on 07 3151 9000 to draw our attention to the exit date & time.`n`n"
        $NoteData += "If you need access for a departing user terminated immediately, please call us on 07 3151 9000 immediately.`n`n"
        $NoteData += "#### User To Offboard`n"
        $NoteData += "$EmployeeTicket`n`n"
        $NoteData += "#### Exit Date & Time`n"
        $NoteData += "$EndDate`n`n"

        return $NoteData
    }
    catch {
        Write-Error -Message "Failed to set SEA Exit User initial ticket note: $_"
        return $null
    }
}

function Get-TicketByEmployeeName {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FirstName,
        [Parameter(Mandatory = $true)]
        [string]$LastName
    )
    Write-Message "Searching for existing ticket for $FirstName $LastName"

    $searchString = "User Onboarding: $FirstName $LastName"
    
    try {
        $tickets = Get-CWM-Tickets -Condition "summary contains '$searchString'"
        
        # Handle both single objects and arrays
        if ($tickets) {
            # Convert to array if it's a single object
            if ($tickets -isnot [array]) {
                $tickets = @($tickets)
            }
            
            if ($tickets.Count -gt 0) {
                $recentTicket = $tickets | Sort-Object dateEntered -Descending | Select-Object -First 1
                Write-Message "Found ticket #$($recentTicket.id) for $FirstName $LastName"
                return $recentTicket.id
            }
        }
        else {
            Write-Message "No ticket found for $FirstName $LastName with search: $searchString" "WARN"
            return $null
        }
    }
    catch {
        Write-Message "Error searching for ticket: $_" "ERROR"
        return $null
    }
}

# ================================================================================
# WEBHOOK PAYLOAD FORMATTING FUNCTIONS
# ================================================================================

function Format-OnboardingPayload {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FirstName,
        [Parameter(Mandatory = $true)]
        [string]$LastName,
        [Parameter(Mandatory = $true)]
        [string]$StartDate,
        [Parameter(Mandatory = $true)]
        [string]$JobTitle,
        [Parameter(Mandatory = $false)]
        [string]$ManagerName = "",
        [Parameter(Mandatory = $false)]
        [string]$ManagerEmail = "",
        [Parameter(Mandatory = $true)]
        [string]$Department,
        [Parameter(Mandatory = $true)]
        [string]$Location,
        [Parameter(Mandatory = $true)]
        [string]$Persona,
        [Parameter(Mandatory = $false)]
        [string]$MobilePersonal = "",
        [Parameter(Mandatory = $false)]
        [string]$ComputerRequirement = "No",
        [Parameter(Mandatory = $false)]
        [string]$ComputerIdentifier = "",
        [Parameter(Mandatory = $false)]
        [string]$MobileRequirement = "No",
        [Parameter(Mandatory = $false)]
        [string]$MobileIdentifier = "",
        [Parameter(Mandatory = $false)]
        [string]$AdditionalNotes = "",
        [Parameter(Mandatory = $false)]
        [string]$TicketId = ""
    )

    $mailDomain = "seasonsliving.com.au"

    $payload = @{
        "cwpsa_ticket"                         = [int]$TicketId
        [string]"status_request_type"          = "user_onboarding"
        [string]"status_source"                = "Employment Hero"
        [string]"status_preapproved"           = "Yes"
        [string]"status_approved"              = "No"
        [string]"status_sspr"                  = "No"
        [string]"status_user_created"          = "No"
        [string]"status_license_assigned"      = "No"
        [string]"status_offboarding_scheduled" = "No"
        [string]"user_firstname"               = $FirstName
        [string]"user_lastname"                = $LastName
        [string]"user_fullname"                = "$($FirstName) $($LastName)"
        [string]"user_displayname"             = "$($FirstName) $($LastName)"
        [string]"user_password_reset_required" = "Yes"
        [string]"user_startdate"               = $StartDate
        [string]"user_company"                 = "Seasons Living"
        [string]"user_department"              = $Department
        [string]"user_title"                   = $JobTitle
        [string]"user_member_type"             = "Member"
        [string]"task_cs_operation01"          = "No"
        [string]"organisation_persona"         = $Persona
        [string]"organisation_site"            = $Location
        [string]"organisation_environment"     = "Cloud"
        [string]"asset02_required"             = $MobileRequirement
        [string]"asset01_required"             = $ComputerRequirement
        [string]"exchange_microsoft_domain"    = $mailDomain
        [string]"exchange_email_domain"        = $mailDomain
    }

    if (-not [string]::IsNullOrEmpty($ManagerName)) {
        $payload["user_manager_name"] = $ManagerName
    }
    
    if (-not [string]::IsNullOrEmpty($ManagerEmail)) {
        $payload["user_manager_upn"] = $ManagerEmail
    }

    if (-not [string]::IsNullOrEmpty($MobilePersonal)) {
        $payload["user_mobile"] = $MobilePersonal
    }
    
    if (-not [string]::IsNullOrEmpty($MobileIdentifier)) {
        $payload["asset01_tag"] = $MobileIdentifier
    }
    
    if (-not [string]::IsNullOrEmpty($ComputerIdentifier)) {
        $payload["asset02_tag"] = $ComputerIdentifier
    }

    if (-not [string]::IsNullOrEmpty($AdditionalNotes)) {
        $payload["task_operation01"] = $AdditionalNotes
    }

    return $payload
}

function Format-OnboardedPayload {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TicketId,
        [Parameter(Mandatory = $true)]
        [string]$MobilePersonal
    )

    $payload = @{
        "cwpsa_ticket"                = [int]$TicketId
        [string]"status_request_type" = "user_onboarded"
        [string]"status_source"       = "Employment Hero"
        [string]"user_mobile"         = $MobilePersonal
    }

    return $payload
}

function Format-OffboardingPayload {
    param (
        [Parameter(Mandatory = $true)]
        [string]$EmployeeName,
        [Parameter(Mandatory = $true)]
        [string]$EmployeeEmail,
        [Parameter(Mandatory = $true)]
        [string]$EndDate,
        [Parameter(Mandatory = $false)]
        [string]$ManagerName = "",
        [Parameter(Mandatory = $false)]
        [string]$ManagerEmail = "",
        [Parameter(Mandatory = $false)]
        [string]$TicketId = ""
    )

    $payload = @{
        "cwpsa_ticket"                = [int]$TicketId
        [string]"status_request_type" = "user_offboarding"
        [string]"status_source"       = "Employment Hero"
        "status_preapproved"          = "Yes"
        [string]"user_fullname"       = $EmployeeName
        [string]"user_upn"            = $EmployeeEmail
        [string]"user_primary_smtp"   = $employeeEmail
        [string]"user_enddate"        = $EndDate
    }
    
    if (-not [string]::IsNullOrEmpty($ManagerName)) {
        $payload["user_manager_name"] = $ManagerName
    }
    
    if (-not [string]::IsNullOrEmpty($ManagerEmail)) {
        $payload["user_manager_upn"] = $ManagerEmail
    }

    return $payload
}

# ================================================================================
# WEBHOOK COMMUNICATION FUNCTIONS
# ================================================================================

function Send-ToWebhookHelper {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$Payload,
        [Parameter(Mandatory = $false)]
        [bool]$Test = $false
    )

    $webhookUrl = "https://1ce978e3-0165-4625-bc15-bdcdfaeb9c7e.webhook.ae.azure-automation.net/webhooks?token=Q88nSvS565POfFpyw7Wo6fciyjG%2beX7rJjNJ%2f%2bjC0RA%3d"
    $JsonPayload = $Payload | ConvertTo-Json -Depth 10

    Write-Message "[INFO]: Payload to be sent:"
    Write-Message $JsonPayload

    if (-not $Test) {
        try {
            Write-Message "[INFO]: Sending webhook request to $webhookUrl"
            $response = Invoke-RestMethod -Method Post -Uri $webhookUrl -Body $JsonPayload -ContentType 'application/json'
            Write-Message "[INFO]: Webhook response: $($response | ConvertTo-Json -Compress)"
            return $true
        }
        catch {
            Write-Message "[ERROR]: Failed to send webhook request: $_" "ERROR"
            return $false
        }
    }
    else {
        Write-Message "[TEST]: Test mode enabled, webhook not sent"
        return $true
    }
}

# ================================================================================
# MAIN PROCESSING FUNCTIONS
# ================================================================================

function New-EmployeeCreated {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$PreCheckData,
        [Parameter(Mandatory = $true)]
        [object]$Employee,
        [Parameter(Mandatory = $true)]
        [string]$EmployeeId,
        [Parameter(Mandatory = $true)]
        [string]$EHOrganizationId,
        [Parameter(Mandatory = $true)]
        [string]$EHAuthorization
    )

    Write-Message "Starting New-EmployeeCreated process."

    $ClientIdentifier = "SEA"
    $ContactName = "System Notifications"
    $ServiceBoard = "ASIO"
    $TicketStatus = "New (portal)~"
    $TicketType = "Request"
    $TicketSubType = "Account Management"
    $TicketItem = "ADD User Account"

    Write-Message "Getting company and contact information."
    $CWMCompany = Get-CWM-Company -ClientIdentifier $ClientIdentifier
    if ($null -eq $CWMCompany) { Write-Message "Failed to get company information." ; return }
    $CWMContact = Get-CWM-Contact -ClientIdentifier $ClientIdentifier -ContactName $ContactName

    Write-Message "Getting ConnectWise Manage Locations."
    $CWMLocations = @()
    $ParentId = $CWMCompany.id
    $Sites = Get-CWM-CompanySites -CompanyId $ParentId
    if ($null -eq $Sites) { Write-Message "Failed to get company sites." ; return }
    $CWMLocations = $Sites.Name

    Write-Message "Getting board information."
    $CWMBoard = Get-CWM-Board -ServiceBoardName $ServiceBoard
    $CWMBoardStatus = if ($TicketStatus -ne '') { Get-CWM-BoardStatus -ServiceBoardName $ServiceBoard -ServiceBoardStatusName $TicketStatus }
    else { $null }
    $CWMBoardType = if ($TicketType -ne '') { Get-CWM-BoardType -ServiceBoardName $ServiceBoard -ServiceBoardTypeName $TicketType }
    else { $null }
    $CWMBoardSubType = if ($TicketSubType -ne '') { Get-CWM-BoardSubType -ServiceBoardName $ServiceBoard -ServiceBoardSubTypeName $TicketSubType }
    else { $null }
    $CWMBoardItem = if ($TicketItem -ne '') { Get-CWM-BoardItem -ServiceBoardName $ServiceBoard -ServiceBoardTypeName $TicketType -ServiceBoardSubTypeName $TicketSubType -ServiceBoardItemName $TicketItem }
    else { $null }
    
    Write-Message "Setting employee information."
    $FirstName = $Employee.first_name
    $LastName = $Employee.last_name
    $MobilePersonal = $Employee.personal_mobile_number
    $JobTitle = $Employee.job_title
    $EmployeeLocation = $Employee.location

    Write-Message "Getting employee manager."
    $ManagerName = ""
    $ManagerEmail = ""
    $ManagerTicket = ""

    if ($Employee.primary_manager -and $Employee.primary_manager.id) {
        $EmployeeManagerId = $Employee.primary_manager.id
        $Manager = Get-Employee -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization -EmployeeId $EmployeeManagerId
        if ($null -eq $Manager) {
            Write-Message "Failed to get manager details."
            return
        }
        else {
            $ManagerName = "$($Manager.first_name) $($Manager.last_name)"
            $ManagerEmail = $Manager.company_email
            Write-Output $Manager
            $ManagerTicket = if ($ManagerName -and $ManagerEmail) { "$ManagerName <$ManagerEmail>" } else { "" }
            Write-Output $ManagerTicket
        }
    }
    else { Write-Message "No manager assigned for this employee." }
    
    Write-Message "Getting team details."
    $Teams = Get-Teams -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization
    if ($null -eq $Teams) { Write-Message "Failed to get team details." ; return }
    else { Write-Output $Teams }
 
    Write-Message "Getting employee team names."
    $EmployeeTeamNames = Get-EmployeeTeamNames -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization -EmployeeId $EmployeeId -Teams $Teams
    if ($null -eq $EmployeeTeamNames) { Write-Message "Failed to get employee team names." ; return }
    else { Write-Output $EmployeeTeamNames }

    Write-Message "Setting employee start date."
    $parsedDate = $null
    if (Test-DateValid -dateString $Employee.start_date -parsedDate ([ref]$parsedDate)) {
        $StartDate = Set-EmployeeDate -inputDate $Employee.start_date -type "start"
        $StartDateForPayload = Set-EmployeeDateForPayload -inputDate $Employee.start_date
        Write-Message "Formatted start date: $StartDate"
        Write-Message "Formatted start date for payload: $StartDateForPayload"
    }
    else { Write-Message "Invalid date format for start date." }
    
    Write-Message "Setting Persona."
    $Persona = Set-Persona -EmployeeTeamNames $EmployeeTeamNames -PersonaLocations $PersonaLocations -PersonaRoles $PersonaRoles
    if ($null -eq $Persona) { Write-Message "Failed to set Persona." ; return }
    else { Write-Output $Persona }

    Write-Message "Setting department."
    $Department = Set-Department -Persona $Persona -PersonaRoles $PersonaRoles
    if ($null -eq $Department) { Write-Message "Failed to set department." ; return }
    else { Write-Output $Department }

    Write-Message "Setting location details."
    $Location = Get-Location -EmployeeLocation $EmployeeLocation -CWMLocations $CWMLocations
    if ($null -eq $Location) { Write-Message "Failed to set location." ; return }
    else { Write-Output $Location }

    Write-Message "Setting role."
    $Role = Set-Role -Persona $Persona -PersonaRoles $PersonaRoles
    if ($null -eq $Role) { Write-Message "Failed to set role." ; return }
    else { Write-Output $Role }

    Write-Message "Setting computer requirements."
    $ComputerRequirementDetails = Set-ComputerRequirement -Role $Role -PersonaRoles $PersonaRoles
    $ComputerRequirement = $ComputerRequirementDetails[0]
    $ComputerIdentifier = $ComputerRequirementDetails[1]
    Write-Output $ComputerRequirement
    Write-Output $ComputerIdentifier

    Write-Message "Setting mobile requirements."
    $MobileRequirementDetails = Set-MobileRequirement -Role $Role -PersonaRoles $PersonaRoles
    $MobileRequirement = $MobileRequirementDetails[0]
    $MobileIdentifier = $MobileRequirementDetails[1]
    Write-Output $MobileRequirement
    Write-Output $MobileIdentifier

    Write-Message "Setting additional notes."
    $AdditionalNotes = Set-AdditionalNotes -Role $Role
    Write-Output $AdditionalNotes

    Write-Message "Setting ticket summary and initial note."
    $CWMTicketSummary = Set-TicketSummaryNewUser -FirstName $FirstName -LastName $LastName -StartDate $StartDate
    $CWMTicketNoteInitial = Set-TicketNoteInitialNewUser -FirstName $FirstName `
        -LastName $LastName `
        -StartDate $StartDate `
        -MobilePersonal $MobilePersonal `
        -Company $Company `
        -JobTitle $JobTitle `
        -ManagerTicket $ManagerTicket `
        -Department $Department `
        -Location $Location `
        -Persona $Persona `
        -ComputerRequirement $ComputerRequirement `
        -ComputerIdentifier $ComputerIdentifier `
        -MobileRequirement $MobileRequirement `
        -MobileIdentifier $MobileIdentifier `
        -AdditionalNotes $AdditionalNotes

    Write-Message "Creating ticket."
    $Ticket = New-CWM-Ticket -CWMCompany $CWMCompany `
        -CWMContact $CWMContact `
        -CWMBoard $CWMBoard `
        -CWMTicketSummary $CWMTicketSummary `
        -CWMTicketNoteInitial $CWMTicketNoteInitial `
        -CWMBoardStatus $CWMBoardStatus `
        -CWMBoardType $CWMBoardType `
        -CWMBoardSubType $CWMBoardSubType `
        -CWMBoardItem $CWMBoardItem
        
    if ($null -eq $Ticket) {
        Write-Message "Failed to create ticket. Please check ConnectWise Manage connection and permissions." "ERROR"
        return
    }
    else {
        Write-Message "Successfully created onboarding ticket #$($Ticket.id) for $FirstName $LastName with start date $StartDate"
    }
    
    Write-Message "Adding Employment Hero note to ticket."
    $EmploymentHeroNote = "[Employment Hero] This user will be created without a mobile number until they have signed the contract in Employment Hero.`n`nThe mobile number will be added to this ticket once available.`n`nIf the mobile number is not in the system 1 day before the due date, an automated email will be sent advising this and instructing them to get the user to complete onboarding in Employment Hero."
    
    try {
        New-CWM-TicketNote -TicketId $Ticket.id -NoteText $EmploymentHeroNote
        Write-Message "Successfully added Employment Hero note to ticket #$($Ticket.id)"
    }
    catch {
        Write-Message "Failed to add Employment Hero note to ticket #$($Ticket.id): $_" "WARN"
    }
        
    Write-Message "Creating webhook payload."
    $TicketId = $Ticket.id
    $WebhookPayload = Format-OnboardingPayload -FirstName $FirstName `
        -LastName $LastName `
        -StartDate $StartDateForPayload `
        -JobTitle $JobTitle `
        -ManagerName $ManagerName `
        -ManagerEmail $ManagerEmail `
        -Department $Department `
        -Location $Location `
        -Persona $Persona `
        -MobilePersonal $MobilePersonal `
        -ComputerRequirement $ComputerRequirement `
        -ComputerIdentifier $ComputerIdentifier `
        -MobileRequirement $MobileRequirement `
        -MobileIdentifier $MobileIdentifier `
        -AdditionalNotes $AdditionalNotes `
        -TicketId $TicketId
        
    Write-Message "Sending payload to webhook helper."
    $Result = Send-ToWebhookHelper -Payload $WebhookPayload
        
    if ($Result) {
        Write-Message "Successfully sent onboarding data to webhook helper."
    }
    else {
        Write-Message "Failed to send onboarding data to webhook helper." "ERROR"
    }
    Write-Message "New-EmployeeCreated process completed."
}

function New-EmployeeOnboarded {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$PreCheckData,
        [Parameter(Mandatory = $true)]
        [object]$Employee,
        [Parameter(Mandatory = $true)]
        [string]$EmployeeId,
        [Parameter(Mandatory = $true)]
        [string]$EHOrganizationId,
        [Parameter(Mandatory = $true)]
        [string]$EHAuthorization
    )

    Write-Message "Starting New-EmployeeOnboarded process - updating mobile number."
    
    $FirstName = $Employee.first_name
    $LastName = $Employee.last_name
    $MobilePersonal = $Employee.personal_mobile_number
    
    Write-Message "Looking for original onboarding ticket for $FirstName $LastName"
    $TicketNumber = Get-TicketByEmployeeName -FirstName $FirstName -LastName $LastName
    
    if ($TicketNumber) {
        Write-Message "Creating mobile number update payload."
        $WebhookPayload = Format-OnboardedPayload -TicketId $TicketNumber -MobilePersonal $MobilePersonal
        
        Write-Message "Sending mobile number update to webhook helper for ticket #$TicketNumber"
        $Result = Send-ToWebhookHelper -Payload $WebhookPayload
        
        if ($Result) {
            Write-Message "Successfully sent mobile number update for $FirstName $LastName"
            
            $TicketNote = "[Employment Hero] - Mobile number added [$MobilePersonal]"
            New-CWM-TicketNote -TicketId $TicketNumber -TicketNote $TicketNote -InternalFlag $true
            Write-Message "Added ticket note for mobile number update on ticket #$TicketNumber"
        }
        else {
            Write-Message "Failed to send mobile number update for $FirstName $LastName" "ERROR"
        }
    }
    else {
        Write-Message "Could not find original ticket for $FirstName $LastName - unable to update mobile number" "ERROR"
    }

    Write-Message "New-EmployeeOnboarded process completed."
}

function New-EmployeeOffboarding {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$PreCheckData,
        [Parameter(Mandatory = $true)]
        [object]$Employee,
        [Parameter(Mandatory = $true)]
        [string]$EmployeeId,
        [Parameter(Mandatory = $true)]
        [string]$EHOrganizationId,
        [Parameter(Mandatory = $true)]
        [string]$EHAuthorization
    )

    Write-Message "Starting New-EmployeeOffboarding process."

    $ClientIdentifier = "SEA"
    $ContactName = "System Notifications"
    $ServiceBoard = "ASIO"
    $TicketStatus = "New (portal)~"
    $TicketType = "Request"
    $TicketSubType = "Account Management"
    $TicketItem = "EXIT User"

    Write-Message "Getting company and contact information."
    $CWMCompany = Get-CWM-Company -ClientIdentifier $ClientIdentifier
    if ($null -eq $CWMCompany) { Write-Message "Failed to get company information." ; return }
    $CWMContact = Get-CWM-Contact -ClientIdentifier $ClientIdentifier -ContactName $ContactName
    
    Write-Message "Getting board information."
    $CWMBoard = Get-CWM-Board -ServiceBoardName $ServiceBoard
    $CWMBoardStatus = if ($TicketStatus -ne '') { Get-CWM-BoardStatus -ServiceBoardName $ServiceBoard -ServiceBoardStatusName $TicketStatus }
    else { $null }
    $CWMBoardType = if ($TicketType -ne '') { Get-CWM-BoardType -ServiceBoardName $ServiceBoard -ServiceBoardTypeName $TicketType }
    else { $null }
    $CWMBoardSubType = if ($TicketSubType -ne '') { Get-CWM-BoardSubType -ServiceBoardName $ServiceBoard -ServiceBoardSubTypeName $TicketSubType }
    else { $null }
    $CWMBoardItem = if ($TicketItem -ne '') { Get-CWM-BoardItem -ServiceBoardName $ServiceBoard -ServiceBoardTypeName $TicketType -ServiceBoardSubTypeName $TicketSubType -ServiceBoardItemName $TicketItem }
    else { $null }    

    Write-Message "Setting employee information."
    $FirstName = $Employee.first_name
    $LastName = $Employee.last_name
    $EmployeeName = "$($Employee.first_name) $($Employee.last_name)"
    $EmployeeEmail = $Employee.company_email
    $EmployeeTicket = "$EmployeeName <$EmployeeEmail>"
    if ($null -eq $EmployeeTicket) { Write-Message "Failed to set employee details." ; return }
    else { Write-Message $EmployeeTicket }

    Write-Message "Getting employee manager."
    $ManagerName = ""
    $ManagerEmail = ""
    $ManagerTicket = ""

    if ($Employee.primary_manager -and $Employee.primary_manager.id) {
        $EmployeeManagerId = $Employee.primary_manager.id
        $Manager = Get-Employee -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization -EmployeeId $EmployeeManagerId
        if ($null -eq $Manager) {
            Write-Message "Failed to get manager details."
            return
        }
        else {
            $ManagerName = "$($Manager.first_name) $($Manager.last_name)"
            $ManagerEmail = $Manager.company_email
            Write-Output $Manager
            $ManagerTicket = if ($ManagerName -and $ManagerEmail) { "$ManagerName <$ManagerEmail>" } else { "" }
            Write-Output $ManagerTicket
        }
    }
    else { Write-Message "No manager assigned for this employee." }

    Write-Message "Setting employee end date."
    
    $EndDate = $null
    $EndDateForPayload = $null
    $EmploymentHistory = Get-EmployeeEmploymentHistory -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization -EmployeeId $EmployeeId
    if ($null -ne $EmploymentHistory) {
        Write-Message "Found employment history records: $($EmploymentHistory.Count)"
        
        $ActiveEmployment = $EmploymentHistory | Where-Object { $null -ne $_.end_date } | Sort-Object end_date -Descending | Select-Object -First 1
        
        if ($null -ne $ActiveEmployment -and $null -ne $ActiveEmployment.end_date) {
            Write-Message "Found end date in employment history: $($ActiveEmployment.end_date)"
            $parsedDate = $null
            if (Test-DateValid -dateString $ActiveEmployment.end_date -parsedDate ([ref]$parsedDate)) {
                $EndDate = Set-EmployeeDate -inputDate $ActiveEmployment.end_date -type "end"
                $EndDateForPayload = Set-EmployeeDateForPayload -inputDate $ActiveEmployment.end_date
                Write-Message "Formatted end date: $EndDate"
                Write-Message "Formatted end date for payload: $EndDateForPayload"
            }
        }
    }
    if ($null -eq $EndDate) { Write-Message "No end date found in employment history." ; return }

    Write-Message "Setting ticket summary and initial note."
    $CWMTicketSummary = Set-TicketSummaryExitUser -FirstName $FirstName -LastName $LastName -EndDate $EndDate
    $CWMTicketNoteInitial = Set-TicketNoteInitialExitUser -EmployeeTicket $EmployeeTicket -EndDate $EndDate

    Write-Message "Creating ticket."
    $Ticket = New-CWM-Ticket `
        -CWMCompany $CWMCompany `
        -CWMContact $CWMContact `
        -CWMBoard $CWMBoard `
        -CWMTicketSummary $CWMTicketSummary `
        -CWMTicketNoteInitial $CWMTicketNoteInitial `
        -CWMBoardStatus $CWMBoardStatus `
        -CWMBoardType $CWMBoardType `
        -CWMBoardSubType $CWMBoardSubType `
        -CWMBoardItem $CWMBoardItem
        
    if ($null -eq $Ticket) {
        Write-Message "Failed to create ticket. Please check ConnectWise Manage connection and permissions." "ERROR"
        return
    }
    else {
        Write-Message "Successfully created offboarding ticket #$($Ticket.id) for $FirstName $LastName with end date $EndDate"
        
        Write-Message "Creating webhook payload using Format-OffboardingPayload."
        $TicketId = $Ticket.id
        $WebhookPayload = Format-OffboardingPayload -EmployeeName $EmployeeName `
            -EmployeeEmail $EmployeeEmail `
            -EndDate $EndDateForPayload `
            -ManagerName $ManagerName `
            -ManagerEmail $ManagerEmail `
            -TicketId $TicketId
            
        Write-Message "Sending simplified payload to webhook helper."
        $Result = Send-ToWebhookHelper -Payload $WebhookPayload
        
        if ($Result) {
            Write-Message "Successfully sent offboarding data to webhook helper."
        }
        else {
            Write-Message "Failed to send offboarding data to webhook helper." "ERROR"
        }
    }

    Write-Message "New-EmployeeOffboarding process completed."
}

# ================================================================================
# MAIN EXECUTION
# ================================================================================

function ProcessMainFunction {
    Write-Message "Starting ProcessMainFunction."

    try {
        $ParsedWebhookData = Get-WebhookData -WebhookData $WebhookData
        Write-Message "Webhook data parsed successfully"
    }
    catch {
        Write-ScriptError "Failed to parse WebhookData. Error: $_"
    }

    Write-Message "The Webhook Header"
    Write-Message $($ParsedWebhookData.RequestHeader)
    
    Write-Message 'The Webhook Name'
    Write-Message $($ParsedWebhookData.WebhookName)
    
    Write-Message 'Extracted Employee Data'
    Write-Message ($ParsedWebhookData.EmployeeData | ConvertTo-Json -Depth 2)
    
    Write-Message 'Extracted Event'
    Write-Message $ParsedWebhookData.Event

    Write-Message "Getting PreCheck details."
    $PreCheckData = Get-PreCheck -ParsedWebhookData $ParsedWebhookData
        
    if (-not $PreCheckData) { Write-ScriptError "PreCheckData is empty. Exiting script." }
    else {
        $EmployeeId = $PreCheckData.EmployeeId
        $EmployeeEvent = $PreCheckData.Event
        Write-Message 'EmployeeId'
        Write-Message $EmployeeId
        Write-Message 'Event Type'
        Write-Message $EmployeeEvent
    }
    
    Write-Message "Getting Secrets from AzureVault."
    $AzKeyVaultName = Get-AutomationVariable -Name 'AzKeyVaultName'
    $Secrets = Get-Secrets -AzKeyVaultName $AzKeyVaultName

    Write-Message "Connecting to CWManage."
    ConnectCWM -AzKeyVaultName $AzKeyVaultName `
        -CWMClientIdName 'MIT-CWAApi-ClientId' `
        -CWMPublicKeyName 'MIT-CWMApi-PubKey' `
        -CWMPrivateKeyName 'MIT-CWMApi-PrivateKey' `
        -CWMCompanyIdName 'MIT-CWMApi-CompanyId' `
        -CWMUrlName 'MIT-CWMApi-Server'

    Write-Message "Setting Secret Variables."
    $EHclient_Id = $Secrets.EHclient_Id
    $EHclient_secret = $Secrets.EHclient_secret
    $EHcode = $Secrets.EHcode
    $EHrefresh_token = $Secrets.EHrefresh_token
    $EHOrganizationId = $Secrets.EHOrganizationId
    $EHRedirectUri = $Secrets.EHRedirectUri
    $EHAuthorization = Get-EHAuthorization -EHclient_Id $EHclient_Id -EHclient_secret $EHclient_secret -EHcode $EHcode -EHrefresh_token $EHrefresh_token -EHRedirectUri $EHRedirectUri

    Write-Message "Getting employee details."
    $Employee = Get-Employee -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization -EmployeeId $EmployeeId
    if ($null -eq $Employee) { throw "Failed to fetch employee details." }
    else { Write-Output $Employee }

    Write-Message "Checking event type."
    if ($EmployeeEvent -eq "employee_created") {
        Write-Message 'New User event'
        New-EmployeeCreated -PreCheckData $PreCheckData -Employee $Employee -EmployeeId $EmployeeId -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization
    }
    elseif ($EmployeeEvent -eq "employee_offboarding") {
        Write-Message 'Exit User event'
        New-EmployeeOffboarding -PreCheckData $PreCheckData -Employee $Employee -EmployeeId $EmployeeId -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization
    }
    elseif ($EmployeeEvent -eq "employee_onboarded") {
        Write-Message 'User Onboarded event'
        New-EmployeeOnboarded -PreCheckData $PreCheckData -Employee $Employee -EmployeeId $EmployeeId -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization
    }
    else { Write-Message "Unhandled event type: $EmployeeEvent" "WARN" }

    Write-Message "ProcessMainFunction completed."
}

ProcessMainFunction

Write-Message "Script execution completed."