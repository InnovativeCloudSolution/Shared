function Write-MessageLog {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Output "$timestamp [$Level] $Message"
}

function Get-WebhookData {
    param (
        [Parameter(Mandatory = $true)]
        [object]$WebhookData
    )

    if (-not $WebhookData) {
        Write-Error "WebhookData is null. Cannot convert from JSON."
    }

    $Data = $WebhookData | ConvertFrom-Json

    return @{
        EmployeeData = $Data.data
        Event        = $Data.event
    }
}

function Get-PreCheck {
    param (
        [Parameter(Mandatory = $true)]
        [string]$WebhookData
    )

    $Result = Get-WebhookData -WebhookData $WebhookData
    $EmployeeData = $Result.EmployeeData
    $EmployeeId = $EmployeeData.id
    $Event = $Result.Event

    return @{
        Event        = $Event
        EmployeeData = $EmployeeData
        EmployeeId   = $EmployeeId
    }
}

function Write-Error {
    param (
        [string]$Message
    )
    Write-MessageLog $Message "ERROR"
    throw $Message
}

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

function Set-EmployeeDetails {
    param ($Employee)

    if ($null -eq $Employee) { 
        Write-Host "Failed to fetch manager details."
        return $null 
    }
    return "$($Employee.first_name) $($Employee.last_name) <$($Employee.company_email)>"
}

function Set-ManagerDetails {
    param ($ManagerDetails)

    if ($null -eq $ManagerDetails) { 
        Write-Host "Failed to fetch manager details."
        return $null 
    }
    return "$($ManagerDetails.first_name) $($ManagerDetails.last_name) <$($ManagerDetails.company_email)>"
}

function Set-AccessLevel {
    param (
        $TeamNames,
        $AccessLocations,
        $AccessRoles
    ) 
    if ($TeamNames.Count -eq 1 -or $TeamNames.Count -gt 2) {
        Write-Host "ERROR: Invalid number of team names."
        return $null
    }

    $MatchedLocation = ""
    $MatchedRole = ""

    foreach ($TeamName in $TeamNames) {
        foreach ($AccessLocation in $AccessLocations) {
            if ($AccessLocation -eq $TeamName) {
                $MatchedLocation = $AccessLocation
            }
        }

        foreach ($AccessRole in $AccessRoles) {
            if ($AccessRole -eq $TeamName) {
                $MatchedRole = $AccessRole
            }
        }
    }

    return "$MatchedLocation-$MatchedRole"
}

function Set-Department {
    param (
        $AccessLevel,
        $AccessRoles
    )
    $Role = $null
    $Department = "Seasons"
    
    foreach ($AccessRole in $AccessRoles) {
        if ($AccessLevel.IndexOf($AccessRole, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            $Role = $AccessRole
            if ($Role -eq "BUS DRIVER" -or $Role -eq "CHEF" -or $Role -eq "CLEANER" -or $Role -eq "HOSPITALITY" -or $Role -eq "MAINTENANCE" -or $Role -eq "PERSONAL CARE WORKER") {
                $Department = "F3 Seasons"
            }
            break
        }
    }
    return $Department
}

function Set-Role {
    param (
        $AccessLevel,
        $AccessRoles
    )
    $Role = $null
    
    foreach ($AccessRole in $AccessRoles) {
        if ($AccessLevel.IndexOf($AccessRole, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            $Role = $AccessRole
            break
        }
    }
    return $Role
}

function Get-Location {
    param (
        $TeamNames,
        $CWMLocations
    )
    $MatchedLocation = ""
    foreach ($TeamName in $TeamNames) {
        foreach ($CWMLocation in $CWMLocations) {
            if ($CWMLocation.IndexOf($TeamName, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $MatchedLocation = $CWMLocation
            }
        }
    }
    return $MatchedLocation
}

function Set-LocationDetails {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Sites,
        [Parameter(Mandatory = $true)]
        [string]$Location
    )

    $LocationDetails = @{}
    foreach ($Site in $Sites) {
        $LocationName = $Site.name
        if ($LocationName -eq $Location) {
            $LocationDetails = @{
                Address  = $Site.addressLine1
                City     = $Site.City
                State    = $Site.stateReference.name
                Postcode = $Site.zip
            }
            break
        }
    }
    return $LocationDetails
}

function Set-ComputerRequirement {
    param (
        $Role,
        $AccessRoles
    )
    $ComputerRequirement = "No"
    $ComputerIdentifier = $null
    
    if ($AccessRoles -contains $Role) {
        if ($Role -in @("BUS DRIVER", "CLEANER", "HOSPITALITY", "PERSONAL CARE WORKER")) {
            $ComputerRequirement = "No"
        }
        else {
            $ComputerRequirement = "Yes"
            if ($Role -in @("CHEF", "ENROLLED NURSE", "LIFESTYLE", "MAINTENANCE", "REGISTERED NURSE")) {
                $ComputerIdentifier = "Check with Britt if the new user is hot-desking, using a shared laptop, or getting a new laptop"
            }
            else {
                $ComputerIdentifier = "Check with Britt if the new user is getting a new endpoint or re-purposing an endpoint"
            }
        }
    }
    return @($ComputerRequirement, $ComputerIdentifier)
}

function Set-MobileRequirement {
    param (
        $Role,
        $AccessRoles
    )
    $MobileRequirement = "No"
    $MobileIdentifier = $null
    
    if ($AccessRoles -contains $Role) {
        if ($Role -ne "ADMINISTRATION") {
            $MobileRequirement = "Yes"
            if ($Role -in @("BUS DRIVER", "CLEANER", "HOSPITALITY", "LIFESTYLE", "PERSONAL CARE WORKER")) {
                $MobileIdentifier = "Check with Britt on PCW mobile phone asset tagging"
            }
            else {
                $MobileIdentifier = "Check with Britt if the new user is getting a new mobile or re-purposing a mobile"
            }
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
        $AdditionalNotes += "Part of SUPPORT OFFICE, check with Brittany Smith what shared mailboxes are needed`n"
        $AdditionalNotes += "Part of SUPPORT OFFICE, Add Reviewer calendar access to OutOfOfficeCalendar@seasonsliving.com.au"
    }
    return $AdditionalNotes
}

function Set-TicketSummaryNewUser {
    param (
        [string]$FirstName,
        [string]$LastName,
        [string]$StartDate
    )

    try {
        $SummaryData = "New Standard User: $FirstName $LastName, Start Date $StartDate"
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
        $SummaryData = "User Exit: $FirstName $LastName, departing on $EndDate"
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
        [string]$JobTitle,
        [string]$ManagerTicket,
        [string]$Department,
        [string]$Location,
        [string]$AccessLevel,
        [string]$LocationAddress,
        [string]$LocationCity,
        [string]$LocationState,
        [string]$LocationPostcode,
        [string]$MobilePersonal,
        [string]$ComputerRequirement,
        [string]$ComputerIdentifier,
        [string]$MobileRequirement,
        [string]$MobileIdentifier,
        [string]$AdditionalNotes
    )

    try {
        $NoteData = "### Pre-Approved Form`n"
        $NoteData += "This form is pre-approved and will be completed without requiring additional approval once submitted.`n`n"
        $NoteData += "### New User Setup`n"
        $NoteData += "The new user will be set up exactly as specified here. Please ensure all information is accurate, and avoid typing or spelling mistakes. Select the correct options. If you are unsure about any part of the form, call us at 07 3151 9000 before submitting.`n`n"
        $NoteData += "### Given Name(s)`n"
        $NoteData += "$FirstName`n`n"
        $NoteData += "### Last Name`n"
        $NoteData += "$LastName`n`n"
        $NoteData += "### Start Date`n"
        $NoteData += "$StartDate`n`n"
        $NoteData += "### Job Title`n"
        $NoteData += "$JobTitle`n`n"
        $NoteData += "### Manager`n"
        $NoteData += "$ManagerTicket`n`n"
        $NoteData += "### Department`n"
        $NoteData += "$Department`n`n"
        $NoteData += "### Location`n"
        $NoteData += "$Location`n`n"
        $NoteData += "### Access Level`n"
        $NoteData += "$AccessLevel`n`n"
        $NoteData += "### Street Address`n"
        $NoteData += "$LocationAddress`n`n"
        $NoteData += "### City`n"
        $NoteData += "$LocationCity`n`n"
        $NoteData += "### State`n"
        $NoteData += "$LocationState`n`n"
        $NoteData += "### Postcode`n"
        $NoteData += "$LocationPostcode`n`n"
        $NoteData += "### Personal Mobile Number`n"
        $NoteData += "$MobilePersonal`n`n"
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

        $NoteData += "### Confirmation of Information`n"
        $NoteData += "By pressing submit, you agree that all information you have provided is accurate. Errors in spelling or incorrect mobile numbers may cause delays in setting up the user's account."

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
        [string]$EndDate,
        [string]$ManagerTicket
    )

    try {
        $NoteData = "### Response Time`n"
        $NoteData += "It may take us up to four business hours to respond to your request and it may take several days to fulfill requests.`n`n"
        $NoteData += "If your user is exiting the organisation within the next two business days please complete the form in full and then call us on 07 3151 9000 to draw our attention to the exit date & time.`n`n"
        $NoteData += "If you need access for a departing user terminated immediately, please call us on 07 3151 9000 immediately.`n`n"
        $NoteData += "### User`n"
        $NoteData += "$EmployeeTicket`n`n"
        $NoteData += "### Exit Date and Time`n"
        $NoteData += "$EndDate`n`n"

        if (-not [string]::IsNullOrEmpty($ManagerTicket)) {
            $NoteData += "### Auto Response - Redirected`n"
            $NoteData += "$ManagerTicket"
        }

        return $NoteData
    }
    catch {
        Write-Error -Message "Failed to set SEA Exit User initial ticket note: $_"
        return $null
    }
}
