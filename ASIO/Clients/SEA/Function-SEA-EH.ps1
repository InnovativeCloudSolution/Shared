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
    $formats = @("dd/MM/yyyy hh:mm tt", "dd/MM/yyyy hh:mmtt", "yyyy-MM-ddTHH:mm:sszzz", "yyyy-MM-dd")
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
        Write-Host "Failed to fetch details."
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

function Set-Persona {
    param (
        $TeamNames,
        $Location,
        $AccessRoles
    ) 

    $MatchedLocation = $Location.ToUpper()
    if ($MatchedLocation -like "MANGO HILL*") {
        $MatchedLocation = "MANGO HILL"
    }

    $MatchedRole = ""

    foreach ($TeamName in $TeamNames) {
        foreach ($AccessRole in $AccessRoles) {
            if ($AccessRole -eq $TeamName) {
                $MatchedRole = $AccessRole
            }
        }
    }

    return "$MatchedRole-$MatchedLocation"
}

function Set-Department {
    param (
        $Persona,
        $AccessRoles
    )
    $Role = $null
    $Department = "Seasons"
    
    foreach ($AccessRole in $AccessRoles) {
        if ($Persona.IndexOf($AccessRole, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
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
        $Persona,
        $AccessRoles
    )
    $Role = $null
    
    foreach ($AccessRole in $AccessRoles) {
        if ($Persona.IndexOf($AccessRole, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            $Role = $AccessRole
            break
        }
    }
    return $Role
}

function Get-Location {
    param (
        $Location,
        $CWMLocations
    )
    $MatchedLocation = ""
        foreach ($CWMLocation in $CWMLocations) {
            if ($CWMLocation.IndexOf($Location, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $MatchedLocation = $CWMLocation
            }
        }
    return $MatchedLocation
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
                $ComputerIdentifier = "Please check with Britt if the new user is hot-desking, using a shared laptop, or getting a new laptop"
            }
            else {
                $ComputerIdentifier = "Please check with Britt if the new user is getting a new endpoint or re-purposing an endpoint"
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
                $MobileIdentifier = "Please check with Britt on PCW mobile phone asset tagging"
            }
            else {
                $MobileIdentifier = "Please check with Britt if the new user is getting a new mobile or re-purposing a mobile"
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
        [string]$Persona,
        [string]$ComputerRequirement,
        [string]$ComputerIdentifier,
        [string]$MobileRequirement,
        [string]$MobileIdentifier,
        [string]$AdditionalNotes
    )

    try {
        $NoteData = "### New User Setup`n"
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
        $NoteData += "### Persona`n"
        $NoteData += "$Persona`n`n"
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
            $NoteData += "### Mailbox Delegate Access`n"
            $NoteData += "$ManagerTicket"
        }

        return $NoteData
    }
    catch {
        Write-Error -Message "Failed to set SEA Exit User initial ticket note: $_"
        return $null
    }
}

function Set-License {
    param (
        [string]$LicenseGroup
    )
    try {
        $Licenses = switch ($LicenseGroup) {
            "SG.License.M365BusinessPremium;" { 'NCE Microsoft 365 Business Premium:Telstra:null:SPB:Monthly;' }
            "SG.License.OfficeF3;" { 'NCE Microsoft 365 F3:Telstra:null:SPE_F1:Monthly;' }
            default { "" }
        }

        return $Licenses
    }
    catch {
        throw "Failed to determine licenses: $_"
    }
}

function Set-NonSupportUser {
    param (
        [string]$Department
    )
    try {
        $NonSupportUser = if ($Department -eq "F3 Seasons") {
            $True
        } else {
            $False
        }

        return $NonSupportUser
    }
    catch {
        throw "Failed to determine Non-Support user status: $_"
    }
}
