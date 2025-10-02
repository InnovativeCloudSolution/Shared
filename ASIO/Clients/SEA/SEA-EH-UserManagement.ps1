param (
    [Parameter(Mandatory = $false)]
    [object]$WebhookData
)

$WebhookData = Get-Content -Path .\Test-Webhook.json -Raw

. .\Function-ALL-General.ps1
. .\Function-ALL-CWM.ps1
. .\Function-ALL-MSGraph.ps1
. .\Function-ALL-MSGraph-SP.ps1
. .\Function-ALL-EH.ps1
. .\Function-SEA-EH.ps1

function ProcessMainFunction {
    Write-MessageLog -Message "Starting the Main Process."
    Write-MessageLog -Message "The Webhook Header"
    Write-MessageLog -Message $WebhookData.RequestHeader

    Write-MessageLog -Message "The Webhook Header Message"
    Write-MessageLog -Message $WebhookData.RequestHeader.Message

    Write-MessageLog -Message 'The Webhook Name'
    Write-MessageLog -Message $WebhookData.WebhookName

    Write-MessageLog -Message 'The Webhook Request Body'
    Write-MessageLog -Message $WebhookData.RequestBody

    Write-MessageLog -Message "Getting PreCheck details."
    # $PreCheckData = Get-PreCheck -WebhookData $WebhookData.RequestBody
    $PreCheckData = Get-PreCheck -WebhookData $WebhookData

    try {
        Initialize
        Write-MessageLog -Message "Initialization completed successfully."
    }
    catch {
        Write-ErrorLog -Message "Failed during initialization. Error: $_"
        return
    }

    try {
        Write-MessageLog -Message "Connecting to Microsoft Graph."
        # Connect to Microsoft Graph
        ConnectMSGraphTest -TenantUrl $TenantUrl
        # ConnectMSGraph -AzKeyVaultName $AzKeyVaultName -TenantUrl $TenantUrl -ClientIdSecretName $ClientIdSecretName -ClientSecretSecretName $ClientSecretSecretName
        Write-MessageLog -Message "Connected to Microsoft Graph successfully."
    }
    catch {
        Write-ErrorLog -Message "Failed to connect to Microsoft Graph. Error: $_"
        return
    }

    try {
        Write-MessageLog -Message "Connecting to ConnectWise Manage."
        # Connect to ConnectWise Manage
        ConnectCWMTest
        # ConnectCWM -AzKeyVaultName $AzKeyVaultName -CWMClientIdName $CWMClientIdName -CWMPublicKeyName $CWMPublicKeyName -CWMPrivateKeyName $CWMPrivateKeyName -CWMCompanyIdName $CWMCompanyIdName -CWMUrlName $CWMUrlName
        Write-MessageLog -Message "Connected to ConnectWise Manage successfully."
    }
    catch {
        Write-ErrorLog -Message "Failed to connect to ConnectWise Manage. Error: $_"
        return
    }

    if ($PreCheckData) {
        Write-MessageLog -Message "Webhook triggered action."

        # $Result = $WebhookData.RequestBody | ConvertFrom-Json
        $Result = $WebhookData | ConvertFrom-Json
        $EmployeeData = $Result.data
        $EmployeeId = $EmployeeData.id
        $Event = $Result.event

        if ($Event -eq "employee_created") {
            Write-MessageLog -Message "Triggered Event: Employee Created."
            New-EmployeeCreated
        }
        elseif ($Event -eq "employee_onboarded") {
            Write-MessageLog -Message "Triggered Event: Employee Onboarded."
            New-EmployeeOnboarded
        }
        elseif ($Event -eq "employee_offboarding") {
            Write-MessageLog -Message "Triggered Event: Employee Offboarding."
            New-EmployeeOffboarding
        }
        else {
            Write-MessageLog -Message "Unknown event: $Event"
        }
    }
    Write-MessageLog -Message "Script execution completed."
}

function Initialize {
    Write-MessageLog -Message "Initializing script configuration."
    $Global:ClientIdSecretName = "MIT-AutomationApp-ClientID"
    $Global:ClientSecretSecretName = "MIT-AutomationApp-ClientSecret"
    # try {
    #     Write-MessageLog -Message "Retrieving Azure Key Vault name from automation variables."
    #     $Global:AzKeyVaultName = Get-AutomationVariable -Name 'AzKeyVaultName'
    #
    #     if ($null -eq $AzKeyVaultName) {
    #         Write-ErrorLog -Message "Azure Key Vault name is missing. Check your automation variables."
    #         throw "Azure Key Vault name is not defined."
    #     }
    #
    #     Write-MessageLog -Message "Azure Key Vault name retrieved successfully: $AzKeyVaultName."
    # }
    # catch {
    #     Write-ErrorLog -Message "Failed to retrieve Azure Key Vault name. Error: $_"
    #     throw $_
    # }
    $Global:TenantUrl = 'manganoit.com.au'
    Write-MessageLog -Message "Tenant URL set to $Global:TenantUrl."
    $Global:CWMClientIdName = "MIT-CWMApi-ClientId"
    $Global:CWMPublicKeyName = "MIT-CWMApi-PubKey"
    $Global:CWMPrivateKeyName = "MIT-CWMApi-PrivateKey"
    $Global:CWMCompanyIdName = "MIT-CWMApi-CompanyId"
    $Global:CWMUrlName = "MIT-CWMApi-Url"
    $Global:SPSiteUrl = "https://manganoit.sharepoint.com/sites/Automation/SEA"
    $Global:SPSiteName = "SEA"
    $Global:SPFileName = "SEA-Matrix.csv"
    $Global:ClientIdSecretName = "SEA-EH-ClientID"
    $Global:ClientSecretSecretName = "SEA-EH-ClientSecret"
    $Global:CodeSecretName = "SEA-EH-Code"
    $Global:RefreshTokenSecretName = "SEA-EH-RefreshToken"
    $Global:RedirectUriSecretName = "SEA-EH-RedirectUri"
    $Global:EHOrganizationIdSecretName = "SEA-EH-OrganizationId"
    $Global:EHAuthorization = Get-EH-AuthorizationTest
    $Global:EHOrganizationId = "5023b874-f855-4aa4-94c3-07d9cdf441e4"
    $Global:AccessRoles = @(
        "ADMINISTRATION",
        "BUS DRIVER",
        "CARE COORDINATOR",
        "CARE MANAGER",
        "CARE SERVICES SCHEDULER",
        "CHEF",
        "CLEANER",
        "COMMUNITY MANAGEMENT",
        "ENROLLED NURSE",
        "HOSPITALITY",
        "LIFESTYLE",
        "MAINTENANCE",
        "PERSONAL CARE WORKER",
        "REGISTERED NURSE",
        "SUPPORT OFFICE"
    )
    Write-MessageLog -Message "Access roles configured."
    Write-MessageLog -Message "Script configuration initialized."
}

function New-EmployeeCreated {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$PreCheckData
    )

    Write-MessageLog "Starting New-EmployeeCreated process."

    $ClientIdentifier = "SEA"
    $ContactName = "System Notifications"
    $ServiceBoardName = "Pia"
    $ServiceBoardStatusName = "New (portal)~"
    $ServiceBoardTypeName = "Request"
    $ServiceBoardSubTypeName = "Profile"
    $ServiceBoardItemName = "New User"

    Write-MessageLog "Getting employee details."
    $Employee = Get-EH-Employee -EmployeeId $PreCheckData.EmployeeData.id -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization
    if ($Employee -contains "error") { Write-ErrorLog "Failed to fetch employee details." ; return }

    $FirstName = $Employee.first_name
    $LastName = $Employee.last_name
    $JobTitle = $Employee.job_title
    $Location = $Employee.location

    $parsedDate = $null
    if (Test-DateValid -dateString $Employee.start_date -parsedDate ([ref]$parsedDate)) {
        $StartDate = Set-EmployeeDate -inputDate $Employee.start_date -type "start"
    }
    else {
        Write-ErrorLog "Invalid date format for start date."
    }

    Write-MessageLog "Getting company and contact information."
    $CWMCompany = Get-CWM-Company -ClientIdentifier $ClientIdentifier
    if ($null -eq $CWMCompany) { 
        Write-ErrorLog "Failed to get company information." 
        return 
    }

    $CWMContact = Get-CWM-Contact -ClientIdentifier $ClientIdentifier -ContactName $ContactName
    if ($null -eq $CWMContact) { 
        Write-ErrorLog "Failed to get contact information." 
        return 
    }

    Write-MessageLog "Getting ConnectWise Manage Locations."
    $Sites = Get-CWM-CompanySites -ParentId $CWMCompany.id
    $CWMLocations = $Sites.name

    Write-MessageLog "Getting board information."
    $CWMBoard = Get-CWM-Board -ServiceBoardName $ServiceBoardName
    $CWMBoardStatus = if ($ServiceBoardStatusName -ne '') { Get-CWM-BoardStatus -ServiceBoardName $ServiceBoardName -ServiceBoardStatusName $ServiceBoardStatusName } else { $null }
    $CWMBoardType = if ($ServiceBoardTypeName -ne '') { Get-CWM-BoardType -ServiceBoardName $ServiceBoardName -ServiceBoardTypeName $ServiceBoardTypeName } else { $null }
    $CWMBoardSubType = if ($ServiceBoardSubTypeName -ne '') { Get-CWM-BoardSubType -ServiceBoardName $ServiceBoardName -ServiceBoardSubTypeName $ServiceBoardSubTypeName } else { $null }
    $CWMBoardItem = if ($ServiceBoardItemName -ne '') { Get-CWM-BoardItem -ServiceBoardName $ServiceBoardName -ServiceBoardTypeName $ServiceBoardTypeName -ServiceBoardSubTypeName $ServiceBoardSubTypeName -ServiceBoardItemName $ServiceBoardItemName } else { $null }

    Write-MessageLog "Getting employee manager."
    $ManagerDetails = Get-EH-Manager -ManagerId $Employee.primary_manager.id -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization
    if ($null -eq $ManagerDetails) { Write-ErrorLog "Failed to get manager details." ; return }
    $ManagerTicket = Set-ManagerDetails -ManagerDetails $ManagerDetails

    Write-MessageLog "Getting team details."
    $Teams = Get-EH-TeamDetails -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization | Where-Object { $_.status -contains "active" }
    if ($null -eq $Teams) { Write-ErrorLog "Failed to get team details." ; return }

    Write-MessageLog "Setting Persona."
    $TeamNames = Get-EH-EmployeeTeamNames -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization -EmployeeId $Employee.id -Teams $Teams
    $Persona = Set-Persona -TeamNames $TeamNames -Location $Location -AccessRoles $AccessRoles
    if ($null -eq $Persona) { Write-ErrorLog "Failed to set access level." ; return }

    Write-MessageLog "Retrieving Persona Data."
    $csvData = Get-MSGraph-SPCSV -SiteUrl $SPSiteUrl -SiteName $SPSiteName -FileName $SPFileName
    $personaData = Get-MSGraph-SPPersonaData -CSVFromSharePoint $csvData -Persona $Persona
    Write-MessageLog "Persona data processed successfully."

    Write-MessageLog "Setting department."
    $Department = Set-Department -Persona $Persona -AccessRoles $AccessRoles
    if ($null -eq $Department) { Write-ErrorLog "Failed to set department." ; return }

    Write-MessageLog "Setting location details."
    $Location = Get-Location -Location $Location -CWMLocations $CWMLocations

    Write-MessageLog "Setting role."
    $Role = Set-Role -Persona $Persona -AccessRoles $AccessRoles
    if ($null -eq $Role) { Write-ErrorLog "Failed to set role." ; return }

    Write-MessageLog "Setting computer and mobile requirements."
    $ComputerRequirementDetails = Set-ComputerRequirement -Role $Role -AccessRoles $AccessRoles
    $ComputerRequirement = $ComputerRequirementDetails[0]
    $ComputerIdentifier = $ComputerRequirementDetails[1]

    $MobileRequirementDetails = Set-MobileRequirement -Role $Role -AccessRoles $AccessRoles
    $MobileRequirement = $MobileRequirementDetails[0]
    $MobileIdentifier = $MobileRequirementDetails[1]

    Write-MessageLog "Setting additional notes."
    $AdditionalNotes = Set-AdditionalNotes -Role $Role

    Write-MessageLog "Setting ticket summary and initial note."
    $CWMTicketSummary = Set-TicketSummaryNewUser -FirstName $FirstName -LastName $LastName -StartDate $StartDate
    $CWMTicketNoteInitial = Set-TicketNoteInitialNewUser `
        -FirstName $FirstName `
        -LastName $LastName `
        -StartDate $StartDate `
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

    Write-MessageLog "Creating ticket."
    $Ticket = New-CWM-Ticket -CWMCompany $CWMCompany `
        -CWMContact $CWMContact `
        -CWMBoard $CWMBoard `
        -CWMTicketSummary $CWMTicketSummary `
        -CWMTicketNoteInitial $CWMTicketNoteInitial `
        -CWMBoardStatus $CWMBoardStatus `
        -CWMBoardType $CWMBoardType `
        -CWMBoardSubType $CWMBoardSubType `
        -CWMBoardItem $CWMBoardItem

    $UserDetails = @{
        TicketID            = $Ticket.id
        TicketSummary       = $CWMTicketSummary
        StartDate           = (Get-Date $Employee.start_date).ToString("yyyy-MM-dd")
        DisplayName         = "$FirstName $LastName"
        GivenName           = $FirstName
        Surname             = $LastName
        JobTitle            = $JobTitle
        Department          = $Department
        ManagerName         = "$($ManagerDetails.first_name) $($ManagerDetails.last_name)"
        ManagerEmail        = $ManagerDetails.company_email
        Location            = $Location
        LicenseGroup        = $personaData.LicenseGroup
        License             = Set-License -LicenseGroup $personaData.LicenseGroup
        GroupNames          = $personaData.M365Group
        SharedMailbox       = $personaData.SharedMailbox
        DistributionLists   = $personaData.DistributionList
        ComputerRequirement = $ComputerRequirement
        ComputerNote        = $ComputerIdentifier
        MobileRequirement   = $MobileRequirement
        MobileNote          = $MobileIdentifier
        NonSupportUser      = Set-NonSupportUser -Department $Department
        CustomTasks         = if (-not [string]::IsNullOrWhiteSpace($AdditionalNotes)) { ($AdditionalNotes -split "`n" | ForEach-Object { ";Manual Task:$_" }) -join "" }
    }
    Write-MessageLog "Parsed employee details: $($UserDetails | Out-String)"

    $WebhookPayload = $UserDetails | ConvertTo-Json -Depth 10
    $WebhookUrl = "https://your.webhook.url"
    Write-MessageLog -Message "Sending webhook payload: $WebhookPayload"
    
    Invoke-RestMethod -Uri $WebhookUrl -Method Post -ContentType "application/json" -Body $WebhookPayload
    Write-MessageLog -Message "Webhook payload sent successfully to $WebhookUrl"

    Write-MessageLog "New-EmployeeCreated process completed."
}

function New-EmployeeOnboarded {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$PreCheckData
    )

    Write-MessageLog "Getting employee details."
    $Employee = Get-EH-Employee -EmployeeId $PreCheckData.EmployeeData.id -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization
    if ($Employee -contains "error") { Write-ErrorLog "Failed to fetch employee details." ; return }

    $FirstName = $Employee.first_name
    $LastName = $Employee.last_name
    $MobileNumber = $Employee.personal_mobile_number
    $parsedDate = $null
    if (Test-DateValid -dateString $Employee.start_date -parsedDate ([ref]$parsedDate)) {
        $StartDate = Set-EmployeeDate -inputDate $Employee.start_date -type "start"
    }
    else {
        Write-ErrorLog "Invalid date format for start date."
    }

    $CWMTicketSummary = Set-TicketSummaryNewUser -FirstName $FirstName -LastName $LastName -StartDate $StartDate
    $Condition = "summary='$CWMTicketSummary'"
    $Ticket = Get-CWMTicket -condition $Condition

    $UserDetails = @{
        TicketID            = $Ticket.id
        TicketSummary       = $CWMTicketSummary
        StartDate           = (Get-Date $Employee.start_date).ToString("yyyy-MM-dd")
        DisplayName         = "$FirstName $LastName"
        GivenName           = $FirstName
        Surname             = $LastName
        MobileNumber        = $MobileNumber
        ManagerName         = "$($ManagerDetails.first_name) $($ManagerDetails.last_name)"
        ManagerEmail        = $ManagerDetails.company_email
    }
    Write-MessageLog "Parsed employee details: $($UserDetails | Out-String)"

    $WebhookPayload = $UserDetails | ConvertTo-Json -Depth 10
    $WebhookUrl = "https://your.webhook.url"
    Write-MessageLog -Message "Sending webhook payload: $WebhookPayload"
    
    Invoke-RestMethod -Uri $WebhookUrl -Method Post -ContentType "application/json" -Body $WebhookPayload
    Write-MessageLog -Message "Webhook payload sent successfully to $WebhookUrl"

    Write-MessageLog "New-EmployeeOnboarded process completed."
}

function New-EmployeeOffboarding {
    Write-MessageLog "Starting New-EmployeeOffboarding process."

    $ClientIdentifier = "SEA"
    $ContactName = "System Notifications"
    $ServiceBoardName = "Pia"
    $ServiceBoardStatusName = "New (portal)~"
    $ServiceBoardTypeName = "Request"
    $ServiceBoardSubTypeName = "Profile"
    $ServiceBoardItemName = "Terminate User"

    Write-MessageLog "Getting employee details."
    $Employee = Get-EH-Employee -EmployeeId $PreCheckData.EmployeeData.id -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization
    if ($Employee -contains "error") { Write-ErrorLog "Failed to fetch employee details." ; return }

    $FirstName = $Employee.first_name
    $LastName = $Employee.last_name

    Write-MessageLog "Setting employee end date."
    $parsedDate = $null
    if (Test-DateValid -dateString $Employee.termination_date -parsedDate ([ref]$parsedDate)) {
        $EndDate = Set-EmployeeDate -inputDate $Employee.termination_date -type "end"
    }
    else {
        Write-ErrorLog "Invalid date format for end date."
    }

    Write-MessageLog "Getting company and contact information."
    $CWMCompany = Get-CWM-Company -ClientIdentifier $ClientIdentifier
    if ($null -eq $CWMCompany) { 
        Write-ErrorLog "Failed to get company information." 
        return 
    }

    $CWMContact = Get-CWM-Contact -ClientIdentifier $ClientIdentifier -ContactName $ContactName
    if ($null -eq $CWMContact) { 
        Write-ErrorLog "Failed to get contact information." 
        return 
    }

    Write-MessageLog "Getting board information."
    $CWMBoard = Get-CWM-Board -ServiceBoardName $ServiceBoardName
    $CWMBoardStatus = if ($ServiceBoardStatusName -ne '') { Get-CWM-BoardStatus -ServiceBoardName $ServiceBoardName -ServiceBoardStatusName $ServiceBoardStatusName } else { $null }
    $CWMBoardType = if ($ServiceBoardTypeName -ne '') { Get-CWM-BoardType -ServiceBoardName $ServiceBoardName -ServiceBoardTypeName $ServiceBoardTypeName } else { $null }
    $CWMBoardSubType = if ($ServiceBoardSubTypeName -ne '') { Get-CWM-BoardSubType -ServiceBoardName $ServiceBoardName -ServiceBoardSubTypeName $ServiceBoardSubTypeName } else { $null }
    $CWMBoardItem = if ($ServiceBoardItemName -ne '') { Get-CWM-BoardItem -ServiceBoardName $ServiceBoardName -ServiceBoardTypeName $ServiceBoardTypeName -ServiceBoardSubTypeName $ServiceBoardSubTypeName -ServiceBoardItemName $ServiceBoardItemName } else { $null }

    Write-MessageLog "Setting employee information."
    $EmployeeTicket = Set-EmployeeDetails -Employee $Employee
    if ($null -eq $EmployeeTicket) { Write-ErrorLog "Failed to set employee details." ; return }

    Write-MessageLog "Getting manager details."
    $ManagerDetails = Get-EH-Manager -ManagerId $Employee.primary_manager.id -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization
    if ($null -eq $ManagerDetails) { Write-ErrorLog "Failed to get manager details." ; return }

    Write-MessageLog "Getting team details."
    $Teams = Get-EH-TeamDetails -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization | Where-Object { $_.status -contains "active" }
    if ($null -eq $Teams) { Write-ErrorLog "Failed to get team details." ; return }

    Write-MessageLog "Getting employee team names."
    $TeamNames = Get-EH-EmployeeTeamNames -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization -EmployeeId $Employee.id -Teams $Teams
    if ($null -eq $TeamNames) { Write-ErrorLog "Failed to get employee team names." ; return }

    Write-MessageLog "Setting Parsona."
    $Persona = Set-Persona -TeamNames $TeamNames -Location $Location -AccessRoles $AccessRoles
    if ($null -eq $Persona) { Write-ErrorLog "Failed to set access level." ; return }

    Write-MessageLog "Setting role."
    $Role = Set-Role -Persona $Persona -AccessRoles $AccessRoles
    if ($null -eq $Role) { Write-ErrorLog "Failed to set role." ; return }

    if ($Role -eq "CARE MANAGER" -or $Role -eq "CARE SERVICES SCHEDULER" -or $Role -eq "CHEF" -or $Role -eq "COMMUNITY MANAGEMENT" -or $Role -eq "LIFESTYLE" -or $Role -eq "MAINTENANCE" -or $Role -eq "SALES" -or $Role -eq "REGISTERED NURSE") {
        Write-MessageLog "Setting manager information."
        $ManagerTicket = Set-ManagerDetails -ManagerDetails $ManagerDetails
    }

    Write-MessageLog "Setting ticket summary and initial note."
    $CWMTicketSummary = Set-TicketSummaryExitUser -FirstName $FirstName -LastName $LastName -EndDate $EndDate
    $CWMTicketNoteInitial = Set-TicketNoteInitialExitUser `
        -EmployeeTicket $EmployeeTicket `
        -EndDate $EndDate `
        -ManagerTicket $ManagerTicket

    Write-MessageLog "Creating ticket."
    $Ticket = New-CWM-Ticket -CWMCompany $CWMCompany `
        -CWMContact $CWMContact `
        -CWMBoard $CWMBoard `
        -CWMTicketSummary $CWMTicketSummary `
        -CWMTicketNoteInitial $CWMTicketNoteInitial `
        -CWMBoardStatus $CWMBoardStatus `
        -CWMBoardType $CWMBoardType `
        -CWMBoardSubType $CWMBoardSubType `
        -CWMBoardItem $CWMBoardItem

        $UserDetails = @{
            TicketID            = $Ticket.id
            TicketSummary       = $CWMTicketSummary
            EndDate             = (Get-Date $Employee.termination_date).ToString("ddd dd MMM, yyyy HH:mm")
            GivenName           = $FirstName
            Surname             = $LastName
            ManagerName         = "$($ManagerDetails.first_name) $($ManagerDetails.last_name)"
            ManagerTicket       = $ManagerTicket
        }
        Write-MessageLog "Parsed employee details: $($UserDetails | Out-String)"
    
        $WebhookPayload = $UserDetails | ConvertTo-Json -Depth 10
        $WebhookUrl = "https://your.webhook.url"
        Write-MessageLog -Message "Sending webhook payload: $WebhookPayload"
        
        Invoke-RestMethod -Uri $WebhookUrl -Method Post -ContentType "application/json" -Body $WebhookPayload
        Write-MessageLog -Message "Webhook payload sent successfully to $WebhookUrl"

    Write-MessageLog "New-EmployeeOffboarding process completed."
}

ProcessMainFunction