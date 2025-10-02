# Load Dependencies
. .\Function-ALL-General.ps1
. .\Function-ALL-MSGraph.ps1
Write-MessageLog -Message "Dependencies loaded successfully."

# Main Execution
function ProcessMainFunction {
    Write-MessageLog -Message "Starting the main process."

    try {
        Initialize
        Write-MessageLog -Message "Initialization completed successfully."
    }
    catch {
        Write-ErrorLog -Message "Failed during initialization. Error: $_"
        return
    }

    try {
        Write-MessageLog -Message "Processing date range for user: $TeamUserIdentifier."
        ProcessDateRange
        Write-MessageLog -Message "Date range processed successfully for user: $TeamUserIdentifier."
    }
    catch {
        Write-ErrorLog -Message "Failed to process date range for user: $TeamUserIdentifier. Error: $_"
        continue
    }

    try {
        Write-MessageLog -Message "Connecting to Microsoft Graph."
        # ConnectMSGraphTest -TenantUrl $TenantUrl
        ConnectMSGraph -AzKeyVaultName $AzKeyVaultName -TenantUrl $TenantUrl -ClientIdSecretName $ClientIdSecretName -ClientSecretSecretName $ClientSecretSecretName
        Write-MessageLog -Message "Connected to Microsoft Graph successfully."
    }
    catch {
        Write-ErrorLog -Message "Failed to connect to Microsoft Graph. Error: $_"
        return
    }

    $Global:TeamOwner = Get-MgUser -UserId "workflows@manganoit.com.au"
    $Global:Team = Get-MgTeam -Filter "displayName eq 'Team Mangano'"
    $Global:TeamUserIdentifiers = Get-MgGroupMember -GroupId $Team.Id -All | Where-Object { $_.AdditionalProperties.'@odata.type' -eq "#microsoft.graph.user" } | ForEach-Object { $_.AdditionalProperties.userPrincipalName }

    foreach ($TeamUserIdentifier in $TeamUserIdentifiers) {
        Write-MessageLog -Message "Starting process for user: $TeamUserIdentifier."

        try {
            Write-MessageLog -Message "Retrieving user and team information for user: $TeamUserIdentifier."
            GetUserTeamInfo
            Write-MessageLog -Message "User and team information retrieved successfully for user: $TeamUserIdentifier."
        }
        catch {
            Write-ErrorLog -Message "Failed to retrieve user and team information for user: $TeamUserIdentifier. Error: $_"
            continue
        }

        try {
            Write-MessageLog -Message "Deleting existing calendar events for user: $TeamUserIdentifier."
            DeleteExistingEvents
            Write-MessageLog -Message "Existing calendar events deleted successfully for user: $TeamUserIdentifier."
        }
        catch {
            Write-ErrorLog -Message "Failed to delete existing calendar events for user: $TeamUserIdentifier. Error: $_"
            continue
        }

        try {
            Write-MessageLog -Message "Retrieving and filtering shifts for user: $TeamUserIdentifier."
            GetFilteredShifts
            Write-MessageLog -Message "Shifts retrieved and filtered successfully for user: $TeamUserIdentifier."
        }
        catch {
            Write-ErrorLog -Message "Failed to retrieve and filter shifts for user: $TeamUserIdentifier. Error: $_"
            continue
        }

        try {
            Write-MessageLog -Message "Creating new calendar events for user: $TeamUserIdentifier."
            CreateNewEvents
            Write-MessageLog -Message "New calendar events created successfully for user: $TeamUserIdentifier."
        }
        catch {
            Write-ErrorLog -Message "Failed to create new calendar events for user: $TeamUserIdentifier. Error: $_"
            continue
        }

        Write-MessageLog -Message "Completed processing for user: $TeamUserIdentifier."
    }

    Write-MessageLog -Message "Main process completed successfully."
}

# Configuration and Initialization
function Initialize {
    Write-MessageLog "Initializing script configuration."
    $Global:TenantUrl = 'manganoit.com.au'
    $Global:ClientIdSecretName = "MIT-AutomationApp-ClientID"
    $Global:ClientSecretSecretName = "MIT-AutomationApp-ClientSecret"
    try {
        Write-MessageLog -Message "Retrieving Azure Key Vault name from automation variables."
        $Global:AzKeyVaultName = Get-AutomationVariable -Name 'AzKeyVaultName'
    
        if ($null -eq $AzKeyVaultName) {
            Write-ErrorLog -Message "Azure Key Vault name is missing. Check your automation variables."
            throw "Azure Key Vault name is not defined."
        }
    
        Write-MessageLog -Message "Azure Key Vault name retrieved successfully: $AzKeyVaultName."
    } catch {
        Write-ErrorLog -Message "Failed to retrieve Azure Key Vault name. Error: $_"
        throw $_
    }
    $Global:BrisbaneTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("E. Australia Standard Time")
    Write-MessageLog "Script configuration initialized."
}

# Calculate Date Range in Brisbane Time
function ProcessDateRange {
    Write-MessageLog -Message "Calculating date range in Brisbane time."
    $Global:TodayBrisbane = [System.TimeZoneInfo]::ConvertTimeFromUtc((Get-Date).ToUniversalTime(), $BrisbaneTimeZone).Date
    $Global:EndDateBrisbane = $TodayBrisbane.AddMonths(2)
    Write-MessageLog -Message "Date range: Start - $TodayBrisbane, End - $EndDateBrisbane."

    # Convert to UTC for filtering
    $Global:StartFromDateUTC = [System.TimeZoneInfo]::ConvertTimeToUtc($TodayBrisbane, $BrisbaneTimeZone)
    $Global:EndDateUTC = [System.TimeZoneInfo]::ConvertTimeToUtc($EndDateBrisbane, $BrisbaneTimeZone)
    $Global:StartFromDateIso = $StartFromDateUTC.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    $Global:EndDateIso = $EndDateUTC.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    Write-MessageLog -Message "Converted date range in UTC: Start - $StartFromDateIso, End - $EndDateIso."
}

# Retrieve User and Team Information
function GetUserTeamInfo {
    Write-MessageLog -Message "Retrieving user and team owner information."
    $Global:User = Get-MgUser -UserId $TeamUserIdentifier
    Write-MessageLog -Message "Retrieved user: $($User.DisplayName), team owner: $($TeamOwner.DisplayName), and Team ID: $($Team.Id)."
}

# Delete Existing Calendar Events
function DeleteExistingEvents {
    Write-MessageLog -Message "Retrieving calendar events with 'Silver category' to delete."
    $Filter = "start/dateTime ge '$StartFromDateIso'"
    $EventsToDelete = Get-MgUserEvent -UserId $User.id -Filter $Filter -All | Where-Object {
        $_.Categories -contains "Silver category"
    }

    if ($EventsToDelete.Count -gt 0) {
        Write-MessageLog -Message "Found $($EventsToDelete.Count) events to delete."
        foreach ($Event in $EventsToDelete) {
            try {
                Remove-MgUserEvent -UserId $User.id -EventId $Event.Id -Confirm:$false
                Write-MessageLog -Message "Deleted event: $($Event.Subject), Event ID: $($Event.Id)."
            }
            catch {
                Write-MessageLog -Message "Error deleting event: $($Event.Subject). Error: $_"
            }
        }
    }
    else {
        Write-MessageLog -Message "No events with 'Silver category' found to delete."
    }
}

# Retrieve and Filter Shifts
function GetFilteredShifts {
    Write-MessageLog -Message "Retrieving shifts from Microsoft Teams."
    $Filter = "sharedShift/startDateTime ge $StartFromDateIso and sharedShift/endDateTime le $EndDateIso"
    $Headers = @{ "MS-APP-ACTS-AS" = $TeamOwner.id }
    $Global:AllShifts = Get-MgTeamScheduleShift -TeamId $Team.id -Headers $Headers -Filter $Filter
    Write-MessageLog -Message "Retrieved $(if ($AllShifts) { $AllShifts.Count } else { 0 }) shifts."

    Write-MessageLog -Message "Filtering shifts for user: $($User.UserPrincipalName)."
    $Global:FilteredShifts = $AllShifts | Where-Object { $_.userId -eq $User.id }
    Write-MessageLog -Message "Filtered shifts count: $(if ($FilteredShifts) { $FilteredShifts.Count } else { 0 })."
}

# Create New Calendar Events
function CreateNewEvents {
    Write-MessageLog -Message "Processing filtered shifts to create new calendar events."
    foreach ($Shift in $FilteredShifts) {
        Write-MessageLog -Message "Processing shift: $($Shift.sharedShift.displayName)."

        $NewEvent = @{
            subject         = $Shift.sharedShift.displayName
            body            = @{
                contentType = "HTML"
                content     = $Shift.sharedShift.notes
            }
            start           = @{
                dateTime = ([System.TimeZoneInfo]::ConvertTimeFromUtc(([DateTime]::Parse($Shift.sharedShift.startDateTime)), $BrisbaneTimeZone)).ToString("yyyy-MM-ddTHH:mm:ss")
                timeZone = "E. Australia Standard Time"
            }
            end             = @{
                dateTime = ([System.TimeZoneInfo]::ConvertTimeFromUtc(([DateTime]::Parse($Shift.sharedShift.endDateTime)), $BrisbaneTimeZone)).ToString("yyyy-MM-ddTHH:mm:ss")
                timeZone = "E. Australia Standard Time"
            }
            isReminderOn    = $false
            Categories      = @("Silver category")
            attendees       = @(
                @{
                    emailAddress = @{
                        address = $User.UserPrincipalName
                        name    = $User.displayName
                    }
                    type         = "Required"
                }
            )
            isOnlineMeeting = $false
            showAs          = "Free"
        }

        try {
            $CreatedEvent = New-MgUserEvent -UserId $User.id -BodyParameter $NewEvent
            Write-MessageLog -Message "Event created: $($Shift.sharedShift.displayName), ID: $($CreatedEvent.Id)."
        }
        catch {
            Write-MessageLog -Message "Error creating event for shift: $($Shift.sharedShift.displayName). Error: $_"
        }
    }
}

# Run Main Process
ProcessMainFunction