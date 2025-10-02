
. C:\Scripts\Repositories\Private\DEV\Functions\Function-ALL-General.ps1
. C:\Scripts\Repositories\Private\DEV\Functions\Function-ALL-MSGraph.ps1
. C:\Scripts\Repositories\Private\DEV\Functions\Function-ALL-MSGraph-SP.ps1
. C:\Scripts\Repositories\Private\DEV\Functions\Function-ALL-CWM.ps1
Write-MessageLog -Message "Dependencies loaded successfully."

function Initialize {
    Write-MessageLog -Message "Initializing script configuration."
    $Global:CWMClientIdName = "MIT-CWMApi-ClientId"
    $Global:CWMPublicKeyName = "MIT-CWMApi-PubKey"
    $Global:CWMPrivateKeyName = "MIT-CWMApi-PrivateKey"
    $Global:CWMCompanyIdName = "MIT-CWMApi-CompanyId"
    $Global:CWMUrlName = "MIT-CWMApi-Url"
    $Global:ClientIdSecretName = "MIT-AutomationApp-ClientID"
    $Global:ClientSecretSecretName = "MIT-AutomationApp-ClientSecret"
    # try {
    #     Write-MessageLog -Message "Retrieving Azure Key Vault name from automation variables."
    #     $Global:AzKeyVaultName = Get-AutomationVariable -Name 'AzKeyVaultName'
    
    #     if ($null -eq $AzKeyVaultName) {
    #         Write-ErrorLog -Message "Azure Key Vault name is missing. Check your automation variables."
    #         throw "Azure Key Vault name is not defined."
    #     }
    
    #     Write-MessageLog -Message "Azure Key Vault name retrieved successfully: $AzKeyVaultName."
    # } catch {
    #     Write-ErrorLog -Message "Failed to retrieve Azure Key Vault name. Error: $_"
    #     throw $_
    # }
    $Global:TenantUrl = 'manganoit.com.au'
    $Global:SiteUrl = "https://manganoit.sharepoint.com/Clients"
    $Global:SiteName = "Clients"
    $Global:ListName = "Afterhours Support Agreement" 
    Write-MessageLog -Message "Script configuration initialized."
}

function ProcessSPFilteredList {
    param (
        $SPFilteredList
    )

    Write-MessageLog -Message "Starting to process filtered list items."
    foreach ($Item in $SPFilteredList) {
        if ($null -ne $Item.Priority) {
            Write-MessageLog -Message "Processing company: $($Item.CompanyID) - $($Item.CompanyName) with priority items."
            $currentDate = Get-Date
            $StartDate = $currentDate.AddDays(-7)
            $formattedDate = $StartDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
            $Conditions = "company/identifier='$($Item.CompanyID)' AND workType/name='3 - After Hours' AND agreement/type='$($Item.AgreementType)' AND lastUpdated >= [$formattedDate]"
            Write-MessageLog -Message "Constructed conditions for TimeEntry retrieval: $Conditions."
            UpdateTimeEntries -Conditions $Conditions -Item $Item
        } else {
            Write-MessageLog -Message "No priority found for company: $($Item.CompanyName). Skipping."
        }
    }
}

function UpdateTimeEntries {
    param (
        [string]$Conditions,
        $Item
    )

    Write-MessageLog -Message "Fetching TimeEntries using conditions: $Conditions."
    $TimeEntries = Get-CWMTimeEntry -condition $Conditions -all
    if ($null -eq $TimeEntries -or $TimeEntries.Count -eq 0) {
        Write-MessageLog -Message "No TimeEntries found for conditions: $Conditions."
        return
    }
    Write-MessageLog -Message "Retrieved $($TimeEntries.Count) TimeEntries. Processing each entry."
    ProcessTimeEntries -TimeEntries $TimeEntries -Item $Item
}

function ProcessTimeEntries {
    param (
        $TimeEntries,
        $Item
    )

    Write-MessageLog -Message "Starting to process each TimeEntry."
    foreach ($TimeEntry in $TimeEntries) {
        Write-MessageLog -Message "Processing TimeEntry ID: $($TimeEntry.id). Retrieving associated ticket."
        $Ticket = Get-CWMTicket -id $TimeEntry.chargeToId
        if ($null -eq $Ticket) {
            Write-MessageLog -Message "Ticket not found for TimeEntry ID: $($TimeEntry.id). Skipping."
            continue
        }
        Write-MessageLog -Message "Ticket ID: $($Ticket.id) retrieved. Processing VIP support status."
        ProcessVIPContact -Ticket $Ticket -Item $Item -TimeEntry $TimeEntry
    }
    Write-MessageLog -Message "All TimeEntries processed successfully."
}

function ProcessVIPContact {
    param (
        $Ticket,
        $Item,
        $TimeEntry
    )

    $IsVIP = $False

    if ($Item.VIPContact -ne $False) {
        Write-MessageLog -Message "VIP support. Retrieving contact type for company: $($Item.CompanyName)."
        $ContactType = (Get-CWMContactType -conditions "company/identifier='$($Item.CompanyID)'" | Where-Object { $_.description -match "VIP" }).description
        $Contact = Get-CWMContact -id $Ticket.contact.id
        if ($null -eq $Contact) {
            Write-MessageLog -Message "Contact not found for Ticket ID: $($Ticket.id). Skipping VIP support processing."
        } else {
            Write-MessageLog -Message "Contact ID: $($Contact.id) retrieved. Checking VIP conditions."
            if ($Contact.types.name -contains $ContactType) {
                Write-MessageLog -Message "VIP Contact match found for TimeEntry ID: $($TimeEntry.id)."
                $IsVIP = $True
            }
        }
    }

    if ($IsVIP -or ($Item.Priority -contains $Ticket.priority.name)) {
        Write-MessageLog -Message "AH Condition met for TimeEntry ID: $($TimeEntry.id). Setting 'DoNotBill'."
        UpdateTimeEntry -Ticket $Ticket -TimeEntry $TimeEntry -BillableOption "DoNotBill"
    } else {
        Write-MessageLog -Message "AH conditions NOT met for TimeEntry ID: $($TimeEntry.id). Setting 'Billable'."
        UpdateTimeEntry -Ticket $Ticket -TimeEntry $TimeEntry -BillableOption "Billable"
    }
}

function UpdateTimeEntry {
    param (
        $Ticket,
        $TimeEntry,
        [string]$BillableOption
    )

    Write-MessageLog -Message "Checking billable option for TimeEntry ID: $($TimeEntry.id)."
    
    if ($TimeEntry.billableOption -ne $BillableOption) {
        Write-MessageLog -Message "Updating billable option to '$BillableOption' for TimeEntry ID: $($TimeEntry.id)."
        Update-CWMTimeEntry -id $TimeEntry.id -Operation replace -Path billableOption -Value $BillableOption
        Write-MessageLog -Message "TimeEntry ID: $($TimeEntry.id) updated successfully."
        try {
            $TicketNoteUpdate = "TimeEntry ID: $($TimeEntry.id) Updated billable option to '$BillableOption'."
            New-CWMTicketNote -ticketId $Ticket.id -text $TicketNoteUpdate -internalAnalysisFlag $true
            Write-MessageLog -Message "Note added to Ticket ID: $($Ticket.id) successfully."
        } catch {
            Write-ErrorLog -Message "Failed to add note to Ticket ID: $($Ticket.id). Error: $_"
        }
    } else {
        Write-MessageLog -Message "Billable option for TimeEntry ID: $($TimeEntry.id) is already set to '$BillableOption'. No update needed."
    }
}

function ProcessMainFunction {
    Write-MessageLog -Message "Starting the Main Process."
    
    try {
        Initialize
    } catch {
        Write-ErrorLog -Message "Failed to initialize script configuration. Error: $_"
        return
    }

    try {
        Write-MessageLog -Message "Connecting to Microsoft Graph."
        ConnectMSGraphTest -TenantUrl $TenantUrl
        #ConnectMSGraph -AzKeyVaultName $AzKeyVaultName -TenantUrl $TenantUrl -ClientIdSecretName $ClientIdSecretName -ClientSecretSecretName $ClientSecretSecretName
        Write-MessageLog -Message "Connected to Microsoft Graph successfully."
    } catch {
        Write-ErrorLog -Message "Failed to connect to Microsoft Graph. Error: $_"
        return
    }

    try {
        Write-MessageLog -Message "Connecting to ConnectWise Manage."
        ConnectCWMTest
        #ConnectCWM -AzKeyVaultName $AzKeyVaultName -CWMClientIdName $CWMClientIdName -CWMPublicKeyName $CWMPublicKeyName -CWMPrivateKeyName $CWMPrivateKeyName -CWMCompanyIdName $CWMCompanyIdName -CWMUrlName $CWMUrlName
        Write-MessageLog -Message "Connected to ConnectWise Manage successfully."
    } catch {
        Write-ErrorLog -Message "Failed to connect to ConnectWise Manage. Error: $_"
        return
    }

    try {
        Write-MessageLog -Message "Retrieving SharePoint site: $SiteName at URL: $SiteUrl."
        $Site = Get-MSGraph-SPSite -SiteUrl $SiteUrl -SiteName $SiteName
        if ($null -eq $Site) {
            Write-MessageLog -Message "SharePoint site not found. Exiting process." -Level "ERROR"
            return
        }
    } catch {
        Write-ErrorLog -Message "Failed to retrieve SharePoint site. Error: $_"
        return
    }

    try {
        Write-MessageLog -Message "Retrieving SharePoint list: $ListName from site: $SiteName."
        $List = Get-MSGraph-SPList -Site $Site -ListName $ListName
        if ($null -eq $List) {
            Write-MessageLog -Message "SharePoint list not found. Exiting process." -Level "ERROR"
            return
        }
    } catch {
        Write-ErrorLog -Message "Failed to retrieve SharePoint list. Error: $_"
        return
    }

    try {
        Write-MessageLog -Message "Retrieving filtered list items from SharePoint list: $ListName."
        $SPFilteredList = Get-MSGraph-SPFilteredList -Site $Site -List $List
        Write-MessageLog -Message "Filtered list items retrieved. Beginning processing."

        ProcessSPFilteredList -SPFilteredList $SPFilteredList
    } catch {
        Write-ErrorLog -Message "Failed to retrieve or process filtered list items. Error: $_"
        return
    }

    try {
        Write-MessageLog -Message "Disconnecting from ConnectWise Manage."
        Disconnect-CWM
        Write-MessageLog -Message "Disconnected from ConnectWise Manage successfully."
    } catch {
        Write-ErrorLog -Message "Failed to disconnect from ConnectWise Manage. Error: $_"
    }

    Write-MessageLog -Message "Main Process completed successfully."
}

ProcessMainFunction