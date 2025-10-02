# Load Dependencies
. .\Function-ALL-General.ps1
. .\Function-ALL-CWM.ps1
Write-MessageLog -Message "Dependencies loaded successfully."

# Main Execution
function ProcessMainFunction {
    Write-MessageLog -Message "Starting the Main Process."
    
    try {
        # Initialize the script configuration
        Initialize
        Write-MessageLog -Message "Script configuration initialized successfully."
    } catch {
        Write-ErrorLog -Message "Failed to initialize script configuration. Error: $_"
        return
    }

    try {
        Write-MessageLog -Message "Connecting to ConnectWise Manage."
        # Connect to ConnectWise Manage
        ConnectCWMTest
        #ConnectCWM -AzKeyVaultName $AzKeyVaultName -CWMClientIdName $CWMClientIdName -CWMPublicKeyName $CWMPublicKeyName -CWMPrivateKeyName $CWMPrivateKeyName -CWMCompanyIdName $CWMCompanyIdName -CWMUrlName $CWMUrlName
        Write-MessageLog -Message "Connected to ConnectWise Manage successfully."
    } catch {
        Write-ErrorLog -Message "Failed to connect to ConnectWise Manage. Error: $_"
        return
    }

    try {
        # Process the update data
        Write-MessageLog -Message "Starting data processing."

        $TotalItems = $UpdateData.Count
        $CurrentItem = 0

        foreach ($Data in $UpdateData) {
            try {
                $CurrentItem++
                $PercentComplete = [math]::Round(($CurrentItem / $TotalItems) * 100)
                Write-Progress -Activity "Processing Data" -Status "Processing item $CurrentItem of $TotalItems" -PercentComplete $PercentComplete
                Invoke-DataProcessing -Data $Data
                Write-MessageLog -Message "Processed item $CurrentItem of $TotalItems successfully."
            } catch {
                Write-ErrorLog -Message "Failed to process item $CurrentItem of $TotalItems. Error: $_"
            }
        }

        Write-MessageLog -Message "All configurations processed successfully."
    } catch {
        Write-ErrorLog -Message "Error occurred during data processing. Error: $_"
        return
    }

    try {
        # Disconnect from ConnectWise Manage
        Write-MessageLog -Message "Disconnecting from ConnectWise Manage."
        Disconnect-CWM
        Write-MessageLog -Message "Disconnected from ConnectWise Manage successfully."
    } catch {
        Write-ErrorLog -Message "Failed to disconnect from ConnectWise Manage. Error: $_"
    }

    Write-MessageLog -Message "Main Process completed successfully."
}

function Initialize {
    Write-MessageLog -Message "Initializing script configuration."
    $Global:UpdateData = Import-Csv -Path "C:\Scripts\InputData\CSV\20241119.csv"
    $Global:LogPath = 'C:\Scripts\OutputLogs\'
    Write-MessageLog -Message "Script configuration initialized."
}

function Test-DateValid {
    param (
        [string]$dateString,
        [ref]$parsedDate
    )
    $formats = @("dd/MM/yyyy", "dd/M/yyyy", "d/MM/yyyy", "d/M/yyyy")
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

function Get-Contact {
    param (
        $Company,
        $ContactName
    )
    $Names = $ContactName -split ' '

    if ($Names.Count -eq 1) {
        $FirstName = '"' + $Names[0] + '"'
        $ContactCondition = "company/id=$($Company.id) AND firstName=$FirstName"
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
    $ContactCondition = "company/id=$($Company.id) AND firstName=$FirstName AND lastName=$LastName"
    $Contact = Get-CWMCompanyContact -Condition $ContactCondition
    if ($null -eq $Contact -and $Names.Count -gt 2) {
        $FirstName = '"' + $Names[0..1] + '"' -join ' '
        $LastName = '"' + ($Names[2..($Names.Count - 1)] -join ' ') + '"'
        $ContactCondition = "company/id=$($Company.id) AND firstName=$FirstName AND lastName=$LastName"
        $Contact = Get-CWMCompanyContact -Condition $ContactCondition
    }
    if ($null -eq $Contact) {
        Write-MessageLog -Message "Contact not found for company: $($Company.name), contact name: $ContactName" -Level "ERROR"
        return "ERROR"
    }
    return $Contact
}

function Get-Entities {
    param (
        $Data
    )

    Write-MessageLog -Message "Retrieving entities for Config ID: $($Data.Config_RecID), Name: $($Data.'Configuration Name')."

    $Company = $null
    $Config = $null
    $Vendor = $null
    $Site = $null
    $Manufacturer = $null
    $Contact = $null
    $Type = $null

    try {
        $Config = Get-CWMCompanyConfiguration -id $Data.Config_RecID
        if ($($Data.Company)) { $Company = Get-CWMCompany -condition "name='$($Data.Company)'" }
        if ($($Data.Vendor)) { $Vendor = Get-CWMCompany -condition "name='$($Data.Vendor)'" }
        if ($($Data.'Site Name')) { $SiteCondition = "name='$($Data.'Site Name')'" }
        if ($SiteCondition) { $Site = Get-CWMCompanySite -parentId $Company.id -condition $SiteCondition }
        if ($($Data.Manufacturer)) { $Manufacturer = Get-CWMManufacturer -Condition "name='$($Data.Manufacturer)'" }
        if ($Data.Contact) { $Contact = Get-Contact -Company $Company -ContactName $Data.Contact }
        if ($($Data.'Configuration Type')) { $Type = Get-CWMCompanyConfigurationType -Condition "name='$($Data.'Configuration Type')'" }
    } catch {
        Write-MessageLog -Message "Error retrieving entities for Config ID: $($Data.Config_RecID), Name: $($Data.'Configuration Name'). Error: $_" -Level "ERROR"
        return $null
    }

    $result = @{
        Company = $Company
        Config = $Config
        Vendor = $Vendor
        Site = $Site
        Manufacturer = $Manufacturer
        Contact = $Contact
        Type = $Type
    }

    Write-MessageLog -Message "Entities retrieved successfully for Config ID: $($Data.Config_RecID), Name: $($Data.'Configuration Name')."
    return $result
}

function Invoke-DataProcessing {
    param (
        $Data
    )

    Write-MessageLog -Message "Processing Config ID: $($Data.Config_RecID), Name: $($Data.'Configuration Name')."

    $Entities = Get-Entities -Data $Data
    if (-not $Entities) {
        Write-MessageLog -Message "Skipping Config ID: $($Data.Config_RecID) due to missing entities." -Level "ERROR"
        return
    }

    $Company = $Entities.Company
    $Config = $Entities.Config
    $Vendor = $Entities.Vendor
    $Site = $Entities.Site
    $Manufacturer = $Entities.Manufacturer
    $Contact = $Entities.Contact
    $Type = $Entities.Type

    if ($Contact -match "ERROR") {
        Write-MessageLog -Message "Contact not found for Config ID: $($Data.Config_RecID), Name: $($Data.'Configuration Name'). Skipping." -Level "ERROR"
        return
    }

    $parsedDate = $null
    if ($null -ne $Data.'Warranty Expiration Date' -and -not [string]::IsNullOrWhiteSpace($Data.'Warranty Expiration Date')) {
        if (Test-DateValid $Data.'Warranty Expiration Date' ([ref]$parsedDate)) {
            $WarrantyExpirationDate = $parsedDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
        } else {
            Write-MessageLog -Message "Invalid 'Warranty Expiration Date' for Config ID: $($Data.Config_RecID), Name: $($Data.'Configuration Name'). Skipping." -Level "ERROR"
            return
        }
    }

    $parsedDate = $null
    if ($null -ne $Data.'Date Last Sighted' -and -not [string]::IsNullOrWhiteSpace($Data.'Date Last Sighted')) {
        if (Test-DateValid $Data.'Date Last Sighted' ([ref]$parsedDate)) {
            $DateLastSighted = $parsedDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
        } else {
            Write-MessageLog -Message "Invalid 'Date Last Sighted' for Config ID: $($Data.Config_RecID), Name: $($Data.'Configuration Name'). Skipping." -Level "ERROR"
            return
        }
    }

    $parsedDate = $null
    if ($null -ne $Data.'Disposal Date' -and -not [string]::IsNullOrWhiteSpace($Data.'Disposal Date')) {
        if (Test-DateValid $Data.'Disposal Date' ([ref]$parsedDate)) {
            $DisposalDate = $parsedDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
        } else {
            Write-MessageLog -Message "Invalid 'Disposal Date' for Config ID: $($Data.Config_RecID), Name: $($Data.'Configuration Name'). Skipping." -Level "ERROR"
            return
        }
    }

    Update-ConfigurationStatic -Data $Data -Config $Config -WarrantyExpirationDate $WarrantyExpirationDate
    Update-ConfigurationArray -Data $Data -Config $Config -Company $Company -Contact $Contact -Type $Type -Manufacturer $Manufacturer -Vendor $Vendor -Site $Site -DateLastSighted $DateLastSighted -DisposalDate $DisposalDate

    Write-MessageLog -Message "Finished processing Config ID: $($Data.Config_RecID), Name: $($Data.'Configuration Name')."
}

# Run Main Process
ProcessMainFunction