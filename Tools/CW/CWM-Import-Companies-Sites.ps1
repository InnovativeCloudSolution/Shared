# Import the common module
. "$PSScriptRoot\CWM-Common.ps1"

# CSV Paths
$CompanyCsvPath = ".docs\DropBear IT\DropBear-Companies-Template.csv"
$SiteCsvPath = ".docs\DropBear IT\DropBear-Sites-Template.csv"

function Import-CompaniesAndSites {
    # Initialize logging
    $logPath = Initialize-Logging -LogName "CWM-Companies-Sites-Log"
    
    Write-Log "=========================================" -Level "INFO"
    Write-Log "ConnectWise Manage Company & Site Importer" -Level "INFO"
    Write-Log "Log File: $logPath" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "" -Level "INFO"
    
    # Validate CSV files
    if (-not (Test-Path $CompanyCsvPath)) {
        Write-Log "Company CSV file not found: $CompanyCsvPath" -Level "ERROR"
        exit 1
    }
    
    if (-not (Test-Path $SiteCsvPath)) {
        Write-Log "Site CSV file not found: $SiteCsvPath" -Level "ERROR"
        exit 1
    }
    
    Write-Log "Company CSV loaded: $CompanyCsvPath" -Level "SUCCESS"
    Write-Log "Site CSV loaded: $SiteCsvPath" -Level "SUCCESS"
    
    Connect-CWM
    
    $companyCsvData = Import-Csv -Path $CompanyCsvPath
    $siteCsvData = Import-Csv -Path $SiteCsvPath
    
    Write-Log "Total companies to process: $($companyCsvData.Count)" -Level "INFO"
    Write-Log "Total sites to process: $($siteCsvData.Count)" -Level "INFO"
    Write-Log "" -Level "INFO"
    
    # Process companies first
    Write-Log "=========================================" -Level "INFO"
    Write-Log "STEP 1: Importing Companies" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "" -Level "INFO"
    
    foreach ($row in $companyCsvData) {
        $companyName = $row.CompanyName.Trim()
        $identifier = $row.Identifier.Trim()
        
        if ([string]::IsNullOrWhiteSpace($companyName) -or [string]::IsNullOrWhiteSpace($identifier)) {
            Write-Log "Skipping row with missing CompanyName or Identifier" -Level "WARNING"
            continue
        }
        
        Write-Log "Processing: $companyName ($identifier)" -Level "INFO"
        
        # Check if company already exists
        $existingCompany = Get-CWMCompanyByIdentifier -Identifier $identifier
        
        if ($existingCompany) {
            Write-Log "  Company already exists (ID: $($existingCompany.id))" -Level "INFO"
            continue
        }
        
        # Build company data object
        $companyData = @{
            name = $companyName
            identifier = $identifier
        }
        
        # Add status
        if (-not [string]::IsNullOrWhiteSpace($row.Status)) {
            $status = Get-CWMCompanyStatus -StatusName $row.Status.Trim()
            if ($status) {
                $companyData.status = @{ id = $status.id }
            }
        }
        
        # Add types (note: it's an array)
        if (-not [string]::IsNullOrWhiteSpace($row.Type)) {
            $type = Get-CWMCompanyType -TypeName $row.Type.Trim()
            if ($type) {
                $companyData.types = @(
                    @{ id = $type.id }
                )
            }
        }
        
        # Add territory
        if (-not [string]::IsNullOrWhiteSpace($row.Territory)) {
            $territory = Get-CWMTerritory -TerritoryName $row.Territory.Trim()
            if ($territory) {
                $companyData.territory = @{ id = $territory.id }
            }
        }
        
        # Add market
        if (-not [string]::IsNullOrWhiteSpace($row.Market)) {
            $market = Get-CWMMarket -MarketName $row.Market.Trim()
            if ($market) {
                $companyData.market = @{ id = $market.id }
            }
        }
        
        # Add site name (required by CWM)
        if (-not [string]::IsNullOrWhiteSpace($row.AddressLine1)) {
            $companyData.site = @{ name = $row.AddressLine1.Trim() }
        }
        
        # Add phone number
        if (-not [string]::IsNullOrWhiteSpace($row.PhoneNumber)) {
            $companyData.phoneNumber = $row.PhoneNumber.Trim()
        }
        
        # Add website
        if (-not [string]::IsNullOrWhiteSpace($row.Website)) {
            $companyData.website = $row.Website.Trim()
        }
        
        # Create the company
        New-CWMCompany -CompanyData $companyData | Out-Null
    }
    
    Write-Log "" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "STEP 2: Importing Sites" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "" -Level "INFO"
    
    # Process sites
    foreach ($row in $siteCsvData) {
        $companyIdentifier = $row.CompanyIdentifier.Trim()
        $siteName = $row.SiteName.Trim()
        
        if ([string]::IsNullOrWhiteSpace($companyIdentifier) -or [string]::IsNullOrWhiteSpace($siteName)) {
            Write-Log "Skipping row with missing CompanyIdentifier or SiteName" -Level "WARNING"
            continue
        }
        
        Write-Log "Processing: $siteName for company $companyIdentifier" -Level "INFO"
        
        # Get company
        $company = Get-CWMCompanyByIdentifier -Identifier $companyIdentifier
        
        if (-not $company) {
            Write-Log "  Company '$companyIdentifier' not found. Skipping..." -Level "ERROR"
            continue
        }
        
        $companyId = $company.id
        
        # Check if site already exists
        $existingSite = Get-CWMSite -CompanyId $companyId -SiteName $siteName
        
        if ($existingSite) {
            Write-Log "  Site already exists (ID: $($existingSite.id))" -Level "INFO"
            continue
        }
        
        # Build site data object
        $siteData = @{
            name = $siteName
        }
        
        # Add optional fields
        if (-not [string]::IsNullOrWhiteSpace($row.AddressLine1)) {
            $siteData.addressLine1 = $row.AddressLine1.Trim()
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.AddressLine2)) {
            $siteData.addressLine2 = $row.AddressLine2.Trim()
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.City)) {
            $siteData.city = $row.City.Trim()
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.State)) {
            $siteData.stateIdentifier = $row.State.Trim()
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.Zip)) {
            $siteData.zip = $row.Zip.Trim()
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.Country)) {
            $siteData.country = @{ name = $row.Country.Trim() }
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.PhoneNumber)) {
            $siteData.phoneNumber = $row.PhoneNumber.Trim()
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.FaxNumber)) {
            $siteData.faxNumber = $row.FaxNumber.Trim()
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.PrimaryAddressFlag)) {
            $siteData.primaryAddressFlag = [bool]::Parse($row.PrimaryAddressFlag.Trim())
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.DefaultShippingFlag)) {
            $siteData.defaultShippingFlag = [bool]::Parse($row.DefaultShippingFlag.Trim())
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.DefaultBillingFlag)) {
            $siteData.defaultBillingFlag = [bool]::Parse($row.DefaultBillingFlag.Trim())
        }
        
        # Create the site
        New-CWMSite -CompanyId $companyId -SiteData $siteData | Out-Null
    }
    
    Write-Log "" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "All companies and sites processed!" -Level "SUCCESS"
    Write-LogSummary
}

# Main script execution
Import-CompaniesAndSites
