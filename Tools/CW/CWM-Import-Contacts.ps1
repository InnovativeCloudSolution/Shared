# Import the common module
. "$PSScriptRoot\CWM-Common.ps1"

# CSV Path
$CsvPath = ".docs\DropBear IT\DropBear-Contacts-Template.csv"

function Import-Contacts {
    # Initialize logging
    $logPath = Initialize-Logging -LogName "CWM-Contacts-Log"
    
    Write-Log "=========================================" -Level "INFO"
    Write-Log "ConnectWise Manage Contact Importer" -Level "INFO"
    Write-Log "Log File: $logPath" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "" -Level "INFO"
    
    if (-not (Test-Path $CsvPath)) {
        Write-Log "CSV file not found: $CsvPath" -Level "ERROR"
        exit 1
    }
    
    Write-Log "CSV file loaded: $CsvPath" -Level "SUCCESS"
    
    Connect-CWM
    
    $csvData = Import-Csv -Path $CsvPath
    Write-Log "Total CSV entries: $($csvData.Count)" -Level "INFO"
    Write-Log "" -Level "INFO"
    
    foreach ($row in $csvData) {
        $companyIdentifier = $row.CompanyIdentifier.Trim()
        $firstName = $row.FirstName.Trim()
        $lastName = $row.LastName.Trim()
        
        if ([string]::IsNullOrWhiteSpace($companyIdentifier) -or [string]::IsNullOrWhiteSpace($firstName) -or [string]::IsNullOrWhiteSpace($lastName)) {
            Write-Log "Skipping row with missing CompanyIdentifier, FirstName, or LastName" -Level "WARNING"
            continue
        }
        
        Write-Log "Processing: $firstName $lastName for company $companyIdentifier" -Level "INFO"
        
        # Get company
        $company = Get-CWMCompanyByIdentifier -Identifier $companyIdentifier
        
        if (-not $company) {
            Write-Log "  Company '$companyIdentifier' not found. Skipping..." -Level "ERROR"
            continue
        }
        
        $companyId = $company.id
        
        # Check if contact already exists
        $existingContact = Get-CWMContact -CompanyId $companyId -FirstName $firstName -LastName $lastName
        
        if ($existingContact) {
            Write-Log "  Contact already exists (ID: $($existingContact.id))" -Level "INFO"
            continue
        }
        
        # Build contact data object
        $contactData = @{
            firstName = $firstName
            lastName = $lastName
            company = @{ id = $companyId }
        }
        
        # Add optional fields
        if (-not [string]::IsNullOrWhiteSpace($row.Title)) {
            $contactData.title = $row.Title.Trim()
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.Department)) {
            $contactData.department = @{ name = $row.Department.Trim() }
        }
        
        # Build communications array
        $communications = @()
        
        if (-not [string]::IsNullOrWhiteSpace($row.Email)) {
            $communications += @{
                type = @{ id = 1; name = "Email" }
                value = $row.Email.Trim()
                defaultFlag = $true
                communicationType = "Email"
            }
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.PhoneNumber)) {
            $communications += @{
                type = @{ id = 2; name = "Direct" }
                value = $row.PhoneNumber.Trim()
                defaultFlag = $false
                communicationType = "Phone"
            }
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.MobileNumber)) {
            $communications += @{
                type = @{ id = 4; name = "Mobile" }
                value = $row.MobileNumber.Trim()
                defaultFlag = $false
                communicationType = "Phone"
            }
        }
        
        if ($communications.Count -gt 0) {
            $contactData.communicationItems = $communications
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.DefaultContactFlag)) {
            $contactData.defaultFlag = [bool]::Parse($row.DefaultContactFlag.Trim())
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.InactiveFlag)) {
            $contactData.inactiveFlag = [bool]::Parse($row.InactiveFlag.Trim())
        }
        
        # Create the contact
        New-CWMContact -ContactData $contactData | Out-Null
    }
    
    Write-Log "" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "All contacts processed successfully!" -Level "SUCCESS"
    Write-LogSummary
}

# Main script execution
Import-Contacts
