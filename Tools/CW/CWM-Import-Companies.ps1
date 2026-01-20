$CsvPath = ".docs\DropBear IT\DropBear-Companies-Template.csv"
$CWMUrl = "https://api-aus.myconnectwise.net"
$ApiVersion = "v4_6_release/apis/3.0"
$CompanyId = "dropbearit"
$PublicKey = "xAVcYWO20x5dRyG7"
$PrivateKey = "QUC1zTaMuUXbiJqX"
$ClientId = "1748c7f0-976c-4205-afa1-9bc9e1533565"

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogPath = ".logs\CWM-Companies-Log_$timestamp.txt"

$script:ErrorCount = 0
$script:WarningCount = 0
$script:SuccessCount = 0

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    Add-Content -Path $script:LogPath -Value $logEntry
    
    switch ($Level) {
        "ERROR" {
            Write-Host $Message -ForegroundColor Red
            $script:ErrorCount++
        }
        "WARNING" {
            Write-Host $Message -ForegroundColor Yellow
            $script:WarningCount++
        }
        "SUCCESS" {
            Write-Host $Message -ForegroundColor Green
            $script:SuccessCount++
        }
        default {
            Write-Host $Message
        }
    }
}

function Connect-CWM {
    param(
        [string]$CWMUrl,
        [string]$ApiVersion,
        [string]$CompanyId,
        [string]$PublicKey,
        [string]$PrivateKey,
        [string]$ClientId
    )
    
    $authString = "$CompanyId+$PublicKey`:$PrivateKey"
    $encodedAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($authString))
    
    $script:CWMHeaders = @{
        'Authorization' = "Basic $encodedAuth"
        'clientId' = $ClientId
        'Content-Type' = 'application/json'
    }
    
    $script:CWMBaseUrl = "$CWMUrl/$ApiVersion"
    
    Write-Log "Connected to ConnectWise Manage at $CWMUrl" -Level "SUCCESS"
}

function Get-CWMCompanyByIdentifier {
    param(
        [string]$Identifier
    )
    
    $escapedIdentifier = $Identifier -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("identifier='$escapedIdentifier'")
    $uri = "$script:CWMBaseUrl/company/companies?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        
        if ($response.Count -gt 0) {
            return $response[0]
        } else {
            return $null
        }
    } catch {
        Write-Log "Failed to check for existing company '$Identifier': $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMCompanyStatus {
    param(
        [string]$StatusName
    )
    
    if ([string]::IsNullOrWhiteSpace($StatusName)) {
        return $null
    }
    
    $escapedName = $StatusName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/company/companies/statuses?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        
        if ($response.Count -gt 0) {
            return $response[0]
        } else {
            Write-Log "Company status '$StatusName' not found" -Level "WARNING"
            return $null
        }
    } catch {
        Write-Log "Failed to retrieve company status '$StatusName': $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMCompanyType {
    param(
        [string]$TypeName
    )
    
    if ([string]::IsNullOrWhiteSpace($TypeName)) {
        return $null
    }
    
    $escapedName = $TypeName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/company/companies/types?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        
        if ($response.Count -gt 0) {
            return $response[0]
        } else {
            Write-Log "Company type '$TypeName' not found" -Level "WARNING"
            return $null
        }
    } catch {
        Write-Log "Failed to retrieve company type '$TypeName': $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMTerritory {
    param(
        [string]$TerritoryName
    )
    
    if ([string]::IsNullOrWhiteSpace($TerritoryName)) {
        return $null
    }
    
    $escapedName = $TerritoryName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/system/locations?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        
        if ($response.Count -gt 0) {
            return $response[0]
        } else {
            Write-Log "Territory '$TerritoryName' not found" -Level "WARNING"
            return $null
        }
    } catch {
        Write-Log "Failed to retrieve territory '$TerritoryName': $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMMarket {
    param(
        [string]$MarketName
    )
    
    if ([string]::IsNullOrWhiteSpace($MarketName)) {
        return $null
    }
    
    $escapedName = $MarketName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/company/markets?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        
        if ($response.Count -gt 0) {
            return $response[0]
        } else {
            Write-Log "Market '$MarketName' not found" -Level "WARNING"
            return $null
        }
    } catch {
        Write-Log "Failed to retrieve market '$MarketName': $_" -Level "ERROR"
        return $null
    }
}

function New-CWMCompany {
    param(
        [hashtable]$CompanyData
    )
    
    $uri = "$script:CWMBaseUrl/company/companies"
    $body = $CompanyData | ConvertTo-Json -Depth 10
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Post -Body $body
        Write-Log "Created company: $($response.name) (ID: $($response.id))" -Level "SUCCESS"
        return $response
    } catch {
        Write-Log "Failed to create company '$($CompanyData.name)': $_" -Level "ERROR"
        return $null
    }
}

function Import-Companies {
    Write-Log "=========================================" -Level "INFO"
    Write-Log "ConnectWise Manage Company Importer" -Level "INFO"
    Write-Log "Log File: $LogPath" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "" -Level "INFO"
    
    if (-not (Test-Path $CsvPath)) {
        Write-Log "CSV file not found: $CsvPath" -Level "ERROR"
        exit 1
    }
    
    Write-Log "CSV file loaded: $CsvPath" -Level "SUCCESS"
    
    Connect-CWM -CWMUrl $CWMUrl -ApiVersion $ApiVersion -CompanyId $CompanyId -PublicKey $PublicKey -PrivateKey $PrivateKey -ClientId $ClientId
    
    $csvData = Import-Csv -Path $CsvPath
    Write-Log "Total CSV entries: $($csvData.Count)" -Level "INFO"
    Write-Log "" -Level "INFO"
    
    foreach ($row in $csvData) {
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
            identifier = $identifier
            name = $companyName
        }
        
        # Add optional fields
        if (-not [string]::IsNullOrWhiteSpace($row.Status)) {
            $status = Get-CWMCompanyStatus -StatusName $row.Status.Trim()
            if ($status) {
                $companyData.status = @{ id = $status.id }
            }
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.Type)) {
            $type = Get-CWMCompanyType -TypeName $row.Type.Trim()
            if ($type) {
                $companyData.types = @(@{ id = $type.id })
            }
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.Territory)) {
            $territory = Get-CWMTerritory -TerritoryName $row.Territory.Trim()
            if ($territory) {
                $companyData.territory = @{ id = $territory.id }
            }
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.Market)) {
            $market = Get-CWMMarket -MarketName $row.Market.Trim()
            if ($market) {
                $companyData.market = @{ id = $market.id }
            }
        }
        
        # Address information
        if (-not [string]::IsNullOrWhiteSpace($row.AddressLine1)) {
            $companyData.addressLine1 = $row.AddressLine1.Trim()
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.City)) {
            $companyData.city = $row.City.Trim()
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.State)) {
            $companyData.state = $row.State.Trim()
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.Zip)) {
            $companyData.zip = $row.Zip.Trim()
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.Country)) {
            $companyData.country = @{ name = $row.Country.Trim() }
        }
        
        # Phone and website
        if (-not [string]::IsNullOrWhiteSpace($row.PhoneNumber)) {
            $companyData.phoneNumber = $row.PhoneNumber.Trim()
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.Website)) {
            $companyData.website = $row.Website.Trim()
        }
        
        # Create the company
        New-CWMCompany -CompanyData $companyData | Out-Null
    }
    
    Write-Log "" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "All companies processed successfully!" -Level "SUCCESS"
    Write-Log "" -Level "INFO"
    Write-Log "SUMMARY:" -Level "INFO"
    Write-Log "  Total Successes: $script:SuccessCount" -Level "INFO"
    Write-Log "  Total Warnings:  $script:WarningCount" -Level "INFO"
    Write-Log "  Total Errors:    $script:ErrorCount" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "" -Level "INFO"
    Write-Log "Log file saved to: $LogPath" -Level "SUCCESS"
}

# Main script execution
Import-Companies
