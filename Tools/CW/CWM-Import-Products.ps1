# Import the common module
. "$PSScriptRoot\CWM-Common.ps1"

# CSV Path
$CsvPath = ".docs\DropBear IT\DropBear-Products-Template.csv"

function Get-CWMCatalogCategory {
    param([string]$CategoryName)
    
    $escapedName = $CategoryName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/procurement/categories?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) { return $response[0] }
        return $null
    } catch {
        Write-Log "Failed to retrieve category: $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMCatalogSubcategory {
    param([string]$SubcategoryName)
    
    $escapedName = $SubcategoryName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/procurement/subcategories?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) { return $response[0] }
        return $null
    } catch {
        Write-Log "Failed to retrieve subcategory: $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMCatalogType {
    param([string]$TypeName)
    
    $escapedName = $TypeName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/procurement/types?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) { return $response[0] }
        return $null
    } catch {
        Write-Log "Failed to retrieve type: $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMManufacturer {
    param([string]$ManufacturerName)
    
    $escapedName = $ManufacturerName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/procurement/manufacturers?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) { return $response[0] }
        return $null
    } catch {
        Write-Log "Failed to retrieve manufacturer: $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMProductByIdentifier {
    param([string]$Identifier)
    
    $escapedIdentifier = $Identifier -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("identifier='$escapedIdentifier'")
    $uri = "$script:CWMBaseUrl/procurement/catalog?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) { return $response[0] }
        return $null
    } catch {
        Write-Log "Failed to retrieve product '$Identifier': $_" -Level "ERROR"
        return $null
    }
}

function New-CWMProduct {
    param([hashtable]$ProductData)
    
    $uri = "$script:CWMBaseUrl/procurement/catalog"
    $body = $ProductData | ConvertTo-Json -Depth 10
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Post -Body $body
        Write-Log "  Created product: $($response.identifier) - $($response.description) (ID: $($response.id))" -Level "SUCCESS"
        return $response
    } catch {
        Write-Log "Failed to create product '$($ProductData.identifier)': $_" -Level "ERROR"
        return $null
    }
}

function Import-Products {
    # Initialize logging
    $logPath = Initialize-Logging -LogName "CWM-Products-Import"
    
    Write-Log "=========================================" -Level "INFO"
    Write-Log "ConnectWise Manage Product Importer" -Level "INFO"
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
    Write-Log "Total products to process: $($csvData.Count)" -Level "INFO"
    Write-Log "" -Level "INFO"
    
    foreach ($row in $csvData) {
        $identifier = $row.identifier.Trim()
        $description = $row.description.Trim()
        
        if ([string]::IsNullOrWhiteSpace($identifier) -or [string]::IsNullOrWhiteSpace($description)) {
            Write-Log "Skipping row with missing Identifier or Description" -Level "WARNING"
            continue
        }
        
        Write-Log "Processing: $identifier - $description" -Level "INFO"
        
        # Check if product already exists
        $existingProduct = Get-CWMProductByIdentifier -Identifier $identifier
        
        if ($existingProduct) {
            Write-Log "  Product already exists (ID: $($existingProduct.id))" -Level "INFO"
            continue
        }
        
        # Build product data object
        $productData = @{
            identifier = $identifier
            description = $description
        }
        
        # Add category
        if (-not [string]::IsNullOrWhiteSpace($row.Category)) {
            $category = Get-CWMCatalogCategory -CategoryName $row.Category.Trim()
            if ($category) {
                $productData.category = @{ id = $category.id }
            } else {
                Write-Log "  Warning: Category '$($row.Category)' not found" -Level "WARNING"
            }
        }
        
        # Add subcategory
        if (-not [string]::IsNullOrWhiteSpace($row.Subcategory)) {
            $subcategory = Get-CWMCatalogSubcategory -SubcategoryName $row.Subcategory.Trim()
            if ($subcategory) {
                $productData.subcategory = @{ id = $subcategory.id }
            } else {
                Write-Log "  Warning: Subcategory '$($row.Subcategory)' not found" -Level "WARNING"
            }
        }
        
        # Add type
        if (-not [string]::IsNullOrWhiteSpace($row.Type)) {
            $type = Get-CWMCatalogType -TypeName $row.Type.Trim()
            if ($type) {
                $productData.type = @{ id = $type.id }
            } else {
                Write-Log "  Warning: Type '$($row.Type)' not found" -Level "WARNING"
            }
        }
        
        # Add manufacturer
        if (-not [string]::IsNullOrWhiteSpace($row.Manufacturer)) {
            $manufacturer = Get-CWMManufacturer -ManufacturerName $row.Manufacturer.Trim()
            if ($manufacturer) {
                $productData.manufacturer = @{ id = $manufacturer.id }
            } else {
                Write-Log "  Warning: Manufacturer '$($row.Manufacturer)' not found" -Level "WARNING"
            }
        }
        
        # Add optional fields
        if (-not [string]::IsNullOrWhiteSpace($row.ManufacturerPartNumber)) {
            $productData.manufacturerPartNumber = $row.ManufacturerPartNumber.Trim()
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.VendorSKU)) {
            $productData.vendorSku = $row.VendorSKU.Trim()
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.price)) {
            $productData.price = [decimal]$row.price
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.cost)) {
            $productData.cost = [decimal]$row.cost
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.CustomerDescription)) {
            $productData.customerDescription = $row.CustomerDescription.Trim()
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.Taxable)) {
            $productData.taxableFlag = [bool]::Parse($row.Taxable.Trim())
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.Inactive)) {
            $productData.inactiveFlag = [bool]::Parse($row.Inactive.Trim())
        }
        
        # Create the product
        New-CWMProduct -ProductData $productData | Out-Null
    }
    
    Write-Log "" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "All products processed!" -Level "SUCCESS"
    Write-LogSummary
}

# Main script execution
Import-Products
