# Import the common module
. "$PSScriptRoot\CWM-Common.ps1"

# Output paths
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$CsvPath = ".docs\DropBear IT\DropBear-Products-Export_$timestamp.csv"

function Get-CWMProducts {
    # Initialize logging
    $logPath = Initialize-Logging -LogName "CWM-Products-List"
    
    Write-Log "=========================================" -Level "INFO"
    Write-Log "ConnectWise Manage Product List" -Level "INFO"
    Write-Log "Log File: $logPath" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "" -Level "INFO"
    
    Connect-CWM
    
    $uri = "$script:CWMBaseUrl/procurement/catalog"
    
    try {
        Write-Log "Retrieving products from ConnectWise Manage..." -Level "INFO"
        
        $allProducts = @()
        $page = 1
        $pageSize = 1000
        
        do {
            $pagedUri = "$uri`?page=$page&pageSize=$pageSize"
            $response = Invoke-RestMethod -Uri $pagedUri -Headers $script:CWMHeaders -Method Get
            
            if ($response.Count -gt 0) {
                $allProducts += $response
                Write-Log "  Retrieved page $page - $($response.Count) products" -Level "INFO"
                $page++
            }
        } while ($response.Count -eq $pageSize)
        
        Write-Log "" -Level "INFO"
        Write-Log "Total products found: $($allProducts.Count)" -Level "SUCCESS"
        Write-Log "" -Level "INFO"
        
        if ($allProducts.Count -gt 0) {
            # Display products
            Write-Log "=========================================" -Level "INFO"
            Write-Log "PRODUCT LIST:" -Level "INFO"
            Write-Log "=========================================" -Level "INFO"
            
            foreach ($product in $allProducts | Sort-Object identifier) {
                $category = if ($product.category) { $product.category.name } else { "N/A" }
                $subcategory = if ($product.subcategory) { $product.subcategory.name } else { "N/A" }
                $type = if ($product.type) { $product.type.name } else { "N/A" }
                $manufacturer = if ($product.manufacturer) { $product.manufacturer.name } else { "N/A" }
                $price = if ($product.price) { $product.price } else { "0.00" }
                $cost = if ($product.cost) { $product.cost } else { "0.00" }
                
                Write-Log "ID: $($product.id) | Identifier: $($product.identifier) | Description: $($product.description)" -Level "INFO"
                Write-Log "  Category: $category | Subcategory: $subcategory | Type: $type" -Level "INFO"
                Write-Log "  Manufacturer: $manufacturer | Price: `$$price | Cost: `$$cost" -Level "INFO"
                Write-Log "  Inactive: $($product.inactiveFlag)" -Level "INFO"
                Write-Log "" -Level "INFO"
            }
            
            # Export to CSV
            $csvData = $allProducts | Select-Object `
                id,
                identifier,
                description,
                @{Name='Category';Expression={if ($_.category) { $_.category.name } else { "" }}},
                @{Name='Subcategory';Expression={if ($_.subcategory) { $_.subcategory.name } else { "" }}},
                @{Name='Type';Expression={if ($_.type) { $_.type.name } else { "" }}},
                @{Name='Manufacturer';Expression={if ($_.manufacturer) { $_.manufacturer.name } else { "" }}},
                @{Name='ManufacturerPartNumber';Expression={$_.manufacturerPartNumber}},
                @{Name='Vendor';Expression={if ($_.vendor) { $_.vendor.name } else { "" }}},
                @{Name='VendorSKU';Expression={$_.vendorSku}},
                price,
                cost,
                @{Name='UnitOfMeasure';Expression={if ($_.unitOfMeasure) { $_.unitOfMeasure.name } else { "" }}},
                @{Name='Taxable';Expression={$_.taxableFlag}},
                @{Name='Inactive';Expression={$_.inactiveFlag}},
                @{Name='CustomerDescription';Expression={$_.customerDescription}}
            
            $csvData | Export-Csv -Path $CsvPath -NoTypeInformation
            Write-Log "Products exported to: $CsvPath" -Level "SUCCESS"
        }
        
        Write-Log "" -Level "INFO"
        Write-Log "=========================================" -Level "INFO"
        Write-Log "Product retrieval complete!" -Level "SUCCESS"
        Write-LogSummary
        
    } catch {
        Write-Log "Failed to retrieve products: $_" -Level "ERROR"
        Write-Log "Error details: $($_.Exception.Message)" -Level "ERROR"
    }
}

# Main script execution
Get-CWMProducts
