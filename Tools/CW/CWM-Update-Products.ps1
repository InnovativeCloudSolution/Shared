# Import the common module
. "$PSScriptRoot\CWM-Common.ps1"

# CSV Path
$CsvPath = ".docs\DropBear IT\DropBear-Products-Template.csv"

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

function Update-CWMProduct {
    param(
        [int]$ProductId,
        [object]$ExistingProduct,
        [hashtable]$Updates
    )
    
    $uri = "$script:CWMBaseUrl/procurement/catalog/$ProductId"
    
    # Apply updates to existing product object
    foreach ($key in $Updates.Keys) {
        $ExistingProduct.$key = $Updates[$key]
    }
    
    $body = $ExistingProduct | ConvertTo-Json -Depth 10
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Put -Body $body -ContentType "application/json"
        Write-Log "  Updated product: $($response.identifier) - Price: $($response.price), Cost: $($response.cost), Inactive: $($response.inactiveFlag)" -Level "SUCCESS"
        return $response
    } catch {
        Write-Log "Failed to update product ID '$ProductId': $_" -Level "ERROR"
        return $null
    }
}

function Update-Products {
    # Initialize logging
    $logPath = Initialize-Logging -LogName "CWM-Products-Update"
    
    Write-Log "=========================================" -Level "INFO"
    Write-Log "ConnectWise Manage Product Updater" -Level "INFO"
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
        
        if ([string]::IsNullOrWhiteSpace($identifier)) {
            Write-Log "Skipping row with missing Identifier" -Level "WARNING"
            continue
        }
        
        Write-Log "Processing: $identifier" -Level "INFO"
        
        # Check if product exists
        $existingProduct = Get-CWMProductByIdentifier -Identifier $identifier
        
        if (-not $existingProduct) {
            Write-Log "  Product does not exist, skipping" -Level "WARNING"
            continue
        }
        
        # Build update data object (only include fields that need updating)
        $updateData = @{}
        
        # Update price if provided
        if (-not [string]::IsNullOrWhiteSpace($row.price)) {
            $newPrice = [decimal]$row.price
            Write-Log "  DEBUG: Existing price: $($existingProduct.price), New price: $newPrice" -Level "INFO"
            if ($existingProduct.price -ne $newPrice) {
                $updateData.price = $newPrice
                Write-Log "  DEBUG: Adding price to update data" -Level "INFO"
            }
        }
        
        # Update cost if provided
        if (-not [string]::IsNullOrWhiteSpace($row.cost)) {
            $newCost = [decimal]$row.cost
            if ($existingProduct.cost -ne $newCost) {
                $updateData.cost = $newCost
            }
        }
        
        # Update inactive flag if provided
        if (-not [string]::IsNullOrWhiteSpace($row.Inactive)) {
            $newInactive = [bool]::Parse($row.Inactive.Trim())
            if ($existingProduct.inactiveFlag -ne $newInactive) {
                $updateData.inactiveFlag = $newInactive
            }
        }
        
        # Only update if there are changes
        if ($updateData.Count -eq 0) {
            Write-Log "  No changes needed" -Level "INFO"
            continue
        }
        
        # Update the product
        Update-CWMProduct -ProductId $existingProduct.id -ExistingProduct $existingProduct -Updates $updateData | Out-Null
    }
    
    Write-Log "" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "All products processed!" -Level "SUCCESS"
    Write-LogSummary
}

# Main script execution
Update-Products
