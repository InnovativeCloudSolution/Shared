# Import the common module
. "$PSScriptRoot\CWM-Common.ps1"

# CSV Path
$CsvPath = ".docs\DropBear IT\DropBear-Boards-TypeSubtypeItem-Template.csv"

function Import-TypeSubtypeItem {
    # Initialize logging
    $logPath = Initialize-Logging -LogName "CWM-TypeSubtypeItem-Log"
    
    Write-Log "=========================================" -Level "INFO"
    Write-Log "ConnectWise Manage Type/Subtype/Item Importer" -Level "INFO"
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
        $boardName = $row.Board.Trim()
        $typeName = $row.Type.Trim()
        $subtypeName = $row.Subtype.Trim()
        $itemName = $row.Item.Trim()
        
        if ([string]::IsNullOrWhiteSpace($boardName) -or [string]::IsNullOrWhiteSpace($typeName)) {
            Write-Log "Skipping row with missing Board or Type" -Level "WARNING"
            continue
        }
        
        Write-Log "Processing: Board='$boardName', Type='$typeName', Subtype='$subtypeName', Item='$itemName'" -Level "INFO"
        
        # Get board
        $board = Get-CWMServiceBoard -BoardName $boardName
        
        if (-not $board) {
            Write-Log "  Board '$boardName' not found. Skipping..." -Level "ERROR"
            continue
        }
        
        $boardId = $board.id
        
        # Check if type exists
        $existingType = Get-CWMServiceType -BoardId $boardId -TypeName $typeName
        
        if (-not $existingType) {
            # Create type
            $typeData = @{
                name = $typeName
            }
            $existingType = New-CWMServiceType -BoardId $boardId -TypeData $typeData
            
            if (-not $existingType) {
                Write-Log "  Failed to create type. Skipping..." -Level "ERROR"
                continue
            }
        } else {
            Write-Log "  Type already exists (ID: $($existingType.id))" -Level "INFO"
        }
        
        # Process subtype if provided
        if (-not [string]::IsNullOrWhiteSpace($subtypeName)) {
            $existingSubtype = Get-CWMServiceSubType -BoardId $boardId -SubTypeName $subtypeName
            
            if (-not $existingSubtype) {
                # Create subtype
                $subtypeData = @{
                    name = $subtypeName
                }
                $existingSubtype = New-CWMServiceSubType -BoardId $boardId -SubTypeData $subtypeData
                
                if (-not $existingSubtype) {
                    Write-Log "  Failed to create subtype. Skipping item..." -Level "ERROR"
                    continue
                }
            } else {
                Write-Log "    Subtype already exists (ID: $($existingSubtype.id))" -Level "INFO"
            }
        }
        
        # Process item if provided
        if (-not [string]::IsNullOrWhiteSpace($itemName)) {
            $existingItem = Get-CWMServiceItem -BoardId $boardId -ItemName $itemName
            
            if (-not $existingItem) {
                # Create item
                $itemData = @{
                    name = $itemName
                }
                $existingItem = New-CWMServiceItem -BoardId $boardId -ItemData $itemData
                
                if (-not $existingItem) {
                    Write-Log "    Failed to create item." -Level "ERROR"
                }
            } else {
                Write-Log "      Item already exists (ID: $($existingItem.id))" -Level "INFO"
            }
        }
    }
    
    Write-Log "" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "All types, subtypes, and items processed!" -Level "SUCCESS"
    Write-LogSummary
}

# Main script execution
Import-TypeSubtypeItem
