# Import the common module
. "$PSScriptRoot\CWM-Common.ps1"

# Output paths
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$CsvPath = ".docs\DropBear IT\DropBear-Statuses-Export_$timestamp.csv"

function Get-CWMStatuses {
    param([int]$BoardId)
    
    # Initialize logging
    $logPath = Initialize-Logging -LogName "CWM-Statuses-List"
    
    Write-Log "=========================================" -Level "INFO"
    Write-Log "ConnectWise Manage Status List" -Level "INFO"
    Write-Log "Log File: $logPath" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "" -Level "INFO"
    
    Connect-CWM
    
    # Get all service boards
    $uri = "$script:CWMBaseUrl/service/boards"
    $boards = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
    
    Write-Log "Available service boards:" -Level "INFO"
    foreach ($b in $boards) {
        Write-Log "  ID: $($b.id) - $($b.name)" -Level "INFO"
    }
    Write-Log "" -Level "INFO"
    
    # Use the first board or specify one
    $board = $boards[0]
    if (-not $board) {
        Write-Log "No service boards found" -Level "ERROR"
        exit 1
    }
    
    Write-Log "Using service board: $($board.name) (ID: $($board.id))" -Level "SUCCESS"
    
    $uri = "$script:CWMBaseUrl/service/boards/$($board.id)/statuses"
    
    try {
        Write-Log "Retrieving statuses from ConnectWise Manage..." -Level "INFO"
        
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        
        Write-Log "" -Level "INFO"
        Write-Log "Total statuses found: $($response.Count)" -Level "SUCCESS"
        Write-Log "" -Level "INFO"
        
        if ($response.Count -gt 0) {
            # Display statuses
            Write-Log "=========================================" -Level "INFO"
            Write-Log "STATUS LIST:" -Level "INFO"
            Write-Log "=========================================" -Level "INFO"
            
            foreach ($status in $response | Sort-Object sortOrder) {
                Write-Log "ID: $($status.id) | Name: $($status.name) | Sort: $($status.sortOrder)" -Level "INFO"
                Write-Log "  Display on Board: $($status.displayOnBoard) | Closed: $($status.closedStatus) | Default: $($status.defaultFlag)" -Level "INFO"
                Write-Log "  Inactive: $($status.inactive) | Escalation: $($status.escalationStatus)" -Level "INFO"
                Write-Log "" -Level "INFO"
            }
            
            # Export to CSV
            $csvData = $response | Select-Object `
                id,
                name,
                sortOrder,
                displayOnBoard,
                inactive,
                closedStatus,
                timeEntryNotAllowed,
                defaultFlag,
                escalationStatus,
                customerPortalDescription,
                customerPortalFlag,
                @{Name='StatusIndicator';Expression={if ($_.statusIndicator) { $_.statusIndicator.name } else { "" }}},
                customStatusIndicatorName
            
            $csvData | Export-Csv -Path $CsvPath -NoTypeInformation
            Write-Log "Statuses exported to: $CsvPath" -Level "SUCCESS"
        }
        
        Write-Log "" -Level "INFO"
        Write-Log "=========================================" -Level "INFO"
        Write-Log "Status retrieval complete!" -Level "SUCCESS"
        Write-LogSummary
        
    } catch {
        Write-Log "Failed to retrieve statuses: $_" -Level "ERROR"
        Write-Log "Error details: $($_.Exception.Message)" -Level "ERROR"
    }
}

# Main script execution
Get-CWMStatuses
