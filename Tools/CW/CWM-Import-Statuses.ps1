# Parameters must be first
param(
    [Parameter(Mandatory=$false)]
    [string]$BoardName,
    [Parameter(Mandatory=$false)]
    [int]$BoardId,
    [Parameter(Mandatory=$false)]
    [string]$CsvPath = ".docs\DropBear IT\DropBear-Statuses-Template.csv"
)

# Import the common module
. "$PSScriptRoot\CWM-Common.ps1"

function Get-CWMStatusByName {
    param(
        [int]$BoardId,
        [string]$StatusName
    )
    
    $escapedName = $StatusName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/service/boards/$BoardId/statuses?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) { return $response[0] }
        return $null
    } catch {
        Write-Log "Failed to retrieve status '$StatusName': $_" -Level "ERROR"
        return $null
    }
}

function New-CWMStatus {
    param(
        [int]$BoardId,
        [hashtable]$StatusData
    )
    
    $uri = "$script:CWMBaseUrl/service/boards/$BoardId/statuses"
    $body = $StatusData | ConvertTo-Json -Depth 10
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Post -Body $body -ContentType "application/json"
        Write-Log "  Created status: $($response.name) (ID: $($response.id))" -Level "SUCCESS"
        return $response
    } catch {
        Write-Log "Failed to create status '$($StatusData.name)': $_" -Level "ERROR"
        Write-Log "Error details: $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

function Import-Statuses {
    # Initialize logging
    $logPath = Initialize-Logging -LogName "CWM-Statuses-Import"
    
    Write-Log "=========================================" -Level "INFO"
    Write-Log "ConnectWise Manage Status Importer" -Level "INFO"
    Write-Log "Log File: $logPath" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "" -Level "INFO"
    
    if (-not (Test-Path $CsvPath)) {
        Write-Log "CSV file not found: $CsvPath" -Level "ERROR"
        exit 1
    }
    
    Write-Log "CSV file loaded: $CsvPath" -Level "SUCCESS"
    
    Connect-CWM
    
    # Get all service boards
    $uri = "$script:CWMBaseUrl/service/boards"
    $boards = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
    
    Write-Log "Available service boards:" -Level "INFO"
    foreach ($b in $boards) {
        Write-Log "  ID: $($b.id) - $($b.name)" -Level "INFO"
    }
    Write-Log "" -Level "INFO"
    
    # Select board based on parameters
    $board = $null
    if ($BoardId -gt 0) {
        $board = $boards | Where-Object { $_.id -eq $BoardId } | Select-Object -First 1
    } elseif (-not [string]::IsNullOrWhiteSpace($BoardName)) {
        $board = $boards | Where-Object { $_.name -like "*$BoardName*" } | Select-Object -First 1
    } else {
        # Default to Service Desk board
        $board = $boards | Where-Object { $_.name -like "*Service*Desk*" } | Select-Object -First 1
        if (-not $board) {
            $board = $boards[0]
        }
    }
    
    if (-not $board) {
        Write-Log "No service board found matching criteria" -Level "ERROR"
        exit 1
    }
    
    Write-Log "Using service board: $($board.name) (ID: $($board.id))" -Level "SUCCESS"
    Write-Log "" -Level "INFO"
    
    $csvData = Import-Csv -Path $CsvPath
    Write-Log "Total statuses to process: $($csvData.Count)" -Level "INFO"
    Write-Log "" -Level "INFO"
    
    foreach ($row in $csvData) {
        $statusName = $row.name.Trim()
        
        if ([string]::IsNullOrWhiteSpace($statusName)) {
            Write-Log "Skipping row with missing status name" -Level "WARNING"
            continue
        }
        
        # If CSV has BoardName column and no board was specified, use it to find the board
        if ([string]::IsNullOrWhiteSpace($BoardName) -and $BoardId -eq 0 -and -not [string]::IsNullOrWhiteSpace($row.BoardName)) {
            $csvBoardName = $row.BoardName.Trim()
            $tempBoard = $boards | Where-Object { $_.name -eq $csvBoardName } | Select-Object -First 1
            if ($tempBoard -and $tempBoard.id -ne $board.id) {
                $board = $tempBoard
                Write-Log "" -Level "INFO"
                Write-Log "Switching to board: $($board.name) (ID: $($board.id))" -Level "SUCCESS"
                Write-Log "" -Level "INFO"
            }
        }
        
        Write-Log "Processing: $statusName" -Level "INFO"
        
        # Check if status already exists
        $existingStatus = Get-CWMStatusByName -BoardId $board.id -StatusName $statusName
        
        if ($existingStatus) {
            Write-Log "  Status already exists (ID: $($existingStatus.id))" -Level "INFO"
            continue
        }
        
        # Build status data object
        $statusData = @{
            name = $statusName
            board = @{ id = $board.id }
        }
        
        # Add optional fields
        if (-not [string]::IsNullOrWhiteSpace($row.sortOrder)) {
            $statusData.sortOrder = [int]$row.sortOrder
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.displayOnBoard)) {
            $statusData.displayOnBoard = [bool]::Parse($row.displayOnBoard.Trim())
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.inactive)) {
            $statusData.inactive = [bool]::Parse($row.inactive.Trim())
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.closedStatus)) {
            $statusData.closedStatus = [bool]::Parse($row.closedStatus.Trim())
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.timeEntryNotAllowed)) {
            $statusData.timeEntryNotAllowed = [bool]::Parse($row.timeEntryNotAllowed.Trim())
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.defaultFlag)) {
            $statusData.defaultFlag = [bool]::Parse($row.defaultFlag.Trim())
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.escalationStatus)) {
            $statusData.escalationStatus = $row.escalationStatus.Trim()
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.customerPortalDescription)) {
            $statusData.customerPortalDescription = $row.customerPortalDescription.Trim()
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.customerPortalFlag)) {
            $statusData.customerPortalFlag = [bool]::Parse($row.customerPortalFlag.Trim())
        }
        
        # Create the status
        New-CWMStatus -BoardId $board.id -StatusData $statusData | Out-Null
    }
    
    Write-Log "" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "All statuses processed!" -Level "SUCCESS"
    Write-LogSummary
}

# Main script execution
Import-Statuses
