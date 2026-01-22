# Parameters must be first
param(
    [Parameter(Mandatory=$false)]
    [string]$CsvPath = ".docs\DropBear IT\DropBear-Statuses-Template.csv"
)

# Import the common module
. "$PSScriptRoot\CWM-Common.ps1"

# Initialize logging
$logPath = Initialize-Logging -LogName "CWM-Statuses-Update"

Write-Log "=========================================" -Level "INFO"
Write-Log "ConnectWise Manage Status Updater" -Level "INFO"
Write-Log "Log File: $logPath" -Level "INFO"
Write-Log "=========================================" -Level "INFO"
Write-Log "" -Level "INFO"

# Validate CSV file
if (-not (Test-Path $CsvPath)) {
    Write-Log "CSV file not found: $CsvPath" -Level "ERROR"
    exit 1
}

# Load CSV
try {
    $statuses = Import-Csv $CsvPath
    Write-Log "CSV file loaded: $CsvPath" -Level "SUCCESS"
} catch {
    Write-Log "Failed to load CSV file: $_" -Level "ERROR"
    exit 1
}

# Set up API connection
Connect-CWM

# Get all service boards
$uri = "$script:CWMBaseUrl/service/boards"
$boards = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get

Write-Log "Available service boards:" -Level "INFO"
foreach ($b in $boards) {
    Write-Log "  ID: $($b.id) - $($b.name)" -Level "INFO"
}
Write-Log "" -Level "INFO"

# Function to update a status
function Update-CWMStatus {
    param(
        [int]$BoardId,
        [int]$StatusId,
        [object]$ExistingStatus,
        [hashtable]$Updates
    )
    
    $uri = "$script:CWMBaseUrl/service/boards/$BoardId/statuses/$StatusId"
    
    # Apply updates to existing status object
    foreach ($key in $Updates.Keys) {
        $ExistingStatus.$key = $Updates[$key]
    }
    
    $body = $ExistingStatus | ConvertTo-Json -Depth 10
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Put -Body $body -ContentType "application/json"
        Write-Log "  Updated status: $($response.name) - Sort Order: $($response.sortOrder)" -Level "SUCCESS"
        return $response
    } catch {
        Write-Log "Failed to update status ID '$StatusId': $_" -Level "ERROR"
        return $null
    }
}

# Process statuses
$totalStatuses = $statuses.Count
Write-Log "Total statuses to process: $totalStatuses" -Level "INFO"
Write-Log "" -Level "INFO"

$currentBoard = $null
$currentBoardId = 0

foreach ($row in $statuses) {
    # Skip empty rows
    if ([string]::IsNullOrWhiteSpace($row.name)) {
        continue
    }
    
    # Check if we need to switch boards
    if (-not [string]::IsNullOrWhiteSpace($row.BoardName)) {
        $csvBoardName = $row.BoardName.Trim()
        $tempBoard = $boards | Where-Object { $_.name -eq $csvBoardName } | Select-Object -First 1
        
        if ($tempBoard -and ($null -eq $currentBoard -or $tempBoard.id -ne $currentBoard.id)) {
            $currentBoard = $tempBoard
            $currentBoardId = $currentBoard.id
            Write-Log "" -Level "INFO"
            Write-Log "Processing board: $($currentBoard.name) (ID: $($currentBoard.id))" -Level "SUCCESS"
            Write-Log "" -Level "INFO"
        } elseif (-not $tempBoard) {
            Write-Log "Board not found: $csvBoardName - Skipping" -Level "WARNING"
            continue
        }
    }
    
    if ($currentBoardId -eq 0) {
        Write-Log "No board selected - Skipping status: $($row.name)" -Level "WARNING"
        continue
    }
    
    # Get existing statuses for this board
    $uri = "$script:CWMBaseUrl/service/boards/$currentBoardId/statuses"
    $existingStatuses = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
    
    # Find matching status by name
    $existingStatus = $existingStatuses | Where-Object { $_.name -eq $row.name } | Select-Object -First 1
    
    if ($existingStatus) {
        Write-Log "Processing: $($row.name)" -Level "INFO"
        
        # Build update data
        $updateData = @{}
        
        # Only update sortOrder if it's different
        if ($row.sortOrder -ne $existingStatus.sortOrder) {
            $updateData.sortOrder = [int]$row.sortOrder
        }
        
        # Update other fields if they're different
        if ($row.displayOnBoard -ne $existingStatus.displayOnBoard.ToString()) {
            $updateData.displayOnBoard = [bool]::Parse($row.displayOnBoard)
        }
        
        if ($row.inactive -ne $existingStatus.inactive.ToString()) {
            $updateData.inactive = [bool]::Parse($row.inactive)
        }
        
        if ($row.closedStatus -ne $existingStatus.closedStatus.ToString()) {
            $updateData.closedStatus = [bool]::Parse($row.closedStatus)
        }
        
        if ($row.timeEntryNotAllowed -ne $existingStatus.timeEntryNotAllowed.ToString()) {
            $updateData.timeEntryNotAllowed = [bool]::Parse($row.timeEntryNotAllowed)
        }
        
        if ($row.defaultFlag -ne $existingStatus.defaultFlag.ToString()) {
            $updateData.defaultFlag = [bool]::Parse($row.defaultFlag)
        }
        
        if ($row.escalationStatus -ne $existingStatus.escalationStatus) {
            $updateData.escalationStatus = $row.escalationStatus
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.customerPortalDescription) -and $row.customerPortalDescription -ne $existingStatus.customerPortalDescription) {
            $updateData.customerPortalDescription = $row.customerPortalDescription
        }
        
        if ($row.customerPortalFlag -ne $existingStatus.customerPortalFlag.ToString()) {
            $updateData.customerPortalFlag = [bool]::Parse($row.customerPortalFlag)
        }
        
        # Only update if there are changes
        if ($updateData.Count -gt 0) {
            Update-CWMStatus -BoardId $currentBoardId -StatusId $existingStatus.id -ExistingStatus $existingStatus -Updates $updateData
        } else {
            Write-Log "  No changes needed for: $($row.name)" -Level "INFO"
        }
    } else {
        Write-Log "Status not found: $($row.name) - Skipping" -Level "WARNING"
    }
}

Write-Log "" -Level "INFO"
Write-Log "=========================================" -Level "INFO"
Write-Log "All statuses processed!" -Level "SUCCESS"
Write-Log "" -Level "INFO"
Write-Log "=========================================" -Level "INFO"
Write-Log "SUMMARY:" -Level "INFO"
Write-Log "  Total Successes: $script:SuccessCount" -Level "INFO"
Write-Log "  Total Warnings:  $script:WarningCount" -Level "INFO"
Write-Log "  Total Errors:    $script:ErrorCount" -Level "INFO"
Write-Log "=========================================" -Level "INFO"
Write-Log "" -Level "INFO"
Write-Log "Log file saved to: $logPath" -Level "SUCCESS"
