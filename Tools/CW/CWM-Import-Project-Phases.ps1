# Parameters must be first
param(
    [Parameter(Mandatory=$false)]
    [string]$CsvPath = ".docs\DropBear IT\DropBear-Project-Phases-Template.csv"
)

# Import the common module
. "$PSScriptRoot\CWM-Common.ps1"

function Get-CWMProjectByName {
    param(
        [string]$ProjectName
    )
    
    $escapedName = $ProjectName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/project/projects?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) { return $response[0] }
        return $null
    } catch {
        Write-Log "Failed to retrieve project '$ProjectName': $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMProjectPhaseStatus {
    param([string]$StatusName)
    
    $escapedName = $StatusName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/project/statuses?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) { return $response[0] }
        return $null
    } catch {
        Write-Log "Failed to retrieve phase status '$StatusName': $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMProjectPhaseByDescription {
    param(
        [int]$ProjectId,
        [string]$Description
    )
    
    $escapedDesc = $Description -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("description='$escapedDesc'")
    $uri = "$script:CWMBaseUrl/project/projects/$ProjectId/phases?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) { return $response[0] }
        return $null
    } catch {
        Write-Log "Failed to retrieve phase '$Description': $_" -Level "ERROR"
        return $null
    }
}

function New-CWMProjectPhase {
    param(
        [int]$ProjectId,
        [hashtable]$PhaseData
    )
    
    $uri = "$script:CWMBaseUrl/project/projects/$ProjectId/phases"
    $body = $PhaseData | ConvertTo-Json -Depth 10
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Post -Body $body -ContentType "application/json"
        Write-Log "    Created phase: $($response.description) (ID: $($response.id))" -Level "SUCCESS"
        return $response
    } catch {
        Write-Log "Failed to create phase '$($PhaseData.description)': $_" -Level "ERROR"
        Write-Log "Error details: $($_.Exception.Message)" -Level "ERROR"
        if ($_.ErrorDetails.Message) {
            Write-Log "API Error: $($_.ErrorDetails.Message)" -Level "ERROR"
        }
        return $null
    }
}

function Import-ProjectPhases {
    # Initialize logging
    $logPath = Initialize-Logging -LogName "CWM-Project-Phases-Import"
    
    Write-Log "=========================================" -Level "INFO"
    Write-Log "ConnectWise Manage Project Phases Importer" -Level "INFO"
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
        $phases = Import-Csv $CsvPath
        Write-Log "CSV file loaded: $CsvPath" -Level "SUCCESS"
    } catch {
        Write-Log "Failed to load CSV file: $_" -Level "ERROR"
        exit 1
    }
    
    # Connect to API
    Connect-CWM
    
    Write-Log "Total phases to process: $($phases.Count)" -Level "INFO"
    Write-Log "" -Level "INFO"
    
    # Group phases by project
    $phasesByProject = $phases | Group-Object -Property projectName
    
    foreach ($projectGroup in $phasesByProject) {
        $projectName = $projectGroup.Name
        Write-Log "Processing project: $projectName" -Level "INFO"
        
        # Get project
        $project = Get-CWMProjectByName -ProjectName $projectName
        if (-not $project) {
            Write-Log "  Project not found: $projectName" -Level "ERROR"
            continue
        }
        
        Write-Log "  Found project (ID: $($project.id))" -Level "SUCCESS"
        
        # First pass: Create all phases without parent relationships
        $createdPhases = @{}
        
        foreach ($row in $projectGroup.Group) {
            # Skip empty rows
            if ([string]::IsNullOrWhiteSpace($row.phaseDescription)) {
                continue
            }
            
            # Check if phase already exists
            $existingPhase = Get-CWMProjectPhaseByDescription -ProjectId $project.id -Description $row.phaseDescription
            if ($existingPhase) {
                Write-Log "    Phase already exists: $($row.phaseDescription) (ID: $($existingPhase.id))" -Level "WARNING"
                $createdPhases[$row.phaseDescription] = $existingPhase
                continue
            }
            
            # Get status
            $status = Get-CWMProjectPhaseStatus -StatusName $row.statusName
            if (-not $status) {
                Write-Log "    Status not found: $($row.statusName)" -Level "ERROR"
                continue
            }
            
            # Build phase data
            $phaseData = @{
                description = $row.phaseDescription
                status = @{ id = $status.id }
            }
            
            if (-not [string]::IsNullOrWhiteSpace($row.budgetHours)) {
                $phaseData.budgetHours = [decimal]$row.budgetHours
            }
            
            if (-not [string]::IsNullOrWhiteSpace($row.billTime)) {
                $phaseData.billTime = $row.billTime
            }
            
            if (-not [string]::IsNullOrWhiteSpace($row.notes)) {
                $phaseData.notes = $row.notes
            }
            
            if (-not [string]::IsNullOrWhiteSpace($row.scheduledStart)) {
                $startDate = [DateTime]::Parse($row.scheduledStart)
                $phaseData.scheduledStart = $startDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
            }
            
            if (-not [string]::IsNullOrWhiteSpace($row.scheduledEnd)) {
                $endDate = [DateTime]::Parse($row.scheduledEnd)
                $phaseData.scheduledEnd = $endDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
            }
            
            # Create phase
            $result = New-CWMProjectPhase -ProjectId $project.id -PhaseData $phaseData
            if ($result) {
                $createdPhases[$row.phaseDescription] = $result
            }
        }
        
        Write-Log "" -Level "INFO"
    }
    
    Write-Log "=========================================" -Level "INFO"
    Write-Log "All project phases processed!" -Level "SUCCESS"
    Write-Log "" -Level "INFO"
    Write-LogSummary
}

# Main script execution
Import-ProjectPhases
