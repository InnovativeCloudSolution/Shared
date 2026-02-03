# Parameters must be first
param(
    [Parameter(Mandatory=$false)]
    [string]$CsvPath = ".docs\DropBear IT\DropBear-Project-Tickets-Template.csv"
)

# Import the common module
. "$PSScriptRoot\CWM-Common.ps1"

function Get-CWMProjectByName {
    param([string]$ProjectName)
    
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

function Get-CWMPriority {
    param([string]$PriorityName)
    
    $escapedName = $PriorityName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/service/priorities?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) { return $response[0] }
        return $null
    } catch {
        Write-Log "Failed to retrieve priority '$PriorityName': $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMTicketStatus {
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

function New-CWMProjectTicket {
    param([hashtable]$TicketData)
    
    $uri = "$script:CWMBaseUrl/project/tickets"
    $body = $TicketData | ConvertTo-Json -Depth 10
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Post -Body $body -ContentType "application/json"
        Write-Log "      Created ticket: $($response.summary) (ID: $($response.id))" -Level "SUCCESS"
        return $response
    } catch {
        Write-Log "Failed to create ticket '$($TicketData.summary)': $_" -Level "ERROR"
        Write-Log "Error details: $($_.Exception.Message)" -Level "ERROR"
        if ($_.ErrorDetails.Message) {
            Write-Log "API Error: $($_.ErrorDetails.Message)" -Level "ERROR"
        }
        return $null
    }
}

function Import-ProjectTickets {
    # Initialize logging
    $logPath = Initialize-Logging -LogName "CWM-Project-Tickets-Import"
    
    Write-Log "=========================================" -Level "INFO"
    Write-Log "ConnectWise Manage Project Tickets Importer" -Level "INFO"
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
        $tickets = Import-Csv $CsvPath
        Write-Log "CSV file loaded: $CsvPath" -Level "SUCCESS"
    } catch {
        Write-Log "Failed to load CSV file: $_" -Level "ERROR"
        exit 1
    }
    
    # Connect to API
    Connect-CWM
    
    Write-Log "Total tickets to process: $($tickets.Count)" -Level "INFO"
    Write-Log "" -Level "INFO"
    
    # Group tickets by project
    $ticketsByProject = $tickets | Group-Object -Property projectName
    
    foreach ($projectGroup in $ticketsByProject) {
        $projectName = $projectGroup.Name
        Write-Log "Processing project: $projectName" -Level "INFO"
        
        # Get project
        $project = Get-CWMProjectByName -ProjectName $projectName
        if (-not $project) {
            Write-Log "  Project not found: $projectName" -Level "ERROR"
            continue
        }
        
        Write-Log "  Found project (ID: $($project.id))" -Level "SUCCESS"
        
        # Group by phase
        $ticketsByPhase = $projectGroup.Group | Group-Object -Property phaseDescription
        
        foreach ($phaseGroup in $ticketsByPhase) {
            $phaseDesc = $phaseGroup.Name
            Write-Log "  Processing phase: $phaseDesc" -Level "INFO"
            
            # Get phase
            $phase = Get-CWMProjectPhaseByDescription -ProjectId $project.id -Description $phaseDesc
            if (-not $phase) {
                Write-Log "    Phase not found: $phaseDesc" -Level "ERROR"
                continue
            }
            
            foreach ($row in $phaseGroup.Group) {
                # Build ticket data for project ticket
                $ticketData = @{
                    summary = $row.summary
                    project = @{ id = $project.id }
                    phase = @{ id = $phase.id }
                }
                
                if (-not [string]::IsNullOrWhiteSpace($row.budgetHours)) {
                    $ticketData.budgetHours = [decimal]$row.budgetHours
                }
                
                if (-not [string]::IsNullOrWhiteSpace($row.initialDescription)) {
                    $ticketData.initialDescription = $row.initialDescription
                }
                
                # Create project ticket
                $result = New-CWMProjectTicket -TicketData $ticketData
            }
        }
        
        Write-Log "" -Level "INFO"
    }
    
    Write-Log "=========================================" -Level "INFO"
    Write-Log "All project tickets processed!" -Level "SUCCESS"
    Write-Log "" -Level "INFO"
    Write-LogSummary
}

# Main script execution
Import-ProjectTickets
