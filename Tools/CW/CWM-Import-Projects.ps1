# Parameters must be first
param(
    [Parameter(Mandatory=$false)]
    [string]$CsvPath = ".docs\DropBear IT\DropBear-Projects-Template.csv"
)

# Import the common module
. "$PSScriptRoot\CWM-Common.ps1"

function Get-CWMProjectByName {
    param(
        [string]$ProjectName,
        [int]$CompanyId
    )
    
    $escapedName = $ProjectName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName' and company/id=$CompanyId")
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

function Get-CWMProjectStatus {
    param([string]$StatusName)
    
    $escapedName = $StatusName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/project/statuses?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) { return $response[0] }
        return $null
    } catch {
        Write-Log "Failed to retrieve project status '$StatusName': $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMProjectType {
    param([string]$TypeName)
    
    $escapedName = $TypeName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/project/projectTypes?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) { return $response[0] }
        return $null
    } catch {
        Write-Log "Failed to retrieve project type '$TypeName': $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMMember {
    param([string]$Identifier)
    
    $escapedIdentifier = $Identifier -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("identifier='$escapedIdentifier'")
    $uri = "$script:CWMBaseUrl/system/members?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) { return $response[0] }
        return $null
    } catch {
        Write-Log "Failed to retrieve member '$Identifier': $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMDepartment {
    param([string]$Identifier)
    
    $escapedIdentifier = $Identifier -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("identifier='$escapedIdentifier'")
    $uri = "$script:CWMBaseUrl/system/departments?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) { return $response[0] }
        return $null
    } catch {
        Write-Log "Failed to retrieve department '$Identifier': $_" -Level "ERROR"
        return $null
    }
}

function New-CWMProject {
    param([hashtable]$ProjectData)
    
    $uri = "$script:CWMBaseUrl/project/projects"
    $body = $ProjectData | ConvertTo-Json -Depth 10
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Post -Body $body -ContentType "application/json"
        Write-Log "  Created project: $($response.name) (ID: $($response.id))" -Level "SUCCESS"
        return $response
    } catch {
        Write-Log "Failed to create project '$($ProjectData.name)': $_" -Level "ERROR"
        Write-Log "Error details: $($_.Exception.Message)" -Level "ERROR"
        if ($_.ErrorDetails.Message) {
            Write-Log "API Error: $($_.ErrorDetails.Message)" -Level "ERROR"
        }
        return $null
    }
}

function Import-Projects {
    # Initialize logging
    $logPath = Initialize-Logging -LogName "CWM-Projects-Import"
    
    Write-Log "=========================================" -Level "INFO"
    Write-Log "ConnectWise Manage Project Importer" -Level "INFO"
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
        $projects = Import-Csv $CsvPath
        Write-Log "CSV file loaded: $CsvPath" -Level "SUCCESS"
    } catch {
        Write-Log "Failed to load CSV file: $_" -Level "ERROR"
        exit 1
    }
    
    # Connect to API
    Connect-CWM
    
    Write-Log "Total projects to process: $($projects.Count)" -Level "INFO"
    Write-Log "" -Level "INFO"
    
    foreach ($row in $projects) {
        # Skip empty rows
        if ([string]::IsNullOrWhiteSpace($row.name)) {
            continue
        }
        
        Write-Log "Processing: $($row.name)" -Level "INFO"
        
        # Get company
        $company = Get-CWMCompanyByIdentifier -Identifier $row.companyIdentifier
        if (-not $company) {
            Write-Log "  Company not found: $($row.companyIdentifier)" -Level "ERROR"
            continue
        }
        
        # Check if project already exists
        $existingProject = Get-CWMProjectByName -ProjectName $row.name -CompanyId $company.id
        if ($existingProject) {
            Write-Log "  Project already exists (ID: $($existingProject.id))" -Level "WARNING"
            continue
        }
        
        # Get site
        $site = Get-CWMSite -CompanyId $company.id -SiteName $row.siteName
        if (-not $site) {
            Write-Log "  Site not found: $($row.siteName)" -Level "ERROR"
            continue
        }
        
        # Get board
        $board = Get-CWMServiceBoard -BoardName $row.boardName
        if (-not $board) {
            Write-Log "  Board not found: $($row.boardName)" -Level "ERROR"
            continue
        }
        
        # Get project type
        $type = Get-CWMProjectType -TypeName $row.typeName
        if (-not $type) {
            Write-Log "  Project type not found: $($row.typeName)" -Level "ERROR"
            continue
        }
        
        # Get status
        $status = Get-CWMProjectStatus -StatusName $row.statusName
        if (-not $status) {
            Write-Log "  Status not found: $($row.statusName)" -Level "ERROR"
            continue
        }
        
        # Get manager
        $manager = Get-CWMMember -Identifier $row.managerIdentifier
        if (-not $manager) {
            Write-Log "  Manager not found: $($row.managerIdentifier)" -Level "ERROR"
            continue
        }
        
        # Get location
        $location = Get-CWMTerritory -TerritoryName $row.locationName
        if (-not $location) {
            Write-Log "  Location not found: $($row.locationName)" -Level "ERROR"
            continue
        }
        
        # Get department
        $department = Get-CWMDepartment -Identifier $row.departmentIdentifier
        if (-not $department) {
            Write-Log "  Department not found: $($row.departmentIdentifier)" -Level "ERROR"
            continue
        }
        
        # Build project data
        $projectData = @{
            name = $row.name
            company = @{ id = $company.id }
            site = @{ id = $site.id }
            board = @{ id = $board.id }
            type = @{ id = $type.id }
            status = @{ id = $status.id }
            manager = @{ id = $manager.id }
            location = @{ id = $location.id }
            department = @{ id = $department.id }
            budgetHours = [decimal]$row.budgetHours
            billingMethod = $row.billingMethod
            billingRateType = $row.billingRateType
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.description)) {
            $projectData.description = $row.description
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.estimatedStart)) {
            $startDate = [DateTime]::Parse($row.estimatedStart)
            $projectData.estimatedStart = $startDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
        
        if (-not [string]::IsNullOrWhiteSpace($row.estimatedEnd)) {
            $endDate = [DateTime]::Parse($row.estimatedEnd)
            $projectData.estimatedEnd = $endDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
        
        # Create project
        $result = New-CWMProject -ProjectData $projectData
        
        Write-Log "" -Level "INFO"
    }
    
    Write-Log "=========================================" -Level "INFO"
    Write-Log "All projects processed!" -Level "SUCCESS"
    Write-Log "" -Level "INFO"
    Write-LogSummary
}

# Main script execution
Import-Projects
