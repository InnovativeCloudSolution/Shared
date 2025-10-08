#Requires -Module SqlServer

<#
.SYNOPSIS
    CQL Service Tickets Export with embedded SQL functions and chunking capabilities
.DESCRIPTION
    Executes all CQL export operations with embedded SQL queries, connection management,
    and configurable chunking for notes and time entries to handle large datasets.
.PARAMETER ServerInstance
    SQL Server instance name. Defaults to .\SQLEXPRESS
.PARAMETER Database
    Target database name. Defaults to CQL_ServiceTickets
.PARAMETER CompanyRecID
    Company Record ID to filter data. Defaults to 2791
.PARAMETER DelayBetweenOperations
    Delay in seconds between each operation. Defaults to 30 seconds
.PARAMETER NotesChunkSize
    Number of records to process per notes chunk. Defaults to 1000
.PARAMETER TimeEntriesChunkSize
    Number of records to process per time entries chunk. Defaults to 1000
.PARAMETER LogPath
    Path for log file. Defaults to current directory with timestamp
.EXAMPLE
    .\Run-CQL-Export-Embedded.ps1
.EXAMPLE
    .\Run-CQL-Export-Embedded.ps1 -NotesChunkSize 500 -TimeEntriesChunkSize 750
#>

[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()][string]$ServerInstance = ".\SQLEXPRESS",
    [ValidateNotNullOrEmpty()][string]$Database = "CQL_ServiceTickets",
    [int]$CompanyRecID = 2791,
    [int]$DelayBetweenOperations = 5,
    [int]$NotesChunkSize = 1000,
    [int]$TimeEntriesChunkSize = 1000,
    [string]$LogPath = "$PSScriptRoot\CQL-Export-Log-$(Get-Date -Format 'yyyyMMdd-HHmmss').txt"
)

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Level] $Message"
    
    Write-Host $LogMessage -ForegroundColor $(
        switch ($Level) {
            "ERROR" { "Red" }
            "WARNING" { "Yellow" }
            "SUCCESS" { "Green" }
            default { "White" }
        }
    )
    
    Add-Content -Path $LogPath -Value $LogMessage
}

function Write-Progress-Log {
    param(
        [string]$Operation,
        [int]$Current,
        [int]$Total,
        [string]$Status = "Processing"
    )
    
    $PercentComplete = if ($Total -gt 0) { [math]::Round(($Current / $Total) * 100, 2) } else { 0 }
    Write-Progress -Activity $Operation -Status "$Status ($Current of $Total)" -PercentComplete $PercentComplete
    Write-Log "$Operation - $($Status): $Current of $Total ($PercentComplete%)"
}

function Test-SqlConnection {
    param(
        [string]$ServerInstance,
        [string]$Database
    )
    
    try {
        $Query = "SELECT 1 AS TestConnection"
        Invoke-Sqlcmd -ServerInstance $ServerInstance -Database "master" -Query $Query -TrustServerCertificate -ErrorAction Stop | Out-Null
        
        $DbQuery = "SELECT COUNT(*) AS DbExists FROM sys.databases WHERE name = '$Database'"
        $DbResult = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database "master" -Query $DbQuery -TrustServerCertificate -ErrorAction Stop
        
        if ($DbResult.DbExists -eq 0) {
            Write-Log "Database '$Database' does not exist. Creating it..." "WARNING"
            $CreateDbQuery = "CREATE DATABASE [$Database]"
            Invoke-Sqlcmd -ServerInstance $ServerInstance -Database "master" -Query $CreateDbQuery -TrustServerCertificate -ErrorAction Stop
            Write-Log "Database '$Database' created successfully" "SUCCESS"
        }
        
        Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query $Query -TrustServerCertificate -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        Write-Log "Connection test failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Invoke-SafeSqlQuery {
    param(
        [string]$Query,
        [string]$Operation,
        [int]$QueryTimeout = 300
    )
    
    try {
        Write-Log "Starting: $Operation"
        $StartTime = Get-Date
        
        $Result = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -Query $Query -QueryTimeout $QueryTimeout -TrustServerCertificate -ErrorAction Stop
        
        $Duration = (Get-Date) - $StartTime
        Write-Log "$Operation completed successfully in $($Duration.TotalMinutes.ToString('F2')) minutes" "SUCCESS"
        return @{ Success = $true; Result = $Result; Duration = $Duration }
    }
    catch {
        Write-Log "$Operation failed: $($_.Exception.Message)" "ERROR"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
    finally {
        # Force disconnect and reconnect for next operation
        try {
            [System.Data.SqlClient.SqlConnection]::ClearAllPools()
            Start-Sleep -Seconds 2
        }
        catch {
            Write-Log "Connection pool clear warning: $($_.Exception.Message)" "WARNING"
        }
    }
}

function Export-ConfigItems {
    param([int]$CompanyRecID)
    
    # Build query without interpolation to avoid quote escaping issues
    $Query = @"
USE [$Database];

-- Create table if not exists
IF OBJECT_ID('SR_Config_Items') IS NULL
CREATE TABLE SR_Config_Items (
    Config_RecID int PRIMARY KEY,
    Config_Name nvarchar(100),
    Config_Type nvarchar(50),
    Company_RecID int,
    Company_Name nvarchar(250),
    Serial_Number nvarchar(100),
    Vendor_Name nvarchar(100),
    Status nvarchar(50)
);

-- Clear existing data
DELETE FROM SR_Config_Items WHERE Company_RecID = $CompanyRecID;

-- Insert config items
INSERT INTO SR_Config_Items
SELECT 
    ci.Config_RecID,
    ci.Config_Name,
    'Configuration' AS Config_Type,
    ci.Company_RecID,
    c.Company_Name,
    ci.Serial_Number,
    mfg.Company_Name AS Vendor_Name,
    ISNULL(cs.Description, 'Active') AS Status
FROM CW_Report_DB.cwwebapp_mit.dbo.Config ci
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Company c ON ci.Company_RecID = c.Company_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Company mfg ON ci.Mfg_Company_RecID = mfg.Company_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Config_Status cs ON ci.Config_Status_RecID = cs.Config_Status_RecID
WHERE ci.Company_RecID = $CompanyRecID;
"@

    return Invoke-SafeSqlQuery -Query $Query -Operation "Export Config Items"
}

function Export-Events {
    param([int]$CompanyRecID)
    
    $Query = @"
USE [$Database];

-- Create table if not exists
IF OBJECT_ID('SR_Event') IS NULL
CREATE TABLE SR_Event (
    SR_Service_RecID int PRIMARY KEY,
    Summary nvarchar(100),
    Date_Entered datetime,
    Last_Update datetime,
    Company_RecID int,
    Company_Name nvarchar(250),
    Contact_RecID int,
    Contact_Name nvarchar(100),
    SR_Type_Description nvarchar(50),
    SR_SubType_Description nvarchar(50),
    SR_Status_Description nvarchar(50)
);

-- Clear existing data
DELETE FROM SR_Event WHERE Company_RecID = $CompanyRecID;

-- Insert events
INSERT INTO SR_Event
SELECT 
    s.SR_Service_RecID,
    s.Summary,
    s.Date_Entered,
    s.Last_Update,
    s.Company_RecID,
    c.Company_Name,
    s.Contact_RecID,
    CONCAT(cont.First_Name, ' ', cont.Last_Name) AS Contact_Name,
    st.Description AS SR_Type_Description,
    sst.Description AS SR_SubType_Description,
    ss.Description AS SR_Status_Description
FROM CW_Report_DB.cwwebapp_mit.dbo.SR_Service s
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Company c ON s.Company_RecID = c.Company_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Contact cont ON s.Contact_RecID = cont.Contact_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Type st ON s.SR_Type_RecID = st.SR_Type_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_SubType sst ON s.SR_SubType_RecID = sst.SR_SubType_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Status ss ON s.SR_Status_RecID = ss.SR_Status_RecID
WHERE s.Company_RecID = $CompanyRecID
    AND s.SR_Type_RecID = 287;
"@

    return Invoke-SafeSqlQuery -Query $Query -Operation "Export Events"
}

function Export-Incidents {
    param([int]$CompanyRecID)
    
    $Query = @"
USE [$Database];

-- Create table if not exists
IF OBJECT_ID('SR_Incident') IS NULL
CREATE TABLE SR_Incident (
    SR_Service_RecID int PRIMARY KEY,
    Summary nvarchar(100),
    Date_Entered datetime,
    Last_Update datetime,
    Company_RecID int,
    Company_Name nvarchar(250),
    Contact_RecID int,
    Contact_Name nvarchar(100),
    SR_Type_Description nvarchar(50),
    SR_SubType_Description nvarchar(50),
    SR_Status_Description nvarchar(50)
);

-- Clear existing data
DELETE FROM SR_Incident WHERE Company_RecID = $CompanyRecID;

-- Insert incidents
INSERT INTO SR_Incident
SELECT 
    s.SR_Service_RecID,
    s.Summary,
    s.Date_Entered,
    s.Last_Update,
    s.Company_RecID,
    c.Company_Name,
    s.Contact_RecID,
    CONCAT(cont.First_Name, ' ', cont.Last_Name) AS Contact_Name,
    st.Description AS SR_Type_Description,
    sst.Description AS SR_SubType_Description,
    ss.Description AS SR_Status_Description
FROM CW_Report_DB.cwwebapp_mit.dbo.SR_Service s
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Company c ON s.Company_RecID = c.Company_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Contact cont ON s.Contact_RecID = cont.Contact_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Type st ON s.SR_Type_RecID = st.SR_Type_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_SubType sst ON s.SR_SubType_RecID = sst.SR_SubType_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Status ss ON s.SR_Status_RecID = ss.SR_Status_RecID
WHERE s.Company_RecID = $CompanyRecID
    AND s.SR_Type_RecID = 346;
"@

    return Invoke-SafeSqlQuery -Query $Query -Operation "Export Incidents"
}

function Export-Problems {
    param([int]$CompanyRecID)
    
    $Query = @"
USE [$Database];

-- Create table if not exists
IF OBJECT_ID('SR_Problem') IS NULL
CREATE TABLE SR_Problem (
    SR_Service_RecID int PRIMARY KEY,
    Summary nvarchar(100),
    Date_Entered datetime,
    Last_Update datetime,
    Company_RecID int,
    Company_Name nvarchar(250),
    Contact_RecID int,
    Contact_Name nvarchar(100),
    SR_Type_Description nvarchar(50),
    SR_SubType_Description nvarchar(50),
    SR_Status_Description nvarchar(50)
);

-- Clear existing data
DELETE FROM SR_Problem WHERE Company_RecID = $CompanyRecID;

-- Insert problems
INSERT INTO SR_Problem
SELECT 
    s.SR_Service_RecID,
    s.Summary,
    s.Date_Entered,
    s.Last_Update,
    s.Company_RecID,
    c.Company_Name,
    s.Contact_RecID,
    CONCAT(cont.First_Name, ' ', cont.Last_Name) AS Contact_Name,
    st.Description AS SR_Type_Description,
    sst.Description AS SR_SubType_Description,
    ss.Description AS SR_Status_Description
FROM CW_Report_DB.cwwebapp_mit.dbo.SR_Service s
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Company c ON s.Company_RecID = c.Company_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Contact cont ON s.Contact_RecID = cont.Contact_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Type st ON s.SR_Type_RecID = st.SR_Type_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_SubType sst ON s.SR_SubType_RecID = sst.SR_SubType_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Status ss ON s.SR_Status_RecID = ss.SR_Status_RecID
WHERE s.Company_RecID = $CompanyRecID
    AND s.SR_Type_RecID = 281;
"@

    return Invoke-SafeSqlQuery -Query $Query -Operation "Export Problems"
}

function Export-Requests {
    param([int]$CompanyRecID)
    
    $Query = @"
USE [$Database];

-- Create table if not exists
IF OBJECT_ID('SR_Request') IS NULL
CREATE TABLE SR_Request (
    SR_Service_RecID int PRIMARY KEY,
    Summary nvarchar(100),
    Date_Entered datetime,
    Last_Update datetime,
    Company_RecID int,
    Company_Name nvarchar(250),
    Contact_RecID int,
    Contact_Name nvarchar(100),
    SR_Type_Description nvarchar(50),
    SR_SubType_Description nvarchar(50),
    SR_Status_Description nvarchar(50)
);

-- Clear existing data
DELETE FROM SR_Request WHERE Company_RecID = $CompanyRecID;

-- Insert requests
INSERT INTO SR_Request
SELECT 
    s.SR_Service_RecID,
    s.Summary,
    s.Date_Entered,
    s.Last_Update,
    s.Company_RecID,
    c.Company_Name,
    s.Contact_RecID,
    CONCAT(cont.First_Name, ' ', cont.Last_Name) AS Contact_Name,
    st.Description AS SR_Type_Description,
    sst.Description AS SR_SubType_Description,
    ss.Description AS SR_Status_Description
FROM CW_Report_DB.cwwebapp_mit.dbo.SR_Service s
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Company c ON s.Company_RecID = c.Company_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Contact cont ON s.Contact_RecID = cont.Contact_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Type st ON s.SR_Type_RecID = st.SR_Type_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_SubType sst ON s.SR_SubType_RecID = sst.SR_SubType_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Status ss ON s.SR_Status_RecID = ss.SR_Status_RecID
WHERE s.Company_RecID = $CompanyRecID
    AND s.SR_Type_RecID NOT IN (287, 346, 281);
"@

    return Invoke-SafeSqlQuery -Query $Query -Operation "Export Requests"
}

function Export-Notes-Chunked {
    param(
        [int]$CompanyRecID,
        [int]$ChunkSize
    )
    
    Write-Log "Starting chunked notes export with chunk size: $ChunkSize"
    
    # Create table if not exists
    $CreateTableQuery = @"
USE [$Database];

IF OBJECT_ID('SR_Notes') IS NULL
CREATE TABLE SR_Notes (
    SR_Detail_RecID int PRIMARY KEY,
    SR_Service_RecID int,
    Ticket_Summary nvarchar(100),
    Note_Content nvarchar(max),
    Date_Created datetime,
    Last_Update datetime,
    Contact_Name nvarchar(100),
    Company_Name nvarchar(250)
);

-- Clear existing data
DELETE FROM SR_Notes WHERE SR_Service_RecID IN (
    SELECT SR_Service_RecID FROM CW_Report_DB.cwwebapp_mit.dbo.SR_Service WHERE Company_RecID = $CompanyRecID
);
"@

    $CreateResult = Invoke-SafeSqlQuery -Query $CreateTableQuery -Operation "Create Notes Table"
    if (-not $CreateResult.Success) {
        return $CreateResult
    }

    # Get total count for progress tracking
    $CountQuery = @"
SELECT COUNT(*) AS TotalCount
FROM CW_Report_DB.cwwebapp_mit.dbo.SR_Detail sd
JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Service s ON sd.SR_Service_RecID = s.SR_Service_RecID
WHERE s.Company_RecID = $CompanyRecID
    AND (sd.InternalAnalysis_Flag = 0 OR sd.InternalAnalysis_Flag IS NULL);
"@

    $CountResult = Invoke-SafeSqlQuery -Query $CountQuery -Operation "Get Notes Count"
    if (-not $CountResult.Success) {
        return $CountResult
    }

    $TotalRecords = $CountResult.Result.TotalCount
    Write-Log "Total notes records to process: $TotalRecords"

    # Create pagination table with row numbers
    $PrepQuery = @"
USE [$Database];
IF OBJECT_ID('notes_pagination') IS NOT NULL DROP TABLE notes_pagination;
SELECT sd.SR_Detail_RecID, sd.Date_Created,
       ROW_NUMBER() OVER (ORDER BY sd.Date_Created, sd.SR_Detail_RecID) AS rn
INTO notes_pagination
FROM CW_Report_DB.cwwebapp_mit.dbo.SR_Detail sd
JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Service s ON s.SR_Service_RecID = sd.SR_Service_RecID
WHERE s.Company_RecID = $CompanyRecID
    AND (sd.InternalAnalysis_Flag = 0 OR sd.InternalAnalysis_Flag IS NULL);
"@

    $PrepResult = Invoke-SafeSqlQuery -Query $PrepQuery -Operation "Prepare Notes Pagination"
    if (-not $PrepResult.Success) {
        return $PrepResult
    }

    # Process in chunks
    $ProcessedRecords = 0
    $FailedChunks = 0
    
    for ($Start = 1; $Start -le $TotalRecords; $Start += $ChunkSize) {
        $End = [math]::Min($Start + $ChunkSize - 1, $TotalRecords)
        
        Write-Progress-Log -Operation "Export Notes" -Current $ProcessedRecords -Total $TotalRecords -Status "Processing chunk $Start-$End"
        
        $ChunkQuery = @"
INSERT INTO SR_Notes (SR_Detail_RecID, SR_Service_RecID, Ticket_Summary, Note_Content, 
    Date_Created, Last_Update, Contact_Name, Company_Name)
SELECT
    sd.SR_Detail_RecID, sd.SR_Service_RecID, s.Summary, sd.SR_Detail_Notes,
    sd.Date_Created, sd.Last_Update,
    CONCAT(cont.First_Name, ' ', cont.Last_Name) AS Contact_Name,
    c.Company_Name
FROM notes_pagination r
JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Detail sd ON sd.SR_Detail_RecID = r.SR_Detail_RecID
JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Service s ON sd.SR_Service_RecID = s.SR_Service_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Type st ON s.SR_Type_RecID = st.SR_Type_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Company c ON s.Company_RecID = c.Company_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Contact cont ON sd.Contact_RecID = cont.Contact_RecID
WHERE r.rn BETWEEN $Start AND $End;
"@

        $ChunkResult = Invoke-SafeSqlQuery -Query $ChunkQuery -Operation "Export Notes Chunk $Start-$End"
        
        if ($ChunkResult.Success) {
            $ProcessedRecords += ($End - $Start + 1)
        } else {
            $FailedChunks++
            Write-Log "Failed to process notes chunk $Start-$End" "ERROR"
        }
        
        Start-Sleep -Seconds 1
    }

    Write-Progress -Activity "Export Notes" -Completed
    
    # Clean up pagination table
    $CleanupQuery = @"
USE [$Database];
IF OBJECT_ID('notes_pagination') IS NOT NULL DROP TABLE notes_pagination;
"@
    Invoke-SafeSqlQuery -Query $CleanupQuery -Operation "Cleanup Notes Pagination Table" | Out-Null
    
    Write-Log "Notes export completed. Processed: $ProcessedRecords, Failed chunks: $FailedChunks" "SUCCESS"
    
    return @{ Success = $true; ProcessedRecords = $ProcessedRecords; FailedChunks = $FailedChunks }
}

function Export-TimeEntries-Chunked {
    param(
        [int]$CompanyRecID,
        [int]$ChunkSize
    )
    
    Write-Log "Starting chunked time entries export with chunk size: $ChunkSize"
    
    # Create table if not exists
    $CreateTableQuery = @"
USE [$Database];

IF OBJECT_ID('SR_Time_Entries') IS NULL
CREATE TABLE SR_Time_Entries (
    Time_RecID int PRIMARY KEY,
    SR_Service_RecID int,
    Ticket_Summary nvarchar(100),
    Date_Start datetime,
    Notes nvarchar(max),
    Last_Update datetime,
    Company_RecID int,
    Company_Name nvarchar(250),
    Activity_Class_RecID int,
    Activity_Class_Description nvarchar(50),
    Activity_Type_RecID int,
    Activity_Type_Description nvarchar(50)
);

-- Clear existing data
DELETE FROM SR_Time_Entries WHERE SR_Service_RecID IN (
    SELECT SR_Service_RecID FROM CW_Report_DB.cwwebapp_mit.dbo.SR_Service WHERE Company_RecID = $CompanyRecID
);
"@

    $CreateResult = Invoke-SafeSqlQuery -Query $CreateTableQuery -Operation "Create Time Entries Table"
    if (-not $CreateResult.Success) {
        return $CreateResult
    }

    # Get total count for progress tracking
    $CountQuery = @"
SELECT COUNT(*) AS TotalCount
FROM CW_Report_DB.cwwebapp_mit.dbo.Time_Entry te
WHERE te.Company_RecID = $CompanyRecID;
"@

    $CountResult = Invoke-SafeSqlQuery -Query $CountQuery -Operation "Get Time Entries Count"
    if (-not $CountResult.Success) {
        return $CountResult
    }

    $TotalRecords = $CountResult.Result.TotalCount
    Write-Log "Total time entries records to process: $TotalRecords"

    # Create pagination table with row numbers
    $PrepQuery = @"
USE [$Database];
IF OBJECT_ID('time_entries_pagination') IS NOT NULL DROP TABLE time_entries_pagination;
SELECT te.Time_RecID, te.Date_Start,
       ROW_NUMBER() OVER (ORDER BY te.Date_Start, te.Time_RecID) AS rn
INTO time_entries_pagination
FROM CW_Report_DB.cwwebapp_mit.dbo.Time_Entry te
WHERE te.Company_RecID = $CompanyRecID;
"@

    $PrepResult = Invoke-SafeSqlQuery -Query $PrepQuery -Operation "Prepare Time Entries Pagination"
    if (-not $PrepResult.Success) {
        return $PrepResult
    }

    # Process in chunks
    $ProcessedRecords = 0
    $FailedChunks = 0
    
    for ($Start = 1; $Start -le $TotalRecords; $Start += $ChunkSize) {
        $End = [math]::Min($Start + $ChunkSize - 1, $TotalRecords)
        
        Write-Progress-Log -Operation "Export Time Entries" -Current $ProcessedRecords -Total $TotalRecords -Status "Processing chunk $Start-$End"
        
        $ChunkQuery = @"
INSERT INTO SR_Time_Entries (Time_RecID, SR_Service_RecID, Ticket_Summary, Date_Start,
    Notes, Last_Update, Company_RecID, Company_Name, 
    Activity_Class_RecID, Activity_Class_Description, Activity_Type_RecID, Activity_Type_Description)
SELECT 
    te.Time_RecID, te.SR_Service_RecID, s.Summary, te.Date_Start,
    te.Notes, te.Last_Update,
    te.Company_RecID, c.Company_Name,
    te.Activity_Class_RecID, ac.Description AS Activity_Class_Description,
    te.Activity_Type_RecID, at.Description AS Activity_Type_Description
FROM time_entries_pagination r
JOIN CW_Report_DB.cwwebapp_mit.dbo.Time_Entry te ON te.Time_RecID = r.Time_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Service s ON te.SR_Service_RecID = s.SR_Service_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Activity_Class ac ON te.Activity_Class_RecID = ac.Activity_Class_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Activity_Type at ON te.Activity_Type_RecID = at.Activity_Type_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Company c ON te.Company_RecID = c.Company_RecID
WHERE r.rn BETWEEN $Start AND $End;
"@

        $ChunkResult = Invoke-SafeSqlQuery -Query $ChunkQuery -Operation "Export Time Entries Chunk $Start-$End"
        
        if ($ChunkResult.Success) {
            $ProcessedRecords += ($End - $Start + 1)
        } else {
            $FailedChunks++
            Write-Log "Failed to process time entries chunk $Start-$End" "ERROR"
        }
        
        Start-Sleep -Seconds 1
    }

    Write-Progress -Activity "Export Time Entries" -Completed
    
    # Clean up pagination table
    $CleanupQuery = @"
USE [$Database];
IF OBJECT_ID('time_entries_pagination') IS NOT NULL DROP TABLE time_entries_pagination;
"@
    Invoke-SafeSqlQuery -Query $CleanupQuery -Operation "Cleanup Time Entries Pagination Table" | Out-Null
    
    Write-Log "Time entries export completed. Processed: $ProcessedRecords, Failed chunks: $FailedChunks" "SUCCESS"
    
    return @{ Success = $true; ProcessedRecords = $ProcessedRecords; FailedChunks = $FailedChunks }
}

try {
    Write-Log "=========================================="
    Write-Log "Starting CQL Service Tickets Export Process (Embedded)"
    Write-Log "=========================================="
    Write-Log "Server: $ServerInstance, Database: $Database"
    Write-Log "Company RecID: $CompanyRecID"
    Write-Log "Notes Chunk Size: $NotesChunkSize"
    Write-Log "Time Entries Chunk Size: $TimeEntriesChunkSize"
    Write-Log "Delay between operations: $DelayBetweenOperations seconds"
    Write-Log "Log Path: $LogPath"
    
    # Test SQL connection
    Write-Log "Testing SQL connection..."
    if (-not (Test-SqlConnection -ServerInstance $ServerInstance -Database $Database)) {
        Write-Log "Cannot connect to SQL Server instance '$ServerInstance' database '$Database'" "ERROR"
        Write-Log "Please verify:"
        Write-Log "  1. SQL Server is running"
        Write-Log "  2. Database '$Database' exists or can be created"
        Write-Log "  3. You have appropriate permissions"
        Write-Log "  4. SqlServer PowerShell module is installed"
        exit 1
    }
    Write-Log "SQL connection successful" "SUCCESS"
    
    # Define operations in execution order
    $Operations = @(
        @{ Name = "Export Config Items"; Function = { Export-ConfigItems -CompanyRecID $CompanyRecID } },
        @{ Name = "Export Events"; Function = { Export-Events -CompanyRecID $CompanyRecID } },
        @{ Name = "Export Incidents"; Function = { Export-Incidents -CompanyRecID $CompanyRecID } },
        @{ Name = "Export Problems"; Function = { Export-Problems -CompanyRecID $CompanyRecID } },
        @{ Name = "Export Requests"; Function = { Export-Requests -CompanyRecID $CompanyRecID } },
        @{ Name = "Export Notes (Chunked)"; Function = { Export-Notes-Chunked -CompanyRecID $CompanyRecID -ChunkSize $NotesChunkSize } },
        @{ Name = "Export Time Entries (Chunked)"; Function = { Export-TimeEntries-Chunked -CompanyRecID $CompanyRecID -ChunkSize $TimeEntriesChunkSize } }
    )
    
    $SuccessCount = 0
    $FailedOperations = @()
    
    foreach ($Operation in $Operations) {
        Write-Log "=========================================="
        Write-Log "Executing: $($Operation.Name)"
        Write-Log "=========================================="
        
        $Result = & $Operation.Function
        
        if ($Result.Success) {
            $SuccessCount++
            Write-Log "$($Operation.Name) completed successfully" "SUCCESS"
        } else {
            $FailedOperations += $Operation
            Write-Log "$($Operation.Name) failed" "ERROR"
            
            # Ask user if they want to continue
            $Continue = Read-Host "Operation failed. Continue with remaining operations? (y/n)"
            if ($Continue -notmatch '^[yY]') {
                Write-Log "Export process stopped by user" "WARNING"
                break
            }
        }
        
        # Add delay between operations (except for the last one)
        if ($Operation -ne $Operations[-1]) {
            Write-Log "Waiting $DelayBetweenOperations seconds before next operation..."
            Start-Sleep -Seconds $DelayBetweenOperations
        }
    }
    
    # Final summary
    Write-Log "=========================================="
    Write-Log "CQL EXPORT PROCESS COMPLETED (Embedded)"
    Write-Log "=========================================="
    Write-Log "Total operations: $($Operations.Count)"
    Write-Log "Successful: $SuccessCount"
    Write-Log "Failed: $($FailedOperations.Count)"
    
    if ($FailedOperations.Count -gt 0) {
        Write-Log "Failed operations:" "WARNING"
        foreach ($Failed in $FailedOperations) {
            Write-Log "  - $($Failed.Name)" "WARNING"
        }
        Write-Log "You can re-run the script to retry failed operations"
    } else {
        Write-Log "ALL OPERATIONS COMPLETED SUCCESSFULLY!" "SUCCESS"
    }
    
    Write-Log "Log file saved to: $LogPath"
    
} catch {
    Write-Log "Unexpected error occurred: $($_.Exception.Message)" "ERROR"
    exit 1
}
#endregion
