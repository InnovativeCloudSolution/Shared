#Requires -Module SqlServer

<#
.SYNOPSIS
    Test script to export the 20 latest tickets per type with all associated data

.DESCRIPTION
    Exports the 20 most recent tickets for each type (Events, Incidents, Problems, Requests)
    along with all associated Config Items, Notes, and Time Entries for testing purposes.

.PARAMETER ServerInstance
    SQL Server instance name (default: .\SQLEXPRESS)

.PARAMETER Database
    Target database name (default: CQL_ServiceTickets_Test)

.PARAMETER CompanyRecID
    Company RecID to filter by (default: 2791)

.PARAMETER LogPath
    Path for log file (default: script directory)
#>

[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()][string]$ServerInstance = ".\SQLEXPRESS",
    [ValidateNotNullOrEmpty()][string]$Database = "CQL_ServiceTickets_Test",
    [int]$CompanyRecID = 2791,
    [string]$LogPath = "$PSScriptRoot\CQL-Test-Log-$(Get-Date -Format 'yyyyMMdd-HHmmss').txt"
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

#region Test Export Functions
function Export-TestTickets {
    param([int]$CompanyRecID)
    
    $Query = @"
USE [$Database];

-- Create Latest Tickets table
IF OBJECT_ID('Test_Latest_Tickets') IS NOT NULL
    DROP TABLE Test_Latest_Tickets;

CREATE TABLE Test_Latest_Tickets (
    SR_Service_RecID int PRIMARY KEY,
    Ticket_Type nvarchar(20),
    Summary nvarchar(100),
    Date_Entered datetime,
    Entered_By nvarchar(100),
    Last_Update datetime,
    Updated_By nvarchar(100),
    Closed_By nvarchar(100),
    Company_RecID int,
    Company_Name nvarchar(250),
    Contact_RecID int,
    Contact_Name nvarchar(100),
    SR_Type_RecID int,
    SR_Type_Description nvarchar(50),
    SR_SubType_RecID int,
    SR_SubType_Description nvarchar(50),
    SR_Status_RecID int,
    SR_Status_Description nvarchar(50)
);

-- Insert 20 latest Events
INSERT INTO Test_Latest_Tickets
SELECT TOP 20
    s.SR_Service_RecID,
    'Event' AS Ticket_Type,
    s.Summary,
    s.Date_Entered,
    s.Entered_By,
    s.Last_Update,
    s.Updated_By,
    s.Closed_By,
    s.Company_RecID,
    c.Company_Name,
    s.Contact_RecID,
    CONCAT(cont.First_Name, ' ', cont.Last_Name) AS Contact_Name,
    s.SR_Type_RecID,
    st.Description AS SR_Type_Description,
    s.SR_SubType_RecID,
    sst.Description AS SR_SubType_Description,
    s.SR_Status_RecID,
    ss.Description AS SR_Status_Description
FROM CW_Report_DB.cwwebapp_mit.dbo.SR_Service s
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Company c ON s.Company_RecID = c.Company_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Contact cont ON s.Contact_RecID = cont.Contact_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Type st ON s.SR_Type_RecID = st.SR_Type_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_SubType sst ON s.SR_SubType_RecID = sst.SR_SubType_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Status ss ON s.SR_Status_RecID = ss.SR_Status_RecID
WHERE s.Company_RecID = $CompanyRecID
    AND s.SR_Type_RecID = 287
ORDER BY s.Date_Entered DESC;

-- Insert 20 latest Incidents
INSERT INTO Test_Latest_Tickets
SELECT TOP 20
    s.SR_Service_RecID,
    'Incident' AS Ticket_Type,
    s.Summary,
    s.Date_Entered,
    s.Entered_By,
    s.Last_Update,
    s.Updated_By,
    s.Closed_By,
    s.Company_RecID,
    c.Company_Name,
    s.Contact_RecID,
    CONCAT(cont.First_Name, ' ', cont.Last_Name) AS Contact_Name,
    s.SR_Type_RecID,
    st.Description AS SR_Type_Description,
    s.SR_SubType_RecID,
    sst.Description AS SR_SubType_Description,
    s.SR_Status_RecID,
    ss.Description AS SR_Status_Description
FROM CW_Report_DB.cwwebapp_mit.dbo.SR_Service s
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Company c ON s.Company_RecID = c.Company_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Contact cont ON s.Contact_RecID = cont.Contact_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Type st ON s.SR_Type_RecID = st.SR_Type_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_SubType sst ON s.SR_SubType_RecID = sst.SR_SubType_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Status ss ON s.SR_Status_RecID = ss.SR_Status_RecID
WHERE s.Company_RecID = $CompanyRecID
    AND s.SR_Type_RecID = 346
ORDER BY s.Date_Entered DESC;

-- Insert 20 latest Problems
INSERT INTO Test_Latest_Tickets
SELECT TOP 20
    s.SR_Service_RecID,
    'Problem' AS Ticket_Type,
    s.Summary,
    s.Date_Entered,
    s.Entered_By,
    s.Last_Update,
    s.Updated_By,
    s.Closed_By,
    s.Company_RecID,
    c.Company_Name,
    s.Contact_RecID,
    CONCAT(cont.First_Name, ' ', cont.Last_Name) AS Contact_Name,
    s.SR_Type_RecID,
    st.Description AS SR_Type_Description,
    s.SR_SubType_RecID,
    sst.Description AS SR_SubType_Description,
    s.SR_Status_RecID,
    ss.Description AS SR_Status_Description
FROM CW_Report_DB.cwwebapp_mit.dbo.SR_Service s
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Company c ON s.Company_RecID = c.Company_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Contact cont ON s.Contact_RecID = cont.Contact_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Type st ON s.SR_Type_RecID = st.SR_Type_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_SubType sst ON s.SR_SubType_RecID = sst.SR_SubType_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Status ss ON s.SR_Status_RecID = ss.SR_Status_RecID
WHERE s.Company_RecID = $CompanyRecID
    AND s.SR_Type_RecID = 281
ORDER BY s.Date_Entered DESC;

-- Insert 20 latest Requests
INSERT INTO Test_Latest_Tickets
SELECT TOP 20
    s.SR_Service_RecID,
    'Request' AS Ticket_Type,
    s.Summary,
    s.Date_Entered,
    s.Entered_By,
    s.Last_Update,
    s.Updated_By,
    s.Closed_By,
    s.Company_RecID,
    c.Company_Name,
    s.Contact_RecID,
    CONCAT(cont.First_Name, ' ', cont.Last_Name) AS Contact_Name,
    s.SR_Type_RecID,
    st.Description AS SR_Type_Description,
    s.SR_SubType_RecID,
    sst.Description AS SR_SubType_Description,
    s.SR_Status_RecID,
    ss.Description AS SR_Status_Description
FROM CW_Report_DB.cwwebapp_mit.dbo.SR_Service s
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Company c ON s.Company_RecID = c.Company_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Contact cont ON s.Contact_RecID = cont.Contact_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Type st ON s.SR_Type_RecID = st.SR_Type_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_SubType sst ON s.SR_SubType_RecID = sst.SR_SubType_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Status ss ON s.SR_Status_RecID = ss.SR_Status_RecID
WHERE s.Company_RecID = $CompanyRecID
    AND s.SR_Type_RecID NOT IN (287, 346, 281)
ORDER BY s.Date_Entered DESC;
"@

    return Invoke-SafeSqlQuery -Query $Query -Operation "Export Test Tickets"
}

function Export-TestConfigItems {
    param([int]$CompanyRecID)
    
    $Query = @"
USE [$Database];

-- Create Test Config Items table
IF OBJECT_ID('Test_Config_Items') IS NOT NULL
    DROP TABLE Test_Config_Items;

CREATE TABLE Test_Config_Items (
    Config_RecID int PRIMARY KEY,
    Config_Name nvarchar(100),
    Config_Type nvarchar(50),
    Company_RecID int,
    Company_Name nvarchar(250),
    Serial_Number nvarchar(100),
    Vendor_Name nvarchar(100),
    Status nvarchar(50)
);

-- Insert all configs associated with test tickets
INSERT INTO Test_Config_Items
SELECT DISTINCT
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
INNER JOIN CW_Report_DB.cwwebapp_mit.dbo.SR_Config sc ON ci.Config_RecID = sc.Config_RecID
INNER JOIN Test_Latest_Tickets t ON sc.SR_Service_RecID = t.SR_Service_RecID
WHERE ci.Company_RecID = $CompanyRecID;
"@

    return Invoke-SafeSqlQuery -Query $Query -Operation "Export Test Config Items"
}

function Export-TestNotes {
    param([int]$CompanyRecID)
    
    $Query = @"
USE [$Database];

-- Create Test Notes table
IF OBJECT_ID('Test_Notes') IS NOT NULL
    DROP TABLE Test_Notes;

CREATE TABLE Test_Notes (
    SR_Detail_RecID int PRIMARY KEY,
    SR_Service_RecID int,
    Ticket_Type nvarchar(20),
    Ticket_Summary nvarchar(100),
    Note_Content nvarchar(max),
    Date_Created datetime,
    Created_By nvarchar(100),
    Last_Update datetime,
    Updated_By nvarchar(100),
    Internal_Member_Flag bit,
    InternalAnalysis_Flag bit,
    Member_Name nvarchar(100),
    Contact_Name nvarchar(100),
    Company_Name nvarchar(250)
);

-- Insert all notes for test tickets
INSERT INTO Test_Notes
SELECT
    sd.SR_Detail_RecID,
    sd.SR_Service_RecID,
    t.Ticket_Type,
    t.Summary AS Ticket_Summary,
    sd.SR_Detail_Notes AS Note_Content,
    sd.Date_Created,
    sd.Created_By,
    sd.Last_Update,
    sd.Updated_By,
    sd.Internal_Member_Flag,
    sd.InternalAnalysis_Flag,
    CONCAT(m.First_Name, ' ', m.Last_Name) AS Member_Name,
    CONCAT(cont.First_Name, ' ', cont.Last_Name) AS Contact_Name,
    t.Company_Name
FROM CW_Report_DB.cwwebapp_mit.dbo.SR_Detail sd
INNER JOIN Test_Latest_Tickets t ON sd.SR_Service_RecID = t.SR_Service_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Member m ON sd.Member_RecID = m.Member_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Contact cont ON sd.Contact_RecID = cont.Contact_RecID
WHERE (sd.InternalAnalysis_Flag = 0 OR sd.InternalAnalysis_Flag IS NULL)
ORDER BY sd.Date_Created DESC;
"@

    return Invoke-SafeSqlQuery -Query $Query -Operation "Export Test Notes"
}

function Export-TestTimeEntries {
    param([int]$CompanyRecID)
    
    $Query = @"
USE [$Database];

-- Create Test Time Entries table
IF OBJECT_ID('Test_Time_Entries') IS NOT NULL
    DROP TABLE Test_Time_Entries;

CREATE TABLE Test_Time_Entries (
    Time_RecID int PRIMARY KEY,
    SR_Service_RecID int,
    Ticket_Type nvarchar(20),
    Ticket_Summary nvarchar(100),
    Date_Start datetime,
    Hours_Actual decimal(6,2),
    Hours_Bill decimal(6,2),
    Notes nvarchar(max),
    Entered_By nvarchar(100),
    Last_Update datetime,
    Updated_By nvarchar(100),
    Member_RecID int,
    Member_Name nvarchar(100),
    Company_RecID int,
    Company_Name nvarchar(250),
    Activity_Class_RecID int,
    Activity_Class_Description nvarchar(50),
    Activity_Type_RecID int,
    Activity_Type_Description nvarchar(50)
);

-- Insert all time entries for test tickets
INSERT INTO Test_Time_Entries
SELECT
    te.Time_RecID,
    te.SR_Service_RecID,
    t.Ticket_Type,
    t.Summary AS Ticket_Summary,
    te.Date_Start,
    te.Hours_Actual,
    te.Hours_Bill,
    te.Notes,
    te.Entered_By,
    te.Last_Update,
    te.Updated_By,
    te.Member_RecID,
    CONCAT(m.First_Name, ' ', m.Last_Name) AS Member_Name,
    te.Company_RecID,
    t.Company_Name,
    te.Activity_Class_RecID,
    ac.Description AS Activity_Class_Description,
    te.Activity_Type_RecID,
    at.Description AS Activity_Type_Description
FROM CW_Report_DB.cwwebapp_mit.dbo.Time_Entry te
INNER JOIN Test_Latest_Tickets t ON te.SR_Service_RecID = t.SR_Service_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Member m ON te.Member_RecID = m.Member_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Activity_Class ac ON te.Activity_Class_RecID = ac.Activity_Class_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Activity_Type at ON te.Activity_Type_RecID = at.Activity_Type_RecID
ORDER BY te.Date_Start DESC;
"@

    return Invoke-SafeSqlQuery -Query $Query -Operation "Export Test Time Entries"
}

function Export-TestConfigLinks {
    param([int]$CompanyRecID)
    
    $Query = @"
USE [$Database];

-- Create Test Config Links table
IF OBJECT_ID('Test_Config_Links') IS NOT NULL
    DROP TABLE Test_Config_Links;

CREATE TABLE Test_Config_Links (
    Link_RecID int IDENTITY(1,1) PRIMARY KEY,
    SR_Service_RecID int,
    Config_RecID int,
    Ticket_Type nvarchar(20),
    Ticket_Summary nvarchar(100),
    Config_Name nvarchar(100),
    Config_Type nvarchar(50),
    Link_Type nvarchar(50),
    Company_Name nvarchar(250)
);

-- Insert config-to-ticket relationships for test tickets
INSERT INTO Test_Config_Links (SR_Service_RecID, Config_RecID, Ticket_Type, Ticket_Summary, Config_Name, Config_Type, Link_Type, Company_Name)
SELECT
    sc.SR_Service_RecID,
    sc.Config_RecID,
    t.Ticket_Type,
    t.Summary AS Ticket_Summary,
    ci.Config_Name,
    'Configuration' AS Config_Type,
    'Service-Config Link' AS Link_Type,
    t.Company_Name
FROM CW_Report_DB.cwwebapp_mit.dbo.SR_Config sc
INNER JOIN Test_Latest_Tickets t ON sc.SR_Service_RecID = t.SR_Service_RecID
LEFT JOIN CW_Report_DB.cwwebapp_mit.dbo.Config ci ON sc.Config_RecID = ci.Config_RecID
WHERE sc.Config_RecID IS NOT NULL;
"@

    return Invoke-SafeSqlQuery -Query $Query -Operation "Export Test Config Links"
}

function Generate-TestSummary {
    param([int]$CompanyRecID)
    
    $Query = @"
USE [$Database];

-- Generate test summary report
SELECT 
    'TICKETS' AS Report_Section,
    Ticket_Type AS Item_Type,
    COUNT(*) AS Total_Count,
    MIN(Date_Entered) AS Earliest_Date,
    MAX(Date_Entered) AS Latest_Date
FROM Test_Latest_Tickets
GROUP BY Ticket_Type

UNION ALL

SELECT 
    'CONFIG_ITEMS' AS Report_Section,
    'Total' AS Item_Type,
    COUNT(*) AS Total_Count,
    NULL AS Earliest_Date,
    NULL AS Latest_Date
FROM Test_Config_Items

UNION ALL

SELECT 
    'NOTES' AS Report_Section,
    'Total' AS Item_Type,
    COUNT(*) AS Total_Count,
    MIN(Date_Created) AS Earliest_Date,
    MAX(Date_Created) AS Latest_Date
FROM Test_Notes

UNION ALL

SELECT 
    'TIME_ENTRIES' AS Report_Section,
    'Total' AS Item_Type,
    COUNT(*) AS Total_Count,
    MIN(Date_Start) AS Earliest_Date,
    MAX(Date_Start) AS Latest_Date
FROM Test_Time_Entries

UNION ALL

SELECT 
    'CONFIG_LINKS' AS Report_Section,
    'Total' AS Item_Type,
    COUNT(*) AS Total_Count,
    NULL AS Earliest_Date,
    NULL AS Latest_Date
FROM Test_Config_Links

ORDER BY Report_Section, Item_Type;
"@

    return Invoke-SafeSqlQuery -Query $Query -Operation "Generate Test Summary"
}
#endregion

#region Main Execution
try {
    Write-Log "=========================================="
    Write-Log "Starting CQL Test Export - Latest 20 Tickets Per Type"
    Write-Log "=========================================="
    Write-Log "Server: $ServerInstance, Database: $Database"
    Write-Log "Company RecID: $CompanyRecID"
    Write-Log "Log Path: $LogPath"
    
    # Test SQL connection
    Write-Log "Testing SQL connection..."
    $ConnectionTest = Test-SqlConnection -ServerInstance $ServerInstance -Database $Database
    if (-not $ConnectionTest) {
        Write-Log "Connection test failed. Exiting." "ERROR"
        exit 1
    }
    Write-Log "SQL connection successful" "SUCCESS"
    
    # Define operations in execution order
    $Operations = @(
        @{ Name = "Export Test Tickets (20 per type)"; Function = { Export-TestTickets -CompanyRecID $CompanyRecID } },
        @{ Name = "Export Associated Config Items"; Function = { Export-TestConfigItems -CompanyRecID $CompanyRecID } },
        @{ Name = "Export Associated Notes"; Function = { Export-TestNotes -CompanyRecID $CompanyRecID } },
        @{ Name = "Export Associated Time Entries"; Function = { Export-TestTimeEntries -CompanyRecID $CompanyRecID } },
        @{ Name = "Export Config-to-Ticket Links"; Function = { Export-TestConfigLinks -CompanyRecID $CompanyRecID } },
        @{ Name = "Generate Test Summary Report"; Function = { Generate-TestSummary -CompanyRecID $CompanyRecID } }
    )
    
    $SuccessCount = 0
    $FailedOperations = @()
    
    foreach ($operation in $Operations) {
        Write-Log "=========================================="
        Write-Log "Executing: $($operation.Name)"
        Write-Log "=========================================="
        
        $result = & $operation.Function
        
        if ($result.Success) {
            $SuccessCount++
            Write-Log "$($operation.Name) completed successfully" "SUCCESS"
        } else {
            $FailedOperations += $operation.Name
            Write-Log "$($operation.Name) failed" "ERROR"
        }
        
        # Small delay between operations
        Start-Sleep -Seconds 2
    }
    
    Write-Log "=========================================="
    Write-Log "CQL TEST EXPORT COMPLETED"
    Write-Log "=========================================="
    Write-Log "Total operations: $($Operations.Count)"
    Write-Log "Successful: $SuccessCount"
    Write-Log "Failed: $($FailedOperations.Count)"
    
    if ($FailedOperations.Count -gt 0) {
        Write-Log "Failed operations:" "WARNING"
        foreach ($failed in $FailedOperations) {
            Write-Log "  - $failed" "WARNING"
        }
    }
    
    Write-Log "=========================================="
    Write-Log "TEST TABLES CREATED:"
    Write-Log "- Test_Latest_Tickets (80 tickets max: 20 per type)"
    Write-Log "- Test_Config_Items (configs linked to test tickets)"
    Write-Log "- Test_Notes (all notes for test tickets)"
    Write-Log "- Test_Time_Entries (all time entries for test tickets)"
    Write-Log "- Test_Config_Links (config-to-ticket relationships)"
    Write-Log "=========================================="
    Write-Log "Run this query to see summary:"
    Write-Log "SELECT * FROM [$Database].[dbo].[Test_Latest_Tickets] ORDER BY Ticket_Type, Date_Entered DESC"
    Write-Log "Log file saved to: $LogPath"
}
catch {
    Write-Log "Critical error in main execution: $($_.Exception.Message)" "ERROR"
    exit 1
}
#endregion
