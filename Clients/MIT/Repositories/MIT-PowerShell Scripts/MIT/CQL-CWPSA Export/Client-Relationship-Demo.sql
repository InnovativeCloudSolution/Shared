-- =============================================================================
-- CLIENT DEMONSTRATION QUERIES
-- Showing Ticket Relationships: Config Items, Notes, and Time Entries
-- =============================================================================

-- Query 1: TICKET TO CONFIG ITEM RELATIONSHIPS
-- Shows which Configuration Items are attached to which tickets
SELECT 
    t.Ticket_Type,
    t.SR_Service_RecID AS Ticket_ID,
    t.Summary AS Ticket_Summary,
    t.Date_Entered AS Ticket_Date,
    cl.Config_RecID AS Config_ID,
    cl.Config_Name,
    cl.Config_Type,
    t.Company_Name
FROM Test_Latest_Tickets t
INNER JOIN Test_Config_Links cl ON t.SR_Service_RecID = cl.SR_Service_RecID
ORDER BY t.Ticket_Type, t.Date_Entered DESC, cl.Config_Name;

-- Query 2: TICKET TO NOTES RELATIONSHIPS  
-- Shows all notes attached to tickets
SELECT 
    t.Ticket_Type,
    t.SR_Service_RecID AS Ticket_ID,
    t.Summary AS Ticket_Summary,
    t.Date_Entered AS Ticket_Date,
    n.SR_Detail_RecID AS Note_ID,
    LEFT(n.Note_Content, 100) + '...' AS Note_Preview,
    n.Date_Created AS Note_Date,
    n.Created_By AS Note_Author,
    CASE WHEN n.Internal_Member_Flag = 1 THEN 'Internal' ELSE 'External' END AS Note_Type
FROM Test_Latest_Tickets t
INNER JOIN Test_Notes n ON t.SR_Service_RecID = n.SR_Service_RecID
ORDER BY t.Ticket_Type, t.Date_Entered DESC, n.Date_Created DESC;

-- Query 3: TICKET TO TIME ENTRIES RELATIONSHIPS
-- Shows all time entries logged against tickets
SELECT 
    t.Ticket_Type,
    t.SR_Service_RecID AS Ticket_ID,
    t.Summary AS Ticket_Summary,
    t.Date_Entered AS Ticket_Date,
    te.Time_RecID AS Time_Entry_ID,
    te.Date_Start AS Work_Date,
    te.Hours_Actual AS Hours_Worked,
    te.Hours_Bill AS Hours_Billable,
    te.Member_Name AS Technician,
    te.Activity_Class_Description AS Work_Type,
    LEFT(te.Notes, 50) + '...' AS Work_Description
FROM Test_Latest_Tickets t
INNER JOIN Test_Time_Entries te ON t.SR_Service_RecID = te.SR_Service_RecID
ORDER BY t.Ticket_Type, t.Date_Entered DESC, te.Date_Start DESC;

-- Query 4: COMPREHENSIVE TICKET SUMMARY
-- Shows ticket with counts of related items
SELECT 
    t.Ticket_Type,
    t.SR_Service_RecID AS Ticket_ID,
    t.Summary AS Ticket_Summary,
    t.Date_Entered AS Ticket_Date,
    t.SR_Type_Description AS Ticket_Category,
    t.SR_Status_Description AS Ticket_Status,
    t.Contact_Name AS Customer_Contact,
    COUNT(DISTINCT cl.Config_RecID) AS Config_Items_Attached,
    COUNT(DISTINCT n.SR_Detail_RecID) AS Notes_Count,
    COUNT(DISTINCT te.Time_RecID) AS Time_Entries_Count,
    ISNULL(SUM(te.Hours_Actual), 0) AS Total_Hours_Worked,
    ISNULL(SUM(te.Hours_Bill), 0) AS Total_Hours_Billable
FROM Test_Latest_Tickets t
LEFT JOIN Test_Config_Links cl ON t.SR_Service_RecID = cl.SR_Service_RecID
LEFT JOIN Test_Notes n ON t.SR_Service_RecID = n.SR_Service_RecID
LEFT JOIN Test_Time_Entries te ON t.SR_Service_RecID = te.SR_Service_RecID
GROUP BY 
    t.Ticket_Type, t.SR_Service_RecID, t.Summary, t.Date_Entered,
    t.SR_Type_Description, t.SR_Status_Description, t.Contact_Name
ORDER BY t.Ticket_Type, t.Date_Entered DESC;

-- Query 5: CONFIGURATION ITEM USAGE REPORT
-- Shows which configs are referenced across multiple tickets
SELECT 
    ci.Config_RecID AS Config_ID,
    ci.Config_Name,
    ci.Config_Type,
    ci.Vendor_Name,
    ci.Serial_Number,
    COUNT(DISTINCT cl.SR_Service_RecID) AS Tickets_Referencing,
    STRING_AGG(CAST(cl.SR_Service_RecID AS VARCHAR), ', ') AS Ticket_IDs,
    ci.Company_Name
FROM Test_Config_Items ci
INNER JOIN Test_Config_Links cl ON ci.Config_RecID = cl.Config_RecID
GROUP BY 
    ci.Config_RecID, ci.Config_Name, ci.Config_Type, 
    ci.Vendor_Name, ci.Serial_Number, ci.Company_Name
ORDER BY COUNT(DISTINCT cl.SR_Service_RecID) DESC, ci.Config_Name;

-- Query 6: TICKET ACTIVITY TIMELINE
-- Shows chronological activity for a specific ticket (example with ticket ID)
-- Replace 'TICKET_ID_HERE' with actual ticket ID
/*
SELECT 
    Activity_Type,
    Activity_Date,
    Activity_Description,
    Created_By
FROM (
    -- Ticket Creation
    SELECT 
        'TICKET_CREATED' AS Activity_Type,
        Date_Entered AS Activity_Date,
        'Ticket: ' + Summary AS Activity_Description,
        Entered_By AS Created_By
    FROM Test_Latest_Tickets 
    WHERE SR_Service_RecID = TICKET_ID_HERE
    
    UNION ALL
    
    -- Notes Added
    SELECT 
        'NOTE_ADDED' AS Activity_Type,
        Date_Created AS Activity_Date,
        'Note: ' + LEFT(Note_Content, 50) + '...' AS Activity_Description,
        Created_By
    FROM Test_Notes 
    WHERE SR_Service_RecID = TICKET_ID_HERE
    
    UNION ALL
    
    -- Time Entries
    SELECT 
        'TIME_LOGGED' AS Activity_Type,
        Date_Start AS Activity_Date,
        'Work: ' + CAST(Hours_Actual AS VARCHAR) + 'h - ' + LEFT(Notes, 50) + '...' AS Activity_Description,
        Member_Name AS Created_By
    FROM Test_Time_Entries 
    WHERE SR_Service_RecID = TICKET_ID_HERE
    
    UNION ALL
    
    -- Config Attachments
    SELECT 
        'CONFIG_ATTACHED' AS Activity_Type,
        t.Date_Entered AS Activity_Date,
        'Config attached: ' + cl.Config_Name AS Activity_Description,
        t.Entered_By AS Created_By
    FROM Test_Config_Links cl
    INNER JOIN Test_Latest_Tickets t ON cl.SR_Service_RecID = t.SR_Service_RecID
    WHERE cl.SR_Service_RecID = TICKET_ID_HERE
) AS Activities
ORDER BY Activity_Date ASC;
*/

-- Query 7: SUMMARY STATISTICS FOR CLIENT
-- High-level overview of the data export
SELECT 
    Report_Type,
    Item_Count,
    Details
FROM (
    SELECT 'TICKETS' AS Report_Type, COUNT(*) AS Item_Count, 
           STRING_AGG(Ticket_Type + ': ' + CAST(Type_Count AS VARCHAR), ', ') AS Details
    FROM (
        SELECT Ticket_Type, COUNT(*) AS Type_Count 
        FROM Test_Latest_Tickets 
        GROUP BY Ticket_Type
    ) AS TicketSummary
    
    UNION ALL
    
    SELECT 'CONFIG_ITEMS' AS Report_Type, COUNT(*) AS Item_Count, 
           'Configuration items linked to tickets' AS Details
    FROM Test_Config_Items
    
    UNION ALL
    
    SELECT 'NOTES' AS Report_Type, COUNT(*) AS Item_Count,
           'Notes and communications on tickets' AS Details
    FROM Test_Notes
    
    UNION ALL
    
    SELECT 'TIME_ENTRIES' AS Report_Type, COUNT(*) AS Item_Count,
           'Work log entries with ' + CAST(ROUND(SUM(Hours_Actual), 1) AS VARCHAR) + ' total hours' AS Details
    FROM Test_Time_Entries
    
    UNION ALL
    
    SELECT 'RELATIONSHIPS' AS Report_Type, COUNT(*) AS Item_Count,
           'Config-to-Ticket relationships' AS Details
    FROM Test_Config_Links
) AS Summary
ORDER BY 
    CASE Report_Type 
        WHEN 'TICKETS' THEN 1
        WHEN 'CONFIG_ITEMS' THEN 2
        WHEN 'NOTES' THEN 3
        WHEN 'TIME_ENTRIES' THEN 4
        WHEN 'RELATIONSHIPS' THEN 5
    END;

-- =============================================================================
-- INSTRUCTIONS FOR CLIENT:
-- 
-- 1. Run Query 1 to see which Configuration Items are attached to tickets
-- 2. Run Query 2 to see all Notes/Communications for tickets  
-- 3. Run Query 3 to see all Time Entries (work log) for tickets
-- 4. Run Query 4 for a comprehensive summary with counts
-- 5. Run Query 5 to see Configuration Item usage across tickets
-- 6. Uncomment and modify Query 6 to see activity timeline for specific ticket
-- 7. Run Query 7 for high-level statistics
--
-- These queries demonstrate the complete relationship structure between:
-- - Tickets (Events, Incidents, Problems, Requests)  
-- - Configuration Items (servers, workstations, etc.)
-- - Notes (communications, updates, resolutions)
-- - Time Entries (work performed, hours logged)
-- =============================================================================
