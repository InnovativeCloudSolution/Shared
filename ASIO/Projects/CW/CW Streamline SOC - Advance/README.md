ConnectWise SOC Vulnerability Management Workflow

================================================================================

OVERVIEW

This automated SOC-like workflow consolidates vulnerability management tickets 
from ConnectWise NOC, tracks patch deployment status using Configuration Item 
(CI) custom fields, and provides centralized oversight through master tickets.

================================================================================

ARCHITECTURE

Core Components
--------------------------------------------------------------------------------
- Tracking Method: ConnectWise CI Custom Fields (no external database)
- Automation Platform: ConnectWise RPA (ASIO) or scheduled Python scripts
- Boards:
  - Intake: "NOC - Vulnerability Tickets" (dump board for individual device tickets)
  - Action: "Patch Management - Master Tickets" (consolidated oversight)
  - Exceptions: "Patch Management - Failures" (requires manual intervention)

Workflow Bots
--------------------------------------------------------------------------------
1. Bot 1: NOC Ticket Ingestion & CI Update
2. Bot 2: Master Ticket Creator
3. Bot 3: Patch Status Monitor
4. Bot 4: Master Ticket Updater
5. Bot 5: Auto-Closure Handler
6. Bot 6: Exception Handler

================================================================================

CI CUSTOM FIELD SCHEMA

Device CI Custom Fields
--------------------------------------------------------------------------------
Add these custom fields to your ConnectWise device configuration type:

Field Name                     | Type      | Size | Description
-------------------------------|-----------|------|-----------------------------------
Pending_CVEs                   | Text Area | 4000 | JSON array of CVE objects with severity and KB mapping
Pending_KB_Patches             | Text      | 500  | Comma-separated list of KB numbers awaiting installation
Active_Vulnerability_Tickets   | Text      | 500  | Comma-separated ticket IDs related to this device
Last_Vulnerability_Scan        | Date      | -    | Timestamp of last NOC vulnerability scan
Patch_Status_KB{number}        | Dropdown  | -    | Dynamic field per KB: Pending, Patched, Rebooted, Verified, Failed
Patch_Installed_Date_KB{number}| Date      | -    | Timestamp when specific KB was installed
Vulnerability_Summary          | Text Area | 2000 | Human-readable summary of pending vulnerabilities

Example CI Field Values
--------------------------------------------------------------------------------
Pending_CVEs:
[
  {"cve": "CVE-2025-59505", "severity": "Critical", "kb": "KB5068861"},
  {"cve": "CVE-2025-59506", "severity": "High", "kb": "KB5068861"},
  {"cve": "CVE-2025-60703", "severity": "High", "kb": "KB5068861"}
]

Pending_KB_Patches:
KB5068861,KB5067890,KB5069001

Active_Vulnerability_Tickets:
26292,26345,26401

Vulnerability_Summary:
Total Pending: 35 CVEs (Critical: 1, High: 28, Medium: 6)
KB5068861: Pending Install
Last Updated: 2025-12-07 10:30

================================================================================

WORKFLOW PROCESS

Phase 1: Ticket Ingestion (Bot 1)
--------------------------------------------------------------------------------
Trigger: New ticket in NOC board with tag "---##@@UpdateFromNOC@@##---"

Process:
1. Parse NOC ticket notes for CVE IDs, KB numbers, severity counts
2. Extract device CI from ticket configuration
3. Update device CI custom fields with vulnerability data
4. Check if master ticket exists for this KB + Company
5. If no master ticket exists, call Bot 2
6. Link child ticket to master ticket
7. Move child ticket to dump board with status "Awaiting Patch"

Phase 2: Master Ticket Creation (Bot 2)
--------------------------------------------------------------------------------
Trigger: Called by Bot 1 when new KB detected for company

Process:
1. Query all device CIs for company where Pending_KB_Patches contains this KB
2. Aggregate data: total devices, severity counts, affected devices list
3. Create master ticket in "Patch Management - Master Tickets" board
4. Summary: "Patch Deployment - KB{number} - {count} Devices - {cve_count} CVEs"
5. Add detailed notes with device list and status table
6. Update master ticket ID in all related device CIs

Phase 3: Status Monitoring (Bot 3)
--------------------------------------------------------------------------------
Trigger: Scheduled (every 4 hours)

Process:
1. Query all device CIs where Pending_KB_Patches is not empty
2. For each device/KB combination:
   - Check RMM for patch installation status
   - Check device reboot status
   - Query vulnerability scan results
3. Update CI custom fields based on findings:
   - Pending to Patched (KB installed)
   - Patched to Rebooted (device rebooted post-patch)
   - Rebooted to Verified (vulnerability scan confirms CVE resolved)
   - Any status to Failed (patch installation failed)
4. Update child ticket notes with status change
5. If status = Verified, call Bot 5 (Auto-Closure)
6. If status = Failed, call Bot 6 (Exception Handler)

Phase 4: Master Ticket Updates (Bot 4)
--------------------------------------------------------------------------------
Trigger: Scheduled (every 1 hour)

Process:
1. Query all master tickets where status != Closed
2. Extract KB number from master ticket summary
3. Query all device CIs with this KB for the company
4. Aggregate status counts from Patch_Status_KB{number} field
5. Calculate progress percentage
6. Update master ticket notes with current status table
7. Update master ticket board status:
   - New to InProgress (first device patched)
   - InProgress to Complete (all devices verified or failed)

Phase 5: Auto-Closure (Bot 5)
--------------------------------------------------------------------------------
Trigger: Called by Bot 3 when device reaches "Verified" status

Process:
1. Get child ticket ID from device CI Active_Vulnerability_Tickets
2. Close child ticket with resolution note
3. Remove ticket ID from device CI Active_Vulnerability_Tickets
4. Remove resolved CVEs from device CI Pending_CVEs
5. Remove KB from device CI Pending_KB_Patches if all CVEs resolved
6. Query remaining incomplete devices for master ticket
7. If all devices complete, close master ticket after 24h grace period

Phase 6: Exception Handling (Bot 6)
--------------------------------------------------------------------------------
Trigger: Called by Bot 3 when device status = "Failed"

Process:
1. Move child ticket to "Patch Management - Failures" board
2. Add failure alert to master ticket notes
3. Send notification to assigned technician
4. Update device CI with failure reason
5. Increment failure count in master ticket
6. Create remediation checklist in child ticket

================================================================================

INSTALLATION & CONFIGURATION

Prerequisites
--------------------------------------------------------------------------------
- ConnectWise Manage with API access
- ConnectWise RMM with API access (for patch status queries)
- Python 3.8+ (for standalone execution)
- ConnectWise RPA (ASIO) platform (optional, for integrated execution)

Setup Steps
--------------------------------------------------------------------------------
1. Create CI Custom Fields
   - Navigate to System > Setup Tables > Configuration Types
   - Select your device configuration type
   - Add custom fields per schema above
   - Configure dropdown values for Patch_Status_KB{number} field template

2. Create Service Boards
   - Create "NOC - Vulnerability Tickets" board
   - Create "Patch Management - Master Tickets" board
   - Create "Patch Management - Failures" board
   - Configure board statuses and workflows

3. Configure API Access
   - Generate ConnectWise Manage API keys
   - Generate ConnectWise RMM API keys
   - Store credentials in secure location (ConnectWise RPA vault or environment variables)

4. Deploy Bot Scripts
   - Copy bot scripts to your automation platform
   - Update config.ini with your environment details
   - Test each bot individually before full deployment

5. Schedule Bot Execution
   - Bot 1: Event-driven (new ticket webhook) or polling every 15 minutes
   - Bot 2: Called by Bot 1 (no schedule)
   - Bot 3: Schedule every 4 hours
   - Bot 4: Schedule every 1 hour
   - Bot 5: Called by Bot 3 (no schedule)
   - Bot 6: Called by Bot 3 (no schedule)

Configuration File
--------------------------------------------------------------------------------
Edit config.ini with your environment settings:

[ConnectWise]
cw_manage_url = https://your-instance.connectwisemanagedservices.com
cw_manage_company_id = your_company_id
cw_manage_public_key = your_public_key
cw_manage_private_key = your_private_key
cw_manage_client_id = your_client_id

cw_rmm_url = https://your-rmm-instance.connectwisecontrol.com
cw_rmm_api_key = your_rmm_api_key

[Boards]
noc_board_name = NOC - Vulnerability Tickets
master_board_name = Patch Management - Master Tickets
exception_board_name = Patch Management - Failures

[Settings]
noc_tag = ---##@@UpdateFromNOC@@##---
patch_check_interval_hours = 4
master_update_interval_hours = 1

================================================================================

BOT EXECUTION

Manual Execution
--------------------------------------------------------------------------------
cd ICS-Shared/ASIO/Projects/CW/CW Streamline SOC - Advance

python bot_1_noc_ticket_ingestion.py
python bot_3_patch_status_monitor.py
python bot_4_master_ticket_updater.py

Scheduled Execution (Windows Task Scheduler)
--------------------------------------------------------------------------------
schtasks /create /tn "CW SOC Bot 3 - Patch Monitor" /tr "python C:\Path\To\bot_3_patch_status_monitor.py" /sc hourly /mo 4
schtasks /create /tn "CW SOC Bot 4 - Master Updater" /tr "python C:\Path\To\bot_4_master_ticket_updater.py" /sc hourly

Scheduled Execution (Linux Cron)
--------------------------------------------------------------------------------
0 */4 * * * cd /path/to/ICS-Shared/CW && python3 bot_3_patch_status_monitor.py
0 * * * * cd /path/to/ICS-Shared/CW && python3 bot_4_master_ticket_updater.py

================================================================================

CI QUERY EXAMPLES

Find all devices needing specific KB:
GET /company/configurations?conditions=customFields/Pending_KB_Patches contains 'KB5068861'

Find devices by patch status:
GET /company/configurations?conditions=customFields/Patch_Status_KB5068861='Pending'

Find devices with active vulnerability tickets:
GET /company/configurations?conditions=customFields/Active_Vulnerability_Tickets!=null

Find all devices with any pending patches:
GET /company/configurations?conditions=company/id={company_id} AND customFields/Pending_KB_Patches!=null

Find overdue patches (last scan > 30 days):
GET /company/configurations?conditions=customFields/Last_Vulnerability_Scan<[2025-11-07] AND customFields/Pending_KB_Patches!=null

Find critical vulnerabilities:
GET /company/configurations?conditions=customFields/Pending_CVEs contains '"severity": "Critical"'

================================================================================

REPORTING

Daily Status Report
--------------------------------------------------------------------------------
Query all master tickets and aggregate progress:
- Total active patch deployments
- Devices by status (Pending, Patched, Rebooted, Verified, Failed)
- Behind schedule patches
- Critical vulnerabilities pending

Weekly Summary Report
--------------------------------------------------------------------------------
- Total CVEs remediated
- Average time to remediate by severity
- Patch compliance percentage by company
- Failed patch rate and common failures

Compliance Dashboard
--------------------------------------------------------------------------------
- Devices with pending critical CVEs
- Overdue patches (> 30 days)
- Patch deployment success rate
- Mean time to remediate (MTTR) by severity

================================================================================

TROUBLESHOOTING

Bot 1 Not Parsing Tickets
--------------------------------------------------------------------------------
Issue: Tickets not being parsed or CI not updating

Resolution:
1. Verify NOC ticket contains "---##@@UpdateFromNOC@@##---" tag
2. Check ticket has device CI linked in configuration field
3. Verify regex patterns match NOC ticket format
4. Check API permissions for CI updates

Bot 3 Not Detecting Patches
--------------------------------------------------------------------------------
Issue: Patch installed but CI status not updating

Resolution:
1. Verify RMM API credentials and permissions
2. Check device is online and communicating with RMM
3. Verify KB number format matches expected pattern
4. Check CI custom field exists for this KB

Master Ticket Not Creating
--------------------------------------------------------------------------------
Issue: Child tickets created but no master ticket

Resolution:
1. Verify "Patch Management - Master Tickets" board exists
2. Check API permissions for ticket creation
3. Verify company ID is valid
4. Check Bot 2 logs for errors

Child Tickets Not Closing
--------------------------------------------------------------------------------
Issue: Devices verified but tickets remain open

Resolution:
1. Verify Bot 5 is scheduled or being called by Bot 3
2. Check ticket IDs in CI Active_Vulnerability_Tickets are valid
3. Verify API permissions for ticket closure
4. Check ticket status allows closure (not already closed)

================================================================================

MAINTENANCE

Weekly Tasks
--------------------------------------------------------------------------------
- Review exception board for failed patches
- Verify bot execution logs for errors
- Check CI custom field data consistency

Monthly Tasks
--------------------------------------------------------------------------------
- Archive closed master tickets older than 90 days
- Review and optimize bot schedules based on ticket volume
- Update CI custom field templates for new patch types

Quarterly Tasks
--------------------------------------------------------------------------------
- Review patch deployment metrics
- Update NOC ticket parsing patterns if format changes
- Audit CI custom field usage and cleanup stale data

================================================================================

SECURITY CONSIDERATIONS

- Store all API credentials securely (vault or environment variables)
- Use least-privilege API keys (read/write only required endpoints)
- Audit bot actions via ConnectWise API logs
- Encrypt CI custom field data if storing sensitive information
- Implement rate limiting to avoid API throttling

================================================================================

SUPPORT & CONTACT

For issues or questions regarding this workflow:
- Review bot logs in logs/ directory
- Check ConnectWise API audit logs for failed requests
- Contact DevOps team for automation platform issues

================================================================================

Version History

v1.0.0 (2025-12-07): Initial release with 6 core bots
