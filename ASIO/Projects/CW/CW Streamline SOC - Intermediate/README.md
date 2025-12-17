CW Streamline SOC - Intermediate

================================================================================

OVERVIEW

This implementation extends the Beginner OOTB workflows by adding custom 
PowerShell and Python actions to extract CVE data from ticket notes and enrich 
tickets with external CVE information.

Implementation Time: 4-8 hours
Difficulty: Intermediate
Coding Required: Basic PowerShell/Python
Scalability: 10-100 devices

================================================================================

WHAT THIS ADDS BEYOND BEGINNER

1. CVE extraction from ticket notes using regex parsing
2. External CVE lookup from NVD (https://nvd.nist.gov/vuln/detail/CVE-XXXX-XXXXX)
3. CI custom field updates to track vulnerabilities per device
4. Automated KB patch extraction from ticket notes
5. Severity scoring from external CVE database
6. Enhanced ticket enrichment with CVSS scores and descriptions

================================================================================

PREREQUISITES

- All prerequisites from Beginner implementation
- PowerShell 5.1+ or Python 3.8+ available on workflow execution server
- CI custom fields created (see Setup section)
- Network access to https://nvd.nist.gov API

================================================================================

ARCHITECTURE

Custom Actions Overview
--------------------------------------------------------------------------------
Action Name              | Language   | Purpose                                  | Triggered By
-------------------------|------------|------------------------------------------|------------------
Parse CVE from Notes     | PowerShell | Extract CVE IDs from ticket notes        | New NOC ticket
Lookup CVE Details       | Python     | Query NVD API for CVE details            | After CVE parsing
Update CI Fields         | PowerShell | Write vulnerability data to device CI    | After CVE lookup
Extract KB Number        | PowerShell | Parse KB patch number from notes         | New NOC ticket
Build CVE Summary        | PowerShell | Create formatted summary for ticket      | After all CVEs processed

Enhanced Workflow Flow
--------------------------------------------------------------------------------
NOC Ticket Created
  |
  v
[OOTB] Add initial note
  |
  v
[CUSTOM] Parse CVE from Notes (Extract CVE IDs)
  |
  v
[CUSTOM] Extract KB Number (Get patch KB)
  |
  v
[CUSTOM] Lookup CVE Details (Query NVD API for each CVE)
  |
  v
[CUSTOM] Update CI Fields (Write to device configuration)
  |
  v
[CUSTOM] Build CVE Summary (Format ticket description)
  |
  v
[OOTB] Update ticket with enriched data
  |
  v
[OOTB] Link to parent ticket (manual or automated)

================================================================================

CUSTOM ACTIONS

Custom Action 1: Parse CVE from Notes
--------------------------------------------------------------------------------
File: custom_action_parse_cve.ps1

Purpose: Extract all CVE IDs from ticket notes using regex

Input: 
- Ticket ID

Output:
- Array of CVE IDs
- Count of CVEs found

Script: See custom_action_parse_cve.ps1

Custom Action 2: Lookup CVE Details
--------------------------------------------------------------------------------
File: custom_action_lookup_cve.py

Purpose: Query NVD API for CVE severity, CVSS score, and description

Input:
- CVE ID (e.g., CVE-2025-29969)

Output:
- Severity (Critical/High/Medium/Low)
- CVSS Score
- Description
- Published Date

Script: See custom_action_lookup_cve.py

Custom Action 3: Update CI Fields
--------------------------------------------------------------------------------
File: custom_action_update_ci.ps1

Purpose: Write vulnerability data to device CI custom fields

Input:
- CI ID
- CVE data (JSON array)
- KB number

Output:
- Success/failure status

Script: See custom_action_update_ci.ps1

Custom Action 4: Extract KB Number
--------------------------------------------------------------------------------
File: custom_action_extract_kb.ps1

Purpose: Parse KB patch number from ticket notes

Input:
- Ticket notes text

Output:
- KB number (e.g., "KB5068861")

Script: See custom_action_extract_kb.ps1

Custom Action 5: Build CVE Summary
--------------------------------------------------------------------------------
File: custom_action_build_summary.ps1

Purpose: Create formatted summary for ticket description

Input:
- Array of CVE objects with details

Output:
- Formatted markdown summary

Script: See custom_action_build_summary.ps1

================================================================================

SETUP INSTRUCTIONS

Step 1: Create CI Custom Fields
--------------------------------------------------------------------------------
Navigate to System > Setup Tables > Configuration Types > Select device type > Custom Fields

Add the following custom fields:

Field Name                      | Type      | Size | Description
--------------------------------|-----------|------|-------------
Pending_CVEs                    | Text Area | 4000 | JSON array of CVE objects
Pending_KB_Patches              | Text      | 500  | Comma-separated KB numbers
Active_Vulnerability_Tickets    | Text      | 500  | Comma-separated ticket IDs
Last_Vulnerability_Scan         | Date      | -    | Last scan timestamp
Patch_Status_{KB}               | Dropdown  | -    | Pending, Patched, Rebooted, Verified, Failed
Vulnerability_Summary           | Text Area | 2000 | Human-readable summary

Step 2: Install Custom Actions
--------------------------------------------------------------------------------
1. Copy all .ps1 and .py files to workflow execution server
2. Ensure PowerShell execution policy allows scripts: Set-ExecutionPolicy RemoteSigned
3. Install Python dependencies: pip install requests
4. Test each script independently before integrating

Step 3: Configure Workflow to Use Custom Actions
--------------------------------------------------------------------------------
In ConnectWise Workflows, add "Execute Script" actions:

Example: After NOC ticket created
1. [OOTB] Add initial note
2. [CUSTOM] Execute custom_action_parse_cve.ps1 -TicketID {TicketID}
3. [CUSTOM] For each CVE, execute python custom_action_lookup_cve.py {CVEID}
4. [CUSTOM] Execute custom_action_update_ci.ps1 -CIID {ConfigurationID} -CVEDataJSON {CVEData} -KBNumber {KB}
5. [CUSTOM] Execute custom_action_build_summary.ps1 -CVEDataJSON {CVEData}
6. [OOTB] Update ticket description with summary
7. [OOTB] Continue with parent linking workflow

Step 4: Configure Environment Variables
--------------------------------------------------------------------------------
On the workflow execution server, set:

$env:CW_URL = "https://your-instance.connectwisemanagedservices.com"
$env:CW_AUTH = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("CompanyID+PublicKey:PrivateKey"))

Step 5: Test Custom Actions
--------------------------------------------------------------------------------
1. Create a test NOC ticket with sample CVE data in notes
2. Manually execute each custom action
3. Verify CI fields are updated correctly
4. Verify CVE data is retrieved from NVD
5. Integrate into workflows once tested

================================================================================

EXAMPLE NOC TICKET FORMAT

Summary: Vulnerability Management | Total CVEs 35 | Installing the patch will resolve the issue

Notes:
---##@@UpdateFromNOC@@##---

Patch to be installed: KB5068861

Total CVEs linked with this ticket: 35
Critical Severity CVEs: 1
High Severity CVEs: 28
Medium Severity CVEs: 6

---##@@UpdateFromNOC@@##---

https://nvd.nist.gov/vuln/detail/CVE-2025-59505
https://nvd.nist.gov/vuln/detail/CVE-2025-59506
https://nvd.nist.gov/vuln/detail/CVE-2025-59507
(... more CVE links ...)

---##@@UpdateFromNOC@@##---

After custom actions run, ticket is enriched with:
- Severity breakdown with CVSS scores
- CVE descriptions
- Device CI updated with vulnerability tracking data

================================================================================

ADVANTAGES OVER BEGINNER

- Automated CVE extraction: No manual parsing required
- External data enrichment: Real CVSS scores and descriptions from NVD
- CI field tracking: Device configurations track vulnerability status
- Better reporting: CI queries enable vulnerability dashboards
- Scalability: Handles 10-100 devices efficiently

================================================================================

LIMITATIONS

- Still requires manual master ticket creation
- Custom actions require maintenance as APIs change
- Dependent on external NVD API availability
- No automated patch status monitoring
- No intelligent auto-closure

================================================================================

NEXT STEPS

When you outgrow this implementation:
- Move to Advanced for full automation with custom bots
- Implement automated master ticket creation
- Add patch status monitoring
- Enable intelligent auto-closure

================================================================================

FILES INCLUDED

- custom_action_parse_cve.ps1
- custom_action_lookup_cve.py
- custom_action_update_ci.ps1
- custom_action_extract_kb.ps1
- custom_action_build_summary.ps1
- workflow_enhanced_noc_processor.xml
- requirements.txt (Python dependencies)

================================================================================

Version History

v1.0.0 (2025-12-07): Initial workflow + custom actions implementation
