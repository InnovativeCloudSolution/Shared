CW Streamline SOC - Beginner

================================================================================

OVERVIEW

This implementation uses ONLY out-of-the-box ConnectWise RPA workflow actions 
to automate vulnerability ticket management. No coding required.

Implementation Time: 1-2 hours
Difficulty: Beginner
Coding Required: None
Scalability: 1-10 devices

================================================================================

WHAT THIS WORKFLOW DOES

1. Detects NOC vulnerability tickets when they arrive on a designated board
2. Creates parent/child relationships by linking individual device tickets to a manually created master ticket
3. Moves tickets between boards based on status (NOC to Patch Management to Closed)
4. Updates ticket statuses when actions are taken
5. Sends notifications to assigned technicians

================================================================================

PREREQUISITES

- ConnectWise Manage with RPA/Workflow enabled
- Access to create and edit workflows
- Service boards created:
  - "NOC - Vulnerability Tickets" (intake board)
  - "Patch Management - Master Tickets" (master tickets)
  - "Patch Management - In Progress" (active work)

================================================================================

WORKFLOW ARCHITECTURE

Workflow 1: NOC Ticket Processor
--------------------------------------------------------------------------------
Trigger: New ticket created on "NOC - Vulnerability Tickets" board
Conditions: Ticket summary contains "Vulnerability Management"

Actions (OOTB only):
1. Add internal note: "Vulnerability ticket received - awaiting parent assignment"
2. Update ticket status to "New - Awaiting Assignment"
3. Send email notification to Patch Management team

Workflow 2: Link to Parent Ticket
--------------------------------------------------------------------------------
Trigger: Manual execution by technician
Conditions: Ticket is on NOC board

Actions (OOTB only):
1. Prompt user to enter parent ticket ID
2. Link child ticket to parent ticket ID
3. Add internal note: "Linked to master ticket ParentID"
4. Move ticket to "Patch Management - In Progress" board
5. Update status to "Assigned to Parent"

Workflow 3: Status Update - Patch Installed
--------------------------------------------------------------------------------
Trigger: Ticket status changed to "Patch Installed"
Conditions: Ticket is on "Patch Management - In Progress" board

Actions (OOTB only):
1. Add internal note: "Patch installed - awaiting reboot"
2. Update status to "Pending Reboot"
3. Send notification to assignee
4. Add time entry for patch installation (optional)

Workflow 4: Status Update - Patch Verified
--------------------------------------------------------------------------------
Trigger: Ticket status changed to "Verified"
Conditions: Ticket is on "Patch Management - In Progress" board

Actions (OOTB only):
1. Add internal note: "Patch verified - closing ticket"
2. Close ticket with resolution: "Vulnerability remediated"
3. Send notification to parent ticket (add note to parent)
4. Archive ticket

Workflow 5: Parent Ticket Progress Tracker
--------------------------------------------------------------------------------
Trigger: Child ticket closed
Conditions: Parent ticket exists

Actions (OOTB only):
1. Count open child tickets for parent
2. Add note to parent ticket: "Child ticket ChildID closed - X remaining"
3. If all child tickets closed: Update parent status to "Ready for Closure"

================================================================================

SETUP INSTRUCTIONS

Step 1: Create Service Boards
--------------------------------------------------------------------------------
1. Navigate to System > Setup Tables > Service Boards
2. Create three boards:

   NOC - Vulnerability Tickets
     Statuses: New, Awaiting Assignment, Assigned to Parent

   Patch Management - Master Tickets
     Statuses: New, In Progress, Ready for Closure, Closed

   Patch Management - In Progress
     Statuses: Assigned, Patch Installed, Pending Reboot, Verified

Step 2: Create Workflows
--------------------------------------------------------------------------------
1. Navigate to System > Workflows
2. Click "New Workflow"
3. For each workflow above, configure:
   - Name: (e.g., "NOC Ticket Processor")
   - Trigger: Select appropriate trigger type
   - Conditions: Add conditions per workflow specification
   - Actions: Add OOTB actions in sequence

Step 3: Configure OOTB Actions
--------------------------------------------------------------------------------
Each workflow uses only these built-in actions:

Available OOTB Actions:
- Add Note: Add internal or external notes
- Update Ticket: Change status, board, priority, etc.
- Link Tickets: Create parent/child relationships
- Send Email: Notify users or teams
- Close Ticket: Set ticket to closed status
- Add Time Entry: Log time (optional)
- Update Custom Field: Modify ticket custom fields

Action Configuration Example (Workflow 1):

Action 1: Add Note
  Note Type: Internal
  Note Text: "Vulnerability ticket received - awaiting parent assignment"

Action 2: Update Ticket
  Field: Status
  Value: "New - Awaiting Assignment"

Action 3: Send Email
  Recipient: "patchmanagement@company.com"
  Subject: "New Vulnerability Ticket: {TicketSummary}"
  Body: "A new vulnerability ticket has been created. Ticket TicketID requires review and assignment to a master ticket."

Step 4: Create Master Ticket Template
--------------------------------------------------------------------------------
1. Navigate to Service > Tickets
2. Create a new ticket template: "Vulnerability Master Ticket"
3. Configure:
   Board: Patch Management - Master Tickets
   Type: Problem
   Summary: "Patch Deployment - [KB Number] - [Date]"
   Description Template: See MASTER_TICKET_TEMPLATE.md

Step 5: Test Workflows
--------------------------------------------------------------------------------
1. Create a test NOC vulnerability ticket
2. Verify Workflow 1 triggers and adds note
3. Manually create a master ticket using template
4. Use Workflow 2 to link test ticket to master
5. Update test ticket status to "Patch Installed"
6. Verify Workflow 3 triggers
7. Update status to "Verified"
8. Verify Workflow 4 closes ticket and updates parent

================================================================================

USAGE GUIDE

Daily Operations
--------------------------------------------------------------------------------
When NOC tickets arrive:
1. Review new tickets on "NOC - Vulnerability Tickets" board
2. Group tickets by KB patch number
3. Create one master ticket per KB patch (using template)
4. For each child ticket, run Workflow 2 to link to master
5. Child tickets move to "Patch Management - In Progress" board

As patches are deployed:
1. Update child ticket status to "Patch Installed"
2. Workflow 3 auto-updates status to "Pending Reboot"
3. After reboot, update status to "Verified"
4. Workflow 4 auto-closes child ticket
5. Workflow 5 updates parent ticket progress

When all devices patched:
1. Review parent ticket - Workflow 5 sets status to "Ready for Closure"
2. Manually close parent ticket with summary note
3. Archive all related child tickets

================================================================================

LIMITATIONS

- Manual parent ticket creation: Technician must create master ticket
- Manual linking: Technician must run workflow to link child to parent
- No CVE parsing: CVE data stays in notes, not extracted
- No CI field updates: Device CIs not automatically updated
- Manual progress tracking: Parent ticket progress updated by workflow, but counts must be verified manually
- No external API calls: Cannot query NVD or other CVE databases

================================================================================

ADVANTAGES

- No coding required: 100% OOTB workflows
- Quick implementation: Can be set up in 1-2 hours
- Easy maintenance: Workflows are visual and easy to modify
- Low risk: No custom code to debug
- Good for small scale: Works well for 1-10 devices

================================================================================

NEXT STEPS

When you outgrow this implementation:
- Move to Intermediate to add CVE parsing and CI field updates
- Move to Advanced for full automation at scale

================================================================================

WORKFLOW EXPORT FILES

The following files are included in this folder:
- workflow_1_noc_ticket_processor.xml - Import into ConnectWise
- workflow_2_link_to_parent.xml - Import into ConnectWise
- workflow_3_patch_installed.xml - Import into ConnectWise
- workflow_4_patch_verified.xml - Import into ConnectWise
- workflow_5_parent_progress_tracker.xml - Import into ConnectWise
- ticket_template_master.json - Master ticket template

================================================================================

SUPPORT

For issues with OOTB workflows:
- Review ConnectWise Workflow documentation
- Check trigger conditions match ticket criteria
- Verify service board names match workflow configuration
- Test workflows in sandbox environment first

================================================================================

Version History

v1.0.0 (2025-12-07): Initial OOTB workflow implementation
