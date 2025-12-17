# ConnectWise Workflow Configuration Instructions

## Workflows to Create

You'll create 5 workflows using ONLY out-of-the-box ConnectWise RPA actions. No custom code required.

---

## Workflow 1: NOC Ticket Processor

**Purpose**: Automatically process new vulnerability tickets from NOC

**Configuration**:
- **Name**: NOC Ticket Processor
- **Type**: Ticket Workflow
- **Trigger**: Ticket Created
- **Conditions**:
  - Board Name equals "NOC - Vulnerability Tickets"
  - Summary contains "Vulnerability Management"

**Actions** (in order):
1. **Add Note** (OOTB Action)
   - Note Type: Internal
   - Note Text: "Vulnerability ticket received from NOC - awaiting parent ticket assignment"

2. **Update Ticket** (OOTB Action)
   - Field: Status
   - Value: "New - Awaiting Assignment"

3. **Send Email** (OOTB Action)
   - To: patchmanagement@yourcompany.com
   - Subject: "New Vulnerability Ticket: {TicketSummary}"
   - Body: "New vulnerability ticket #{TicketID} created. Requires review and assignment to master ticket."

---

## Workflow 2: Link Child to Parent Ticket

**Purpose**: Manual workflow to link child vulnerability ticket to master ticket

**Configuration**:
- **Name**: Link Child to Parent Ticket
- **Type**: Manual Ticket Workflow
- **Trigger**: Manual Execution
- **Conditions**:
  - Board Name equals "NOC - Vulnerability Tickets"

**Actions** (in order):
1. **Prompt for Input** (OOTB Action)
   - Prompt Text: "Enter Parent Ticket ID:"
   - Variable Name: ParentTicketID
   - Input Type: Number

2. **Link Tickets** (OOTB Action)
   - Parent Ticket ID: {ParentTicketID}
   - Child Ticket ID: {CurrentTicketID}

3. **Add Note** (OOTB Action)
   - Note Type: Internal
   - Note Text: "Linked to master ticket #{ParentTicketID}"

4. **Move Ticket** (OOTB Action)
   - Destination Board: "Patch Management - In Progress"
   - Status: "Assigned to Parent"

5. **Add Note to Parent** (OOTB Action)
   - Ticket ID: {ParentTicketID}
   - Note Type: Internal
   - Note Text: "Child ticket #{CurrentTicketID} linked - Device: {ConfigurationName}"

---

## Workflow 3: Patch Installed Update

**Purpose**: Automate status change when patch is installed

**Configuration**:
- **Name**: Patch Installed Update
- **Type**: Ticket Workflow
- **Trigger**: Ticket Status Changed
- **Conditions**:
  - Board Name equals "Patch Management - In Progress"
  - New Status equals "Patch Installed"

**Actions** (in order):
1. **Add Note** (OOTB Action)
   - Note Type: Internal
   - Note Text: "Patch installed on {ConfigurationName}. Awaiting reboot confirmation."

2. **Update Ticket** (OOTB Action)
   - Field: Status
   - Value: "Pending Reboot"

3. **Send Email** (OOTB Action)
   - To: {AssignedTechnician}
   - Subject: "Patch Installed - Reboot Required: Ticket #{TicketID}"
   - Body: "Patch has been installed on {ConfigurationName}. Please confirm device reboot."

---

## Workflow 4: Patch Verified - Close Ticket

**Purpose**: Automatically close ticket when patch is verified

**Configuration**:
- **Name**: Patch Verified - Close Ticket
- **Type**: Ticket Workflow
- **Trigger**: Ticket Status Changed
- **Conditions**:
  - Board Name equals "Patch Management - In Progress"
  - New Status equals "Verified"

**Actions** (in order):
1. **Add Note** (OOTB Action)
   - Note Type: Internal
   - Note Text: "Patch verified and CVE remediation confirmed. Closing ticket."

2. **Close Ticket** (OOTB Action)
   - Resolution: "Vulnerability remediated - patch installed and verified"
   - Closed Status: "Closed - Resolved"

3. **Add Note to Parent** (OOTB Action - if parent exists)
   - Ticket ID: {ParentTicketID}
   - Note Type: Internal
   - Note Text: "Child ticket #{CurrentTicketID} verified and closed - Device: {ConfigurationName}"

4. **Send Email** (OOTB Action)
   - To: {AssignedTechnician}
   - Subject: "Ticket Closed - Patch Verified: #{TicketID}"
   - Body: "Vulnerability ticket #{TicketID} has been successfully remediated and closed."

---

## Workflow 5: Parent Ticket Progress Tracker

**Purpose**: Update master ticket when child tickets are closed

**Configuration**:
- **Name**: Parent Ticket Progress Tracker
- **Type**: Ticket Workflow
- **Trigger**: Ticket Closed
- **Conditions**:
  - Parent Ticket ID is not empty

**Actions** (in order):
1. **Count Child Tickets** (OOTB Action - if available)
   - Parent Ticket ID: {ParentTicketID}
   - Status Filter: "Open"
   - Store Count In: OpenChildCount

2. **Add Note to Parent** (OOTB Action)
   - Ticket ID: {ParentTicketID}
   - Note Type: Internal
   - Note Text: "Child ticket #{CurrentTicketID} closed. Remaining open child tickets: {OpenChildCount}"

3. **Conditional Action** (OOTB Action)
   - Condition: {OpenChildCount} equals 0
   - Then: Update Parent Ticket Status to "Ready for Closure"

4. **Send Email** (OOTB Action - if all children closed)
   - To: {ParentTicketAssignedTechnician}
   - Subject: "Master Ticket Ready for Closure: #{ParentTicketID}"
   - Body: "All child tickets have been remediated. Master ticket #{ParentTicketID} is ready for final review and closure."

---

## Implementation Steps

1. **Create Each Workflow in ConnectWise**
   - Navigate to System > Workflows
   - Click "New Workflow"
   - Configure trigger and conditions
   - Add actions in exact order listed

2. **Test Each Workflow**
   - Create test tickets matching conditions
   - Verify each action executes correctly
   - Check notes are added properly
   - Confirm emails send (if configured)

3. **Deploy to Production**
   - Enable workflows one at a time
   - Monitor first few executions
   - Verify no conflicts with existing workflows

---

## Notes

- All workflows use ONLY OOTB actions - no custom scripts
- Modify email addresses and text as needed for your environment
- Adjust board names if yours differ
- Some OOTB actions may have different names in your ConnectWise version
- Test thoroughly in sandbox before production deployment

---

## Workflow Action Reference

Common OOTB actions available:
- Add Note
- Update Ticket
- Move Ticket
- Close Ticket
- Send Email
- Link Tickets
- Update Custom Field
- Add Time Entry
- Send Notification
- Conditional Logic
- Count Related Records
- Prompt for Input

