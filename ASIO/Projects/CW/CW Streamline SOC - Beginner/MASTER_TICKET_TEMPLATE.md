# Master Ticket Template

## How to Use

When grouping vulnerability tickets by KB patch, create ONE master ticket per KB using this template.

---

## Ticket Fields

**Board**: Patch Management - Master Tickets  
**Type**: Problem  
**Status**: New  
**Priority**: Set based on Critical CVE count (High if any Critical CVEs)  
**Summary**: `Patch Deployment - KB[NUMBER] - [DATE]`

**Example Summary**:
```
Patch Deployment - KB5068861 - December 2025
```

---

## Ticket Description Template

Copy and paste this into the ticket description, then fill in the placeholders:

```
## Patch Deployment Master Ticket

**KB Number**: KB[ENTER_KB_NUMBER]  
**Patch Name**: [ENTER_PATCH_NAME]  
**Release Date**: [ENTER_DATE]  
**Affected Devices**: [ENTER_COUNT]  

---

## Vulnerability Summary

**Total CVEs Addressed**: [ENTER_COUNT]  
- Critical Severity: [ENTER_COUNT]  
- High Severity: [ENTER_COUNT]  
- Medium Severity: [ENTER_COUNT]  
- Low Severity: [ENTER_COUNT]  

---

## Deployment Status

**Total Child Tickets**: [ENTER_COUNT]  
**Status Breakdown**:
- Pending Assignment: [COUNT]
- Patch Installed: [COUNT]
- Pending Reboot: [COUNT]
- Verified: [COUNT]
- Closed: [COUNT]

---

## Affected Devices

| Device Name | CI ID | Ticket ID | Status | Last Updated |
|-------------|-------|-----------|--------|--------------|
| [DEVICE1] | [CI_ID] | #[TICKET] | Pending | [DATE] |
| [DEVICE2] | [CI_ID] | #[TICKET] | Pending | [DATE] |
| [ADD MORE ROWS AS NEEDED] | | | | |

---

## Installation Instructions

1. Deploy KB[NUMBER] via RMM or manual installation
2. Verify installation via Windows Update history
3. Reboot device if required
4. Run vulnerability scan to confirm CVE remediation
5. Update child ticket status as progress is made

---

## Notes

Updates will be automatically added by workflows as child tickets progress.

---

## Closure Criteria

- All child tickets closed successfully
- All devices show patch installed
- Vulnerability scan confirms CVE remediation
- No failed installations
```

---

## Example Completed Master Ticket

```
## Patch Deployment Master Ticket

**KB Number**: KB5068861  
**Patch Name**: 2025-12 Cumulative Update for Windows Server 2022  
**Release Date**: December 10, 2025  
**Affected Devices**: 47  

---

## Vulnerability Summary

**Total CVEs Addressed**: 35  
- Critical Severity: 1  
- High Severity: 28  
- Medium Severity: 6  
- Low Severity: 0  

---

## Deployment Status

**Total Child Tickets**: 47  
**Status Breakdown**:
- Pending Assignment: 0
- Patch Installed: 5
- Pending Reboot: 12
- Verified: 18
- Closed: 12

**Progress**: 64% Complete

---

## Affected Devices

| Device Name | CI ID | Ticket ID | Status | Last Updated |
|-------------|-------|-----------|--------|--------------|
| SERVER-DC01 | 4567 | #26292 | Verified | 2025-12-07 14:30 |
| SERVER-SQL02 | 4568 | #26345 | Pending Reboot | 2025-12-07 15:00 |
| SERVER-WEB01 | 4569 | #26401 | Patch Installed | 2025-12-07 15:15 |
| SERVER-APP01 | 4570 | #26455 | Verified | 2025-12-07 14:45 |

---

## Installation Instructions

1. Deploy KB5068861 via ConnectWise RMM automated patch policy
2. Verify installation via Windows Update history
3. Reboot device (required for this KB)
4. Run Qualys/Tenable scan to confirm CVE remediation
5. Update child ticket status as progress is made

---

## Notes

**2025-12-07 10:30** - Master ticket created, 47 child tickets identified  
**2025-12-07 14:00** - Patch deployment started via RMM  
**2025-12-07 15:30** - 12 devices verified, 18 pending reboot  
```

---

## Tips for Using This Template

1. **Create master ticket BEFORE linking children**
   - Easier to reference the master ticket ID

2. **Update device table as tickets are linked**
   - Use workflow notes to track which devices linked

3. **Update status breakdown daily**
   - Query child tickets to get accurate counts

4. **Use custom fields for tracking** (optional)
   - Add custom fields for Total Devices, Pending Count, etc.
   - Update via workflows or manually

5. **Close master ticket ONLY when all children closed**
   - Workflow 5 will set status to "Ready for Closure"
   - Manually verify all devices patched before closing

---

## Quick Copy Template (Minimal)

For faster creation, use this minimal template:

```
KB: [NUMBER]
Devices: [COUNT]
CVEs: [COUNT] (Critical: [X], High: [Y], Medium: [Z])
Status: [PENDING/IN PROGRESS/READY FOR CLOSURE]
```

Then let workflow notes add detail as tickets progress.

