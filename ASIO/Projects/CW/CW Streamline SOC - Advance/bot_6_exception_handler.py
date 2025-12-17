import os
import re
import json
import base64
import requests
import configparser
from datetime import datetime

config = configparser.ConfigParser()
config.read('config.ini')

CW_URL = os.getenv("CW_MANAGE_URL", config.get('ConnectWise', 'cw_manage_url'))
CW_COMPANY = os.getenv("CW_COMPANY_ID", config.get('ConnectWise', 'cw_manage_company_id'))
CW_PUBLIC_KEY = os.getenv("CW_PUBLIC_KEY", config.get('ConnectWise', 'cw_manage_public_key'))
CW_PRIVATE_KEY = os.getenv("CW_PRIVATE_KEY", config.get('ConnectWise', 'cw_manage_private_key'))
CW_CLIENT_ID = os.getenv("CW_CLIENT_ID", config.get('ConnectWise', 'cw_manage_client_id'))

EXCEPTION_BOARD = config.get('Boards', 'exception_board_name')
MASTER_BOARD = config.get('Boards', 'master_board_name')


def get_auth_header():
    auth_string = f"{CW_COMPANY}+{CW_PUBLIC_KEY}:{CW_PRIVATE_KEY}"
    encoded = base64.b64encode(auth_string.encode()).decode()
    return {
        "Authorization": f"Basic {encoded}",
        "Content-Type": "application/json",
        "clientId": CW_CLIENT_ID
    }


def cw_get(endpoint):
    url = f"{CW_URL}/v4_6_release/apis/3.0/{endpoint}"
    response = requests.get(url, headers=get_auth_header(), timeout=30)
    if response.status_code == 200:
        return response.json()
    return None


def cw_post(endpoint, data):
    url = f"{CW_URL}/v4_6_release/apis/3.0/{endpoint}"
    response = requests.post(url, headers=get_auth_header(), json=data, timeout=30)
    if response.status_code in [200, 201]:
        return response.json()
    return None


def cw_patch(endpoint, data):
    url = f"{CW_URL}/v4_6_release/apis/3.0/{endpoint}"
    response = requests.patch(url, headers=get_auth_header(), json=data, timeout=30)
    if response.status_code == 200:
        return response.json()
    return None


def get_ci_field(ci_data, field_name):
    custom_fields = ci_data.get("customFields", [])
    for field in custom_fields:
        if field.get("caption") == field_name:
            return field.get("value", "")
    return ""


def move_ticket_to_exception_board(ticket_id, kb_number, reason):
    exception_board = cw_get(f"service/boards?conditions=name='{EXCEPTION_BOARD}'")
    if not exception_board or len(exception_board) == 0:
        print(f"  ERROR: Exception board '{EXCEPTION_BOARD}' not found")
        return False
    
    board_id = exception_board[0]["id"]
    
    operations = [
        {"op": "replace", "path": "board/id", "value": board_id},
        {"op": "replace", "path": "status/name", "value": "Failed - Requires Attention"}
    ]
    
    result = cw_patch(f"service/tickets/{ticket_id}", operations)
    
    if result:
        note_data = {
            "text": f"Patch installation failed for {kb_number}. Reason: {reason}\n\nTicket moved to exception board for manual remediation.",
            "detailDescriptionFlag": False,
            "internalAnalysisFlag": False,
            "externalFlag": True
        }
        cw_post(f"service/tickets/{ticket_id}/notes", note_data)
        return True
    
    return False


def add_failure_alert_to_master(master_ticket_id, device_name, ci_id, kb_number, reason):
    alert = f"## FAILURE ALERT\n\n"
    alert += f"**Device**: {device_name} (CI: {ci_id})\n"
    alert += f"**KB**: {kb_number}\n"
    alert += f"**Failure Time**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
    alert += f"**Reason**: {reason}\n"
    alert += f"**Action Required**: Manual remediation needed\n\n"
    alert += f"---\n"
    
    note_data = {
        "text": alert,
        "detailDescriptionFlag": False,
        "internalAnalysisFlag": True
    }
    
    return cw_post(f"service/tickets/{master_ticket_id}/notes", note_data)


def send_failure_notification(ticket_id, device_name, kb_number):
    assignee = cw_get(f"service/tickets/{ticket_id}")
    if not assignee:
        return False
    
    assigned_tech = assignee.get("owner", {}).get("identifier")
    if not assigned_tech:
        return False
    
    note_data = {
        "text": f"ALERT: Patch installation failed on {device_name} for {kb_number}. Manual intervention required. See ticket #{ticket_id}.",
        "detailDescriptionFlag": False,
        "internalAnalysisFlag": False,
        "externalFlag": True
    }
    
    return cw_post(f"service/tickets/{ticket_id}/notes", note_data)


def create_remediation_checklist(ticket_id, kb_number):
    checklist = f"## Remediation Checklist for {kb_number}\n\n"
    checklist += f"- [ ] Verify device is online and accessible\n"
    checklist += f"- [ ] Check Windows Update service is running\n"
    checklist += f"- [ ] Review Windows Update logs for errors\n"
    checklist += f"- [ ] Check disk space on system drive\n"
    checklist += f"- [ ] Manually download and install {kb_number}\n"
    checklist += f"- [ ] Verify installation via Windows Update history\n"
    checklist += f"- [ ] Reboot device\n"
    checklist += f"- [ ] Run vulnerability scan to confirm remediation\n"
    checklist += f"- [ ] Update ticket status once resolved\n\n"
    checklist += f"If unable to resolve, escalate to L2/L3 support.\n"
    
    note_data = {
        "text": checklist,
        "detailDescriptionFlag": False,
        "internalAnalysisFlag": False,
        "externalFlag": True
    }
    
    return cw_post(f"service/tickets/{ticket_id}/notes", note_data)


def handle_failed_device(ci):
    ci_id = ci.get("id")
    ci_name = ci.get("name")
    company_id = ci.get("company", {}).get("id")
    
    kb_str = get_ci_field(ci, "Pending_KB_Patches")
    kbs = [k.strip() for k in kb_str.split(",") if k.strip()] if kb_str else []
    
    if not kbs:
        return False
    
    print(f"Processing failed device: {ci_name} (CI: {ci_id})")
    
    handled = False
    
    for kb in kbs:
        status = get_ci_field(ci, f"Patch_Status_{kb}")
        
        if status == "Failed":
            print(f"  {kb} has failed - handling exception")
            
            ticket_ids_str = get_ci_field(ci, "Active_Vulnerability_Tickets")
            ticket_ids = [t.strip() for t in ticket_ids_str.split(",") if t.strip()] if ticket_ids_str else []
            
            for ticket_id in ticket_ids:
                ticket = cw_get(f"service/tickets/{ticket_id}")
                if not ticket or ticket.get("closedFlag"):
                    continue
                
                notes = cw_get(f"service/tickets/{ticket_id}/notes")
                if notes:
                    full_notes = "\n".join([n.get("text", "") for n in notes])
                    if kb in full_notes:
                        move_ticket_to_exception_board(ticket_id, kb, "Patch installation failed")
                        create_remediation_checklist(ticket_id, kb)
                        send_failure_notification(ticket_id, ci_name, kb)
                        
                        if company_id:
                            params = f"conditions=summary contains '{kb}' and company/id={company_id} and board/name='{MASTER_BOARD}' and closedFlag=false"
                            master_tickets = cw_get(f"service/tickets?{params}")
                            
                            if master_tickets and len(master_tickets) > 0:
                                add_failure_alert_to_master(master_tickets[0]["id"], ci_name, ci_id, kb, "Patch installation failed")
                        
                        handled = True
                        break
    
    return handled


def main():
    print("=" * 60)
    print("Bot 6: Exception Handler")
    print("=" * 60)
    
    params = "conditions=customFields/Pending_KB_Patches!=null&pageSize=1000"
    cis = cw_get(f"company/configurations?{params}")
    
    if not cis:
        print("No devices with pending patches found")
        return
    
    print(f"Found {len(cis)} devices with pending patches")
    
    handled_count = 0
    for ci in cis:
        try:
            if handle_failed_device(ci):
                handled_count += 1
        except Exception as e:
            print(f"ERROR processing CI {ci.get('id')}: {str(e)}")
    
    print(f"\nCompleted: {handled_count} failed devices handled")


if __name__ == "__main__":
    main()

