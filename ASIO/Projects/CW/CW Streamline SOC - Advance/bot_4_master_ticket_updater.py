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


def extract_kb_from_summary(summary):
    kb_match = re.search(r'KB(\d+)', summary)
    if kb_match:
        return f"KB{kb_match.group(1)}"
    return None


def get_status_counts_for_kb(company_id, kb_number):
    params = f"conditions=company/id={company_id} and customFields/Pending_KB_Patches contains '{kb_number}'&pageSize=1000"
    cis = cw_get(f"company/configurations?{params}")
    
    status_counts = {
        "Pending": 0,
        "Patched": 0,
        "Rebooted": 0,
        "Verified": 0,
        "Failed": 0
    }
    
    devices_by_status = {
        "Pending": [],
        "Patched": [],
        "Rebooted": [],
        "Verified": [],
        "Failed": []
    }
    
    if cis:
        for ci in cis:
            status = get_ci_field(ci, f"Patch_Status_{kb_number}")
            
            if not status:
                status = "Pending"
            
            if status in status_counts:
                status_counts[status] += 1
                devices_by_status[status].append({
                    "id": ci.get("id"),
                    "name": ci.get("name")
                })
    
    return status_counts, devices_by_status


def build_status_update_note(kb_number, status_counts, devices_by_status):
    total = sum(status_counts.values())
    
    if total == 0:
        return None
    
    progress = 0
    if total > 0:
        completed = status_counts.get("Verified", 0) + status_counts.get("Failed", 0)
        progress = int((completed / total) * 100)
    
    note = f"## Patch Deployment Status - UPDATED\n\n"
    note += f"KB Number: {kb_number}\n"
    note += f"Total Devices: {total}\n\n"
    note += f"Status Breakdown:\n"
    note += f"- Pending: {status_counts['Pending']} ({int(status_counts['Pending']/total*100) if total > 0 else 0}%)\n"
    note += f"- Patched: {status_counts['Patched']} ({int(status_counts['Patched']/total*100) if total > 0 else 0}%)\n"
    note += f"- Rebooted: {status_counts['Rebooted']} ({int(status_counts['Rebooted']/total*100) if total > 0 else 0}%)\n"
    note += f"- Verified: {status_counts['Verified']} ({int(status_counts['Verified']/total*100) if total > 0 else 0}%)\n"
    note += f"- Failed: {status_counts['Failed']} ({int(status_counts['Failed']/total*100) if total > 0 else 0}%)\n\n"
    note += f"Progress: {progress}% Complete\n\n"
    note += f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
    
    if devices_by_status["Pending"]:
        note += f"Critical Devices Pending ({len(devices_by_status['Pending'])}):\n"
        for device in devices_by_status["Pending"][:10]:
            note += f"- {device['name']} (CI: {device['id']})\n"
        if len(devices_by_status["Pending"]) > 10:
            note += f"- ...and {len(devices_by_status['Pending']) - 10} more\n"
        note += "\n"
    
    if devices_by_status["Failed"]:
        note += f"Failed Installations ({len(devices_by_status['Failed'])}):\n"
        for device in devices_by_status["Failed"]:
            note += f"- {device['name']} (CI: {device['id']})\n"
        note += "\n"
    
    return note


def update_master_ticket(ticket_id, company_id, kb_number):
    print(f"Updating master ticket #{ticket_id} for {kb_number}")
    
    status_counts, devices_by_status = get_status_counts_for_kb(company_id, kb_number)
    
    note_text = build_status_update_note(kb_number, status_counts, devices_by_status)
    
    if not note_text:
        print(f"  ERROR: No devices found for {kb_number}")
        return False
    
    note_data = {
        "text": note_text,
        "detailDescriptionFlag": False,
        "internalAnalysisFlag": True
    }
    
    result = cw_post(f"service/tickets/{ticket_id}/notes", note_data)
    
    total = sum(status_counts.values())
    completed = status_counts.get("Verified", 0) + status_counts.get("Failed", 0)
    
    if total > 0 and completed == total:
        print(f"  All devices complete - updating status to 'Complete'")
        operations = [
            {"op": "replace", "path": "status/name", "value": "Complete"}
        ]
        cw_patch(f"service/tickets/{ticket_id}", operations)
    elif status_counts.get("Patched", 0) > 0 or status_counts.get("Rebooted", 0) > 0:
        operations = [
            {"op": "replace", "path": "status/name", "value": "In Progress"}
        ]
        cw_patch(f"service/tickets/{ticket_id}", operations)
    
    print(f"  SUCCESS: Updated master ticket with current status")
    return True


def main():
    print("=" * 60)
    print("Bot 4: Master Ticket Updater")
    print("=" * 60)
    
    params = f"conditions=board/name='{MASTER_BOARD}' and closedFlag=false&pageSize=100"
    master_tickets = cw_get(f"service/tickets?{params}")
    
    if not master_tickets:
        print("No open master tickets found")
        return
    
    print(f"Found {len(master_tickets)} open master tickets")
    
    updated_count = 0
    for ticket in master_tickets:
        try:
            ticket_id = ticket["id"]
            summary = ticket["summary"]
            company_id = ticket.get("company", {}).get("id")
            
            kb_number = extract_kb_from_summary(summary)
            
            if not kb_number:
                print(f"SKIP: Could not extract KB from ticket #{ticket_id} summary")
                continue
            
            if update_master_ticket(ticket_id, company_id, kb_number):
                updated_count += 1
        except Exception as e:
            print(f"ERROR updating ticket {ticket.get('id')}: {str(e)}")
    
    print(f"\nCompleted: {updated_count} master tickets updated")


if __name__ == "__main__":
    main()

