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


def find_devices_for_kb(company_id, kb_number):
    params = f"conditions=company/id={company_id} and customFields/Pending_KB_Patches contains '{kb_number}'&pageSize=1000"
    cis = cw_get(f"company/configurations?{params}")
    
    devices = []
    if cis:
        for ci in cis:
            pending_cves_str = get_ci_field(ci, "Pending_CVEs")
            pending_cves = json.loads(pending_cves_str) if pending_cves_str else []
            
            ticket_ids_str = get_ci_field(ci, "Active_Vulnerability_Tickets")
            ticket_ids = [t.strip() for t in ticket_ids_str.split(",") if t.strip()] if ticket_ids_str else []
            
            critical_count = sum(1 for cve in pending_cves if cve.get("severity") == "Critical" and cve.get("kb") == kb_number)
            high_count = sum(1 for cve in pending_cves if cve.get("severity") == "High" and cve.get("kb") == kb_number)
            medium_count = sum(1 for cve in pending_cves if cve.get("severity") == "Medium" and cve.get("kb") == kb_number)
            
            devices.append({
                "id": ci.get("id"),
                "name": ci.get("name"),
                "cve_list": [cve for cve in pending_cves if cve.get("kb") == kb_number],
                "critical_count": critical_count,
                "high_count": high_count,
                "medium_count": medium_count,
                "ticket_ids": ticket_ids
            })
    
    return devices


def create_master_ticket(company_id, kb_number):
    print(f"Creating master ticket for {kb_number}, Company {company_id}")
    
    devices = find_devices_for_kb(company_id, kb_number)
    
    if not devices:
        print(f"ERROR: No devices found for {kb_number}")
        return None
    
    total_devices = len(devices)
    total_cves = sum(len(d["cve_list"]) for d in devices)
    total_critical = sum(d["critical_count"] for d in devices)
    total_high = sum(d["high_count"] for d in devices)
    total_medium = sum(d["medium_count"] for d in devices)
    
    board_data = cw_get(f"service/boards?conditions=name='{MASTER_BOARD}'")
    if not board_data or len(board_data) == 0:
        print(f"ERROR: Board '{MASTER_BOARD}' not found")
        return None
    
    board_id = board_data[0]["id"]
    
    priority = 1 if total_critical > 0 else 2 if total_high > 0 else 3
    
    summary = f"Patch Deployment - {kb_number} - {total_devices} Devices - {total_cves} CVEs"
    
    description = f"## Patch Deployment Status\n\n"
    description += f"KB Number: {kb_number}\n"
    description += f"Total Devices: {total_devices}\n"
    description += f"Total CVEs: {total_cves} (Critical: {total_critical}, High: {total_high}, Medium: {total_medium})\n\n"
    description += f"Status Breakdown:\n"
    description += f"- Pending Patch: {total_devices}\n"
    description += f"- Patched: 0\n"
    description += f"- Rebooted: 0\n"
    description += f"- Verified: 0\n"
    description += f"- Failed: 0\n\n"
    description += f"Progress: 0% Complete\n\n"
    description += f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
    description += f"---\n\n"
    description += f"Affected Devices:\n"
    
    for device in devices:
        ticket_ref = f"#{device['ticket_ids'][0]}" if device['ticket_ids'] else "N/A"
        description += f"- {device['name']} (CI: {device['id']}, Ticket: {ticket_ref})\n"
    
    ticket_data = {
        "summary": summary,
        "board": {"id": board_id},
        "company": {"id": company_id},
        "priority": {"id": priority},
        "status": {"name": "New"},
        "type": {"name": "Problem"},
        "initialDescription": description
    }
    
    master_ticket = cw_post("service/tickets", ticket_data)
    
    if master_ticket:
        master_ticket_id = master_ticket["id"]
        print(f"SUCCESS: Created master ticket #{master_ticket_id}")
        
        for device in devices:
            for ticket_id in device["ticket_ids"]:
                operations = [{"op": "replace", "path": "parentTicketId", "value": master_ticket_id}]
                cw_patch(f"service/tickets/{ticket_id}", operations)
        
        return master_ticket_id
    else:
        print(f"ERROR: Failed to create master ticket for {kb_number}")
        return None


def find_companies_with_pending_kbs():
    print("Scanning for companies with pending KB patches...")
    
    companies_kbs = {}
    
    params = "conditions=customFields/Pending_KB_Patches!=null&pageSize=1000"
    cis = cw_get(f"company/configurations?{params}")
    
    if cis:
        for ci in cis:
            company_id = ci.get("company", {}).get("id")
            if not company_id:
                continue
            
            kb_str = get_ci_field(ci, "Pending_KB_Patches")
            kbs = [k.strip() for k in kb_str.split(",") if k.strip()] if kb_str else []
            
            if company_id not in companies_kbs:
                companies_kbs[company_id] = set()
            
            for kb in kbs:
                companies_kbs[company_id].add(kb)
    
    return companies_kbs


def main():
    print("=" * 60)
    print("Bot 2: Master Ticket Creator")
    print("=" * 60)
    
    companies_kbs = find_companies_with_pending_kbs()
    
    if not companies_kbs:
        print("No companies with pending KB patches found")
        return
    
    print(f"Found {len(companies_kbs)} companies with pending patches")
    
    created_count = 0
    for company_id, kbs in companies_kbs.items():
        print(f"\nCompany {company_id}: {len(kbs)} unique KB patches")
        
        for kb in kbs:
            params = f"conditions=summary contains '{kb}' and company/id={company_id} and board/name='{MASTER_BOARD}' and closedFlag=false"
            existing = cw_get(f"service/tickets?{params}")
            
            if existing and len(existing) > 0:
                print(f"SKIP: Master ticket already exists for {kb} (Ticket #{existing[0]['id']})")
                continue
            
            master_id = create_master_ticket(company_id, kb)
            if master_id:
                created_count += 1
    
    print(f"\nCompleted: {created_count} master tickets created")


if __name__ == "__main__":
    main()

