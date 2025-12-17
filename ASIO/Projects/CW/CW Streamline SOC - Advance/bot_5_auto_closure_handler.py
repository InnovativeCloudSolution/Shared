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


def remove_cves_for_kb(cve_list, kb_number):
    return [cve for cve in cve_list if cve.get("kb") != kb_number]


def remove_kb_from_list(kb_list, kb_number):
    return [kb for kb in kb_list if kb != kb_number]


def remove_ticket_from_list(ticket_list, ticket_id):
    return [tid for tid in ticket_list if str(tid) != str(ticket_id)]


def close_child_ticket(ci_id, ci_name, kb_number):
    ci = cw_get(f"company/configurations/{ci_id}")
    if not ci:
        print(f"  ERROR: Could not get CI {ci_id}")
        return False
    
    ticket_ids_str = get_ci_field(ci, "Active_Vulnerability_Tickets")
    ticket_ids = [t.strip() for t in ticket_ids_str.split(",") if t.strip()] if ticket_ids_str else []
    
    if not ticket_ids:
        print(f"  No active tickets for CI {ci_id}")
        return False
    
    closed_count = 0
    
    for ticket_id in ticket_ids:
        ticket = cw_get(f"service/tickets/{ticket_id}")
        if not ticket:
            continue
        
        if ticket.get("closedFlag"):
            continue
        
        notes = cw_get(f"service/tickets/{ticket_id}/notes")
        if notes:
            full_notes = "\n".join([n.get("text", "") for n in notes])
            if kb_number in full_notes:
                print(f"  Closing child ticket #{ticket_id} for {kb_number}")
                
                note_data = {
                    "text": f"Patch {kb_number} verified and installed. All CVEs remediated. Auto-closing ticket.",
                    "detailDescriptionFlag": False,
                    "internalAnalysisFlag": True
                }
                cw_post(f"service/tickets/{ticket_id}/notes", note_data)
                
                operations = [
                    {"op": "replace", "path": "status/name", "value": "Closed"},
                    {"op": "replace", "path": "closedFlag", "value": True}
                ]
                result = cw_patch(f"service/tickets/{ticket_id}", operations)
                
                if result:
                    closed_count += 1
                    
                    updated_ticket_list = remove_ticket_from_list(ticket_ids, ticket_id)
                    
                    operations = [
                        {"op": "replace", "path": "customFields/Active_Vulnerability_Tickets", "value": ",".join(map(str, updated_ticket_list))}
                    ]
                    cw_patch(f"company/configurations/{ci_id}", operations)
    
    return closed_count > 0


def cleanup_ci_fields(ci_id, kb_number):
    ci = cw_get(f"company/configurations/{ci_id}")
    if not ci:
        return False
    
    pending_cves_str = get_ci_field(ci, "Pending_CVEs")
    pending_cves = json.loads(pending_cves_str) if pending_cves_str else []
    
    pending_kbs_str = get_ci_field(ci, "Pending_KB_Patches")
    pending_kbs = [k.strip() for k in pending_kbs_str.split(",") if k.strip()] if pending_kbs_str else []
    
    updated_cves = remove_cves_for_kb(pending_cves, kb_number)
    updated_kbs = remove_kb_from_list(pending_kbs, kb_number)
    
    operations = [
        {"op": "replace", "path": "customFields/Pending_CVEs", "value": json.dumps(updated_cves)},
        {"op": "replace", "path": "customFields/Pending_KB_Patches", "value": ",".join(updated_kbs)},
        {"op": "replace", "path": f"customFields/Patch_Status_{kb_number}", "value": "Verified"}
    ]
    
    result = cw_patch(f"company/configurations/{ci_id}", operations)
    return result is not None


def check_master_ticket_complete(master_ticket_id, kb_number, company_id):
    params = f"conditions=company/id={company_id} and customFields/Patch_Status_{kb_number}!='Verified' and customFields/Patch_Status_{kb_number}!='Failed' and customFields/Pending_KB_Patches contains '{kb_number}'&pageSize=1"
    cis = cw_get(f"company/configurations?{params}")
    
    if not cis or len(cis) == 0:
        print(f"  All devices complete for {kb_number}! Marking master ticket #{master_ticket_id} as Complete")
        
        note_data = {
            "text": f"All devices have been verified or failed. Patch deployment complete.",
            "detailDescriptionFlag": False,
            "internalAnalysisFlag": True
        }
        cw_post(f"service/tickets/{master_ticket_id}/notes", note_data)
        
        operations = [
            {"op": "replace", "path": "status/name", "value": "Complete"}
        ]
        cw_patch(f"service/tickets/{master_ticket_id}", operations)
        
        return True
    
    return False


def process_verified_device(ci):
    ci_id = ci.get("id")
    ci_name = ci.get("name")
    
    kb_str = get_ci_field(ci, "Pending_KB_Patches")
    kbs = [k.strip() for k in kb_str.split(",") if k.strip()] if kb_str else []
    
    if not kbs:
        return False
    
    print(f"Processing verified device: {ci_name} (CI: {ci_id})")
    
    processed = False
    
    for kb in kbs:
        status = get_ci_field(ci, f"Patch_Status_{kb}")
        
        if status == "Verified":
            print(f"  {kb} is verified - closing tickets and cleaning up")
            
            if close_child_ticket(ci_id, ci_name, kb):
                cleanup_ci_fields(ci_id, kb)
                processed = True
                
                company_id = ci.get("company", {}).get("id")
                if company_id:
                    params = f"conditions=summary contains '{kb}' and company/id={company_id} and board/name='{MASTER_BOARD}' and closedFlag=false"
                    master_tickets = cw_get(f"service/tickets?{params}")
                    
                    if master_tickets and len(master_tickets) > 0:
                        check_master_ticket_complete(master_tickets[0]["id"], kb, company_id)
    
    return processed


def main():
    print("=" * 60)
    print("Bot 5: Auto-Closure Handler")
    print("=" * 60)
    
    params = "conditions=customFields/Pending_KB_Patches!=null&pageSize=1000"
    cis = cw_get(f"company/configurations?{params}")
    
    if not cis:
        print("No devices with pending patches found")
        return
    
    print(f"Found {len(cis)} devices with pending patches")
    
    processed_count = 0
    for ci in cis:
        try:
            if process_verified_device(ci):
                processed_count += 1
        except Exception as e:
            print(f"ERROR processing CI {ci.get('id')}: {str(e)}")
    
    print(f"\nCompleted: {processed_count} devices processed for closure")


if __name__ == "__main__":
    main()

