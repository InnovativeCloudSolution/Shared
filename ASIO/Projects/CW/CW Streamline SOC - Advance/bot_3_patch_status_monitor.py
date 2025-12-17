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

CW_RMM_URL = os.getenv("CW_RMM_URL", config.get('ConnectWise', 'cw_rmm_url'))
CW_RMM_KEY = os.getenv("CW_RMM_API_KEY", config.get('ConnectWise', 'cw_rmm_api_key'))


def get_auth_header():
    auth_string = f"{CW_COMPANY}+{CW_PUBLIC_KEY}:{CW_PRIVATE_KEY}"
    encoded = base64.b64encode(auth_string.encode()).decode()
    return {
        "Authorization": f"Basic {encoded}",
        "Content-Type": "application/json",
        "clientId": CW_CLIENT_ID
    }


def get_rmm_header():
    return {
        "Authorization": f"Bearer {CW_RMM_KEY}",
        "Content-Type": "application/json"
    }


def cw_get(endpoint):
    url = f"{CW_URL}/v4_6_release/apis/3.0/{endpoint}"
    response = requests.get(url, headers=get_auth_header(), timeout=30)
    if response.status_code == 200:
        return response.json()
    return None


def cw_patch(endpoint, data):
    url = f"{CW_URL}/v4_6_release/apis/3.0/{endpoint}"
    response = requests.patch(url, headers=get_auth_header(), json=data, timeout=30)
    if response.status_code == 200:
        return response.json()
    return None


def cw_post(endpoint, data):
    url = f"{CW_URL}/v4_6_release/apis/3.0/{endpoint}"
    response = requests.post(url, headers=get_auth_header(), json=data, timeout=30)
    if response.status_code in [200, 201]:
        return response.json()
    return None


def rmm_get(endpoint):
    if not CW_RMM_URL or not CW_RMM_KEY:
        return None
    
    url = f"{CW_RMM_URL}/api/{endpoint}"
    try:
        response = requests.get(url, headers=get_rmm_header(), timeout=30)
        if response.status_code == 200:
            return response.json()
    except:
        pass
    return None


def get_ci_field(ci_data, field_name):
    custom_fields = ci_data.get("customFields", [])
    for field in custom_fields:
        if field.get("caption") == field_name:
            return field.get("value", "")
    return ""


def check_patch_installed(ci_id, kb_number):
    installed_patches = rmm_get(f"devices/{ci_id}/patches")
    
    if installed_patches:
        for patch in installed_patches:
            if kb_number.lower() in patch.get("name", "").lower() or kb_number.lower() in patch.get("id", "").lower():
                return True
    
    return False


def check_device_rebooted(ci_id, patch_installed_date):
    device_info = rmm_get(f"devices/{ci_id}")
    
    if device_info and patch_installed_date:
        last_reboot = device_info.get("lastRebootTime")
        if last_reboot:
            try:
                reboot_time = datetime.fromisoformat(last_reboot.replace('Z', '+00:00'))
                install_time = datetime.fromisoformat(patch_installed_date.replace('Z', '+00:00'))
                return reboot_time > install_time
            except:
                pass
    
    return False


def update_patch_status(ci_id, kb_number, new_status):
    operations = [
        {"op": "replace", "path": f"customFields/Patch_Status_{kb_number}", "value": new_status}
    ]
    
    if new_status == "Patched":
        operations.append({"op": "replace", "path": f"customFields/Patch_Installed_Date_{kb_number}", "value": datetime.now().strftime("%Y-%m-%dT%H:%M:%S")})
    
    result = cw_patch(f"company/configurations/{ci_id}", operations)
    return result is not None


def update_ticket_status(ticket_id, status_message):
    note_data = {
        "text": status_message,
        "detailDescriptionFlag": False,
        "internalAnalysisFlag": True
    }
    return cw_post(f"service/tickets/{ticket_id}/notes", note_data)


def process_device(ci):
    ci_id = ci.get("id")
    ci_name = ci.get("name")
    
    kb_str = get_ci_field(ci, "Pending_KB_Patches")
    kbs = [k.strip() for k in kb_str.split(",") if k.strip()] if kb_str else []
    
    if not kbs:
        return False
    
    print(f"Processing device: {ci_name} (CI: {ci_id}) - {len(kbs)} pending KB patches")
    
    updated = False
    
    for kb in kbs:
        current_status = get_ci_field(ci, f"Patch_Status_{kb}")
        
        if current_status == "Verified" or current_status == "Failed":
            continue
        
        patch_installed = check_patch_installed(ci_id, kb)
        
        if patch_installed:
            if current_status == "Pending":
                print(f"  {kb}: Pending -> Patched")
                update_patch_status(ci_id, kb, "Patched")
                updated = True
            elif current_status == "Patched":
                patch_date = get_ci_field(ci, f"Patch_Installed_Date_{kb}")
                device_rebooted = check_device_rebooted(ci_id, patch_date)
                
                if device_rebooted:
                    print(f"  {kb}: Patched -> Rebooted")
                    update_patch_status(ci_id, kb, "Rebooted")
                    updated = True
            elif current_status == "Rebooted":
                print(f"  {kb}: Rebooted -> Verified (manual scan required)")
                update_patch_status(ci_id, kb, "Verified")
                updated = True
        else:
            if current_status != "Pending":
                print(f"  {kb}: {current_status} -> Failed (patch not found)")
                update_patch_status(ci_id, kb, "Failed")
                updated = True
    
    return updated


def main():
    print("=" * 60)
    print("Bot 3: Patch Status Monitor")
    print("=" * 60)
    
    params = "conditions=customFields/Pending_KB_Patches!=null&pageSize=1000"
    cis = cw_get(f"company/configurations?{params}")
    
    if not cis:
        print("No devices with pending patches found")
        return
    
    print(f"Found {len(cis)} devices with pending patches")
    
    updated_count = 0
    for ci in cis:
        try:
            if process_device(ci):
                updated_count += 1
        except Exception as e:
            print(f"ERROR processing CI {ci.get('id')}: {str(e)}")
    
    print(f"\nCompleted: {updated_count} devices updated")


if __name__ == "__main__":
    main()

