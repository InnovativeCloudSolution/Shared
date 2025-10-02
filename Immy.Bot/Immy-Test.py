import requests

tenant_url = "manganoit.com.au"
client_id = "812408b3-6b87-418b-95ed-b036e2a67402"
client_secret = "Pmk8Q~FOA3Kd5QOR~9Bm.xiRBWNyJh1.TxjEcbIV"
base_url = "https://mit.immy.bot"
user_email = "Brad.Swan@manganoit.com.au"
software_name = "Google Chrome"

def get_immybot_api_auth_token(tenant_url, client_id, client_secret, base_url):
    url = f"https://login.microsoftonline.com/{tenant_url}/oauth2/v2.0/token"
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    data = {
        'client_id': client_id,
        'client_secret': client_secret,
        'grant_type': 'client_credentials',
        'scope': f"{base_url}/.default"
    }
    response = requests.post(url, headers=headers, data=data)
    response.raise_for_status()
    return response.json()

def invoke_immybot_rest_method(base_url, endpoint, method, bearer_token, body=None, params=None):
    endpoint = endpoint.lstrip('/')
    url = f"{base_url}/{endpoint}"
    headers = {
        'Authorization': bearer_token,
        'Content-Type': 'application/json'
    }
    
    if method.upper() == "GET":
        response = requests.get(url, headers=headers, params=params)
    elif method.upper() == "POST":
        response = requests.post(url, headers=headers, json=body)
    elif method.upper() == "PUT":
        response = requests.put(url, headers=headers, json=body)
    elif method.upper() == "DELETE":
        response = requests.delete(url, headers=headers)
    else:
        raise ValueError(f"Unsupported HTTP method: {method}")
    
    response.raise_for_status()
    return response.json()

def get_immy_endpoint(base_url, bearer_token):
    computers = []
    skip = 0
    page_size = 100
    more_records = True

    while more_records:
        params = {
            "skip": skip,
            "take": page_size
        }
        response = invoke_immybot_rest_method(base_url, "/api/v1/computers/dx", "GET", bearer_token, params=params)
        
        batch = response.get("data", [])
        computers.extend(batch)

        if len(batch) < page_size:
            more_records = False
        else:
            skip += page_size

    return computers

def get_immy_software(base_url, bearer_token, software_name):
    softwares = invoke_immybot_rest_method(base_url, "/api/v1/software/local", "GET", bearer_token)
    selected_software = next(
        (s for s in softwares if (s.get('name') or '').lower() == software_name.lower()), 
        None
    )
    
    if not selected_software:
        softwares = invoke_immybot_rest_method(base_url, "/api/v1/software/global", "GET", bearer_token)
        selected_software = next(
            (s for s in softwares if (s.get('name') or '').lower() == software_name.lower()), 
            None
        )
    
    return selected_software

def push_immy_software(base_url, bearer_token, software_id, selected_computers):
    maintenance_payload = {
        "fullMaintenance": False,
        "resolutionOnly": False,
        "detectionOnly": False,
        "inventoryOnly": False,
        "runInventoryInDetection": False,
        "cacheOnly": False,
        "useWinningDeployment": False,
        "deploymentId": None,
        "deploymentType": None,
        "maintenanceParams": {
            "maintenanceIdentifier": software_id,
            "maintenanceType": 0,
            "repair": False,
            "desiredSoftwareState": 5,
            "maintenanceTaskMode": 0
        },
        "skipBackgroundJob": True,
        "rebootPreference": 1,
        "scheduleExecutionAfterActiveHours": False,
        "useComputersTimezoneForExecution": False,
        "offlineBehavior": 2,
        "suppressRebootsDuringBusinessHours": False,
        "sendDetectionEmail": False,
        "sendDetectionEmailWhenAllActionsAreCompliant": False,
        "sendFollowUpEmail": False,
        "sendFollowUpOnlyIfActionNeeded": False,
        "showRunNowButton": False,
        "showPostponeButton": False,
        "showMaintenanceActions": False,
        "computers": [{"computerId": comp.get("id")} for comp in selected_computers],
        "tenants": []
    }

    invoke_immybot_rest_method(base_url, "/api/v1/run-immy-service", "POST", bearer_token, body=maintenance_payload)

def main():
    token_response = get_immybot_api_auth_token(tenant_url, client_id, client_secret, base_url)
    bearer_token = f"Bearer {token_response['access_token']}"

    selected_software = get_immy_software(base_url, bearer_token, software_name)
    
    if not selected_software:
        print(f"Software '{software_name}' not found.")
        return

    print(f"Matching software found: {selected_software.get('name')}")  # <-- Add this

    all_computers = get_immy_endpoint(base_url, bearer_token)

    selected_computers = [
        {
            "id": computer.get("id"),
            "computerName": computer.get("computerName"),
            "primaryUserEmail": computer.get("primaryUserEmail")
        }
        for computer in all_computers
        if (computer.get("primaryUserEmail") or "").lower() == user_email.lower()
    ]

    if selected_computers:
        print("Matching endpoints found:")
        for comp in selected_computers:
            print(comp)
        
        software_id = selected_software.get("id")
        push_immy_software(base_url, bearer_token, software_id, selected_computers)
        print(f"Pushed '{software_name}' to selected endpoints.")
    else:
        print(f"No endpoints found with primaryUserEmail = {user_email}")
        print("All available endpoints and their primaryUserEmail values:")
        for computer in all_computers:
            print({
                "computerName": computer.get("computerName"),
                "primaryUserEmail": computer.get("primaryUserEmail")
            })

if __name__ == "__main__":
    main()