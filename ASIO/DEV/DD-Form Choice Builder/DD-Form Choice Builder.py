import requests
from msal import ConfidentialClientApplication
import pprint
import subprocess
import re
import string
import random
import json

# Azure AD app registration details
CLIENT_ID = input("Enter your Azure AD Client ID: ").strip()
CLIENT_SECRET = input("Enter your Azure AD Client Secret: ").strip()
TENANT_ID = 'f01b3dc6-43a4-4f59-aa9d-90d9658c715b' #input("Enter your Azure AD Tenant ID: ").strip()
ORGANIZATION_NAME = 'sherrinrentals.com.au' #input("Enter your organization name: ").strip()

# MS Graph API endpoints
AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'
SCOPE = ['https://graph.microsoft.com/.default']
GRAPH_API_BASE = 'https://graph.microsoft.com/beta'

def get_access_token():
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=SCOPE)
    if "access_token" in result:
        return result['access_token']
    else:
        raise Exception(f"Could not obtain access token: {result.get('error_description')}")
    
def get_exo_token():
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=['https://outlook.office365.com/.default'])
    if "access_token" in result:
        return result['access_token']
    else:
        raise Exception(f"Could not obtain EXO access token: {result.get('error_description')}")

def get_groups(token):
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    url = f'{GRAPH_API_BASE}/groups'
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        data = response.json().get('value', [])
        groups = []
        for group in data:
            group_name = group['displayName'].strip()
            group_name = group_name.replace('SG.', '', 1).strip()
            group_name = group_name.replace('.', ' ', 1).strip()
            print(f"Processing group: {group_name}")
            groups.append(f"{group_name} [{group['displayName']}]")
        return groups
    else:
        raise Exception(f"Error fetching groups: {response.status_code} - {response.text}")

def get_teams_groups(token):
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    url = f'{GRAPH_API_BASE}/groups?$filter=groupTypes/any(c:c eq \'Unified\')'
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        data = response.json().get('value', [])
        teams_groups = []
        for group in data:
            group_name = group['displayName'].strip()
            group_name = group_name.replace('SG.Teams.', '', 1).strip()
            group_name = group_name.replace('.', ' ', 1).strip()
            print(f"Processing Teams group: {group_name}")
            teams_groups.append(f"{group_name} [{group['displayName']}]")
    else:
        raise Exception(f"Error fetching Teams groups: {response.status_code} - {response.text}")
    
def get_sharepoint_groups(token):
    # Filter groups for SG.SharePoint prefix
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    url = f'{GRAPH_API_BASE}/groups?$filter=startswith(displayName, \'SG.SharePoint\')'
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        data = response.json().get('value', [])
        sharepoint_groups = []
        for group in data:
            group_name = group['displayName'].strip()
            group_name = group_name.replace('SG.SharePoint.', '', 1).strip()
            group_name = group_name.replace('.', ' ').strip()
            print(f"Processing SharePoint group: {group_name}")
            sharepoint_groups.append(f"{group_name} [{group['displayName']}]")
        return sharepoint_groups
    else:
        raise Exception(f"Error fetching SharePoint groups: {response.status_code} - {response.text}")
    
def get_apps_groups(token):
    # Filter groups for SG.App prefix
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    url = f'{GRAPH_API_BASE}/groups?$filter=startswith(displayName, \'SG.App\')'
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        data = response.json().get('value', [])
        app_groups = []
        for group in data:
            group_name = group['displayName'].strip()
            group_name = group_name.replace('SG.App.', '', 1).strip()
            group_name = group_name.replace('.', ' ', 1).strip()
            print(f"Processing App group: {group_name}")
            app_groups.append(f"{group_name} [{group['displayName']}]")
        return app_groups
    else:
        raise Exception(f"Error fetching App groups: {response.status_code} - {response.text}")
    
def get_license_groups(token):
    # Filter groups for SG.License prefix
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }   
    url = f'{GRAPH_API_BASE}/groups?$filter=startswith(displayName, \'SG.License\')'
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        data = response.json().get('value', [])
        license_groups = []
        for group in data:
            group_name = group['displayName'].strip()
            group_name = group_name.replace('SG.License.', '', 1).strip()
            group_name = group_name.replace('.', ' ', 1).strip()
            print(f"Processing License group: {group_name}")
            license_groups.append(f"{group_name} [{group['displayName']}]")
        return license_groups
    else:
        raise Exception(f"Error fetching License groups: {response.status_code} - {response.text}")
    
def get_shared_mailboxes_from_exo(token):
    # Run Powershell command to get shared mailboxes via Connect-ExchangeOnline
    command_1 = "$mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox | Select-Object DisplayName,PrimarySmtpAddress"
    command_2 = "$mailboxes | ForEach-Object { Write-Output( $_.DisplayName ); Write-Output( $_.PrimarySmtpAddress ) }"
    # Run the Powershell command using the EXO token
    try:
        result = subprocess.run(
            ["powershell", "-Command", f"Import-Module ExchangeOnlineManagement; $token = '{token}'; $organisation = '{ORGANIZATION_NAME}'; Connect-ExchangeOnline -AccessToken $token -Organization $organisation -ShowBanner:$false; {command_1}; {command_2}"],
            capture_output=True,
            text=True,
            check=True
        )
        output = result.stdout.strip()
        if output:
            output = output.splitlines()
            #combine display names and primary SMTP addresses into a list of dictionaries
            shared_mailboxes = []
            for i in range(0, len(output), 2):
                if i + 1 < len(output):
                    shared_mailboxes.append(f"{output[i]} (FA) [{output[i + 1]}:FullAccess]")
                    shared_mailboxes.append(f"{output[i]} (RO) [{output[i + 1]}:ReadPermission]")
            return shared_mailboxes
        else:
            return []
    except subprocess.CalledProcessError as e:
        raise Exception(f"Error fetching shared mailboxes: {e.stderr.strip()}")
    
def get_distribution_groups_from_exo(token):
    # Run Powershell command to get distribution groups via Connect-ExchangeOnline
    command_1 = "$distributionGroups = Get-DistributionGroup -ResultSize Unlimited | Select-Object PrimarySmtpAddress,DisplayName"
    command_2 = "$distributionGroups | ForEach-Object { Write-Output( $_.DisplayName ); Write-Output( $_.PrimarySmtpAddress ) }"
    # Run the Powershell command using the EXO token
    try:
        result = subprocess.run(
            ["powershell", "-Command", f"Import-Module ExchangeOnlineManagement; $token = '{token}'; $organisation = '{ORGANIZATION_NAME}'; Connect-ExchangeOnline -AccessToken $token -Organization $organisation -ShowBanner:$false; {command_1}; {command_2}"],
            capture_output=True,
            text=True,
            check=True
        )
        output = result.stdout.strip()
        if output:
            output = output.splitlines()
            #combine display names and primary SMTP addresses into a list of dictionaries
            distribution_groups = []
            for i in range(0, len(output), 2):
                if i + 1 < len(output):
                    distribution_groups.append(f"{output[i]} (Member) [{output[i + 1]}:Member]")
                    distribution_groups.append(f"{output[i]} (Owner) [{output[i + 1]}:Owner]")
            return distribution_groups
        else:
            return []
    except subprocess.CalledProcessError as e:
        raise Exception(f"Error fetching distribution groups: {e.stderr.strip()}")
    
def format_to_dd_choices(input):
    choices = {
        "choices": []
    }
    for line in input:
        identifier = generate_6_char_string()
        choice = {
            "name": line.strip(),
            "identifier": identifier,
        }
        choices["choices"].append(choice)
    return choices

def generate_6_char_string():
    chars = string.ascii_lowercase + string.digits
    return ''.join(random.choices(chars, k=6))

if __name__ == "__main__":
    graph_token = get_access_token()
    exo_token = get_exo_token()

    distribution_list_groups = get_distribution_groups_from_exo(exo_token)
    shared_mailbox_groups = get_shared_mailboxes_from_exo(exo_token)
    sharepoint_groups = get_sharepoint_groups(graph_token)
    app_groups = get_apps_groups(graph_token)
    #teams_groups = get_teams_groups(graph_token)
    license_groups = get_license_groups(graph_token)

    distribution_list_choices = format_to_dd_choices(distribution_list_groups)
    shared_mailbox_choices = format_to_dd_choices(shared_mailbox_groups)
    sharepoint_choices = format_to_dd_choices(sharepoint_groups)
    app_choices = format_to_dd_choices(app_groups)
    #teams_choices = format_to_dd_choices(teams_groups)
    license_choices = format_to_dd_choices(license_groups)

    with open("C:\\Github\\MIT\\PowerShell Scripts\\MIT\\ASIO\\DEV\\ASIO\\DD-Form Choice Builder\\distribution_list_choices.json", "w") as f:
        json.dump(distribution_list_choices, f, indent=2)
    with open("C:\\Github\\MIT\\PowerShell Scripts\\MIT\\ASIO\\DEV\\ASIO\\DD-Form Choice Builder\\shared_mailbox_choices.json", "w") as f:
        json.dump(shared_mailbox_choices, f, indent=2)
    with open("C:\\Github\\MIT\\PowerShell Scripts\\MIT\\ASIO\\DEV\\ASIO\\DD-Form Choice Builder\\sharepoint_choices.json", "w") as f:
        json.dump(sharepoint_choices, f, indent=2)
    with open("C:\\Github\\MIT\\PowerShell Scripts\\MIT\\ASIO\\DEV\\ASIO\\DD-Form Choice Builder\\app_choices.json", "w") as f:
        json.dump(app_choices, f, indent=2)
    #with open("C:\\Github\\MIT\\PowerShell Scripts\\MIT\\ASIO\\DEV\\ASIO\\DD-Form Choice Builder\\teams_choices.json", "w") as f:
    #    json.dump(teams_choices, f, indent=2)
    with open("C:\\Github\\MIT\\PowerShell Scripts\\MIT\\ASIO\\DEV\\ASIO\\DD-Form Choice Builder\\license_choices.json", "w") as f:
        json.dump(license_choices, f, indent=2)
