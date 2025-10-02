import sys
import random
import re
import subprocess
import os
import time
import urllib.parse
import requests
from collections import defaultdict
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

cwpsa_base_url = "https://au.myconnectwise.net/v4_6_release/apis/3.0"
msgraph_base_url = "https://graph.microsoft.com/v1.0"
vault_name = "mit-azu1-prod1-akv1"
data_to_log = {}
log.info("Static variables set")

def record_result(log, level, message):
    log.result_message(level, message)

    if level == ResultLevel.WARNING:
        data_to_log["Result"] = "Fail"
    elif level == ResultLevel.SUCCESS:
        if "Result" not in data_to_log:
            data_to_log["Result"] = "Success"

def execute_api_call(log, http_client, method, endpoint, data=None, retries=5, integration_name=None, headers=None, params=None):
    base_delay = 5
    log.info(f"Executing API call: {method.upper()} {endpoint}")
    for attempt in range(retries):
        try:
            if integration_name:
                response = (
                    getattr(http_client.third_party_integration(integration_name), method)(url=endpoint, json=data)
                    if data else getattr(http_client.third_party_integration(integration_name), method)(url=endpoint)
                )
            else:
                request_args = {"url": endpoint}
                if params:
                    request_args["params"] = params
                if headers:
                    request_args["headers"] = headers
                if data:
                    if (headers and headers.get("Content-Type") == "application/x-www-form-urlencoded"):
                        request_args["data"] = data
                    else:
                        request_args["json"] = data
                response = getattr(requests, method)(**request_args)

            if 200 <= response.status_code < 300:
                return response
            elif response.status_code in [429, 503]:
                retry_after = response.headers.get("Retry-After")
                wait_time = int(retry_after) if retry_after else base_delay * (2 ** attempt) + random.uniform(0, 3)
                log.warning(f"Rate limit exceeded. Retrying in {wait_time:.2f} seconds")
                time.sleep(wait_time)
            elif 400 <= response.status_code < 500:
                if response.status_code == 404:
                    log.warning(f"Skipping non-existent resource [{endpoint}]")
                    return None
                log.error(f"Client error Status: {response.status_code}, Response: {response.text}")
                return response
            elif 500 <= response.status_code < 600:
                log.warning(f"Server error Status: {response.status_code}, attempt {attempt + 1} of {retries}")
                time.sleep(base_delay * (2 ** attempt) + random.uniform(0, 3))
            else:
                log.error(f"Unexpected response Status: {response.status_code}, Response: {response.text}")
                return response

        except Exception as e:
            log.exception(e, f"Exception during API call to {endpoint}")
            return None
    return None

def execute_powershell(log, command, shell="pwsh"):
    try:
        if shell not in ("pwsh", "powershell"):
            raise ValueError("Invalid shell specified. Use 'pwsh' or 'powershell'")

        result = subprocess.run([shell, "-Command", command], capture_output=True, text=True)

        if result.returncode == 0 and not result.stderr.strip():
            log.info(f"{shell} command executed successfully")
            return True, result.stdout.strip()
        else:
            error_message = result.stderr.strip() or result.stdout.strip()
            log.error(f"{shell} execution failed: {error_message}")
            return False, error_message

    except Exception as e:
        log.exception(e, f"Exception occurred during {shell} execution")
        return False, str(e)

def get_secret_value(log, http_client, vault_name, secret_name):
    log.info(f"Fetching secret [{secret_name}] from Key Vault [{vault_name}]")

    secret_url = f"https://{vault_name}.vault.azure.net/secrets/{secret_name}?api-version=7.3"
    response = execute_api_call(log, http_client, "get", secret_url, integration_name="custom_wf_oauth2_client_creds")

    if response and response.status_code == 200:
        secret_value = response.json().get("value", "")
        if secret_value:
            log.info(f"Successfully retrieved secret [{secret_name}]")
            return secret_value

    log.error(f"Failed to retrieve secret [{secret_name}] Status code: {response.status_code if response else 'N/A'}")
    return ""

def get_tenant_id_from_domain(log, http_client, azure_domain):
    try:
        config_url = f"https://login.windows.net/{azure_domain}/.well-known/openid-configuration"
        log.info(f"Fetching OpenID configuration from [{config_url}]")

        response = execute_api_call(log, http_client, "get", config_url)

        if response and response.status_code == 200:
            token_endpoint = response.json().get("token_endpoint", "")
            tenant_id = token_endpoint.split("/")[3] if token_endpoint else ""
            if tenant_id:
                log.info(f"Successfully extracted tenant ID [{tenant_id}]")
                return tenant_id

        log.error(f"Failed to extract tenant ID from domain [{azure_domain}]")
        return ""
    except Exception as e:
        log.exception(e, "Exception while extracting tenant ID from domain")
        return ""

def get_access_token(log, http_client, tenant_id, client_id, client_secret, scope="https://graph.microsoft.com/.default"):
    log.info(f"Requesting access token for scope [{scope}]")

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    payload = urllib.parse.urlencode({
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": scope
    })
    headers = {"Content-Type": "application/x-www-form-urlencoded"}

    response = execute_api_call(log, http_client, "post", token_url, data=payload, retries=3, headers=headers)

    if response and response.status_code == 200:
        token_data = response.json()
        access_token = str(token_data.get("access_token", "")).strip()
        log.info(f"Access token length: {len(access_token)}")
        log.info(f"Access token preview: {access_token[:30]}...")

        if not isinstance(access_token, str) or "." not in access_token:
            log.error("Access token is invalid or malformed")
            return ""

        log.info("Successfully retrieved access token")
        return access_token

    log.error(f"Failed to retrieve access token Status code: {response.status_code if response else 'N/A'}")
    return ""

def get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number):
    log.info(f"Retrieving company details for ticket [{ticket_number}]")
    ticket_endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}"
    ticket_response = execute_api_call(log, http_client, "get", ticket_endpoint, integration_name="cw_psa")

    if ticket_response and ticket_response.status_code == 200:
        ticket_data = ticket_response.json()
        company = ticket_data.get("company", {})

        company_id = company["id"]
        company_identifier = company["identifier"]
        company_name = company["name"]

        log.info(f"Company ID: [{company_id}], Identifier: [{company_identifier}], Name: [{company_name}]")

        company_endpoint = f"{cwpsa_base_url}/company/companies/{company_id}"
        company_response = execute_api_call(log, http_client, "get", company_endpoint, integration_name="cw_psa")

        company_type = ""
        if company_response and company_response.status_code == 200:
            company_data = company_response.json()
            company_type = company_data.get("type", {}).get("name", "")
            log.info(f"Company type for ID [{company_id}]: [{company_type}]")
        else:
            log.warning(f"Unable to retrieve company type for ID [{company_id}]")

        return company_identifier, company_name, company_id, company_type

    elif ticket_response:
        log.error(
            f"Failed to retrieve ticket [{ticket_number}] "
            f"Status: {ticket_response.status_code}, Body: {ticket_response.text}"
        )
    else:
        log.error(f"No response received when retrieving ticket [{ticket_number}]")

    return "", "", 0, ""

def validate_mit_authentication(log, http_client, vault_name, auth_code):
    if not auth_code:
        log.result_message(ResultLevel.FAILED, "Authentication code input is required for MIT")
        return False

    expected_code = get_secret_value(log, http_client, vault_name, "MIT-AuthenticationCode")
    if not expected_code:
        log.result_message(ResultLevel.FAILED, "Failed to retrieve expected authentication code for MIT")
        return False

    if auth_code.strip() != expected_code.strip():
        log.result_message(ResultLevel.FAILED, "Provided authentication code is incorrect")
        return False

    return True

def get_aad_user_data(log, http_client, msgraph_base_url, user_identifier, token):
    log.info(f"Resolving user ID and email for [{user_identifier}]")
    headers = {"Authorization": f"Bearer {token}"}
    if re.fullmatch(
            r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}",
            user_identifier,
    ):
        endpoint = f"{msgraph_base_url}/users/{user_identifier}"
        response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
        if response and response.status_code == 200:
            user = response.json()
            return (
                user.get("id", ""),
                user.get("userPrincipalName", ""),
                user.get("onPremisesSamAccountName", ""),
            )
        return "", "", ""
    filters = [
        f"startswith(displayName,'{user_identifier}')",
        f"startswith(userPrincipalName,'{user_identifier}')",
        f"startswith(mail,'{user_identifier}')",
    ]
    filter_query = " or ".join(filters)
    endpoint = f"{msgraph_base_url}/users?$filter={urllib.parse.quote(filter_query)}"
    response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
    if response and response.status_code == 200:
        users = response.json().get("value", [])
        if len(users) > 1:
            log.error(f"Multiple users found for [{user_identifier}]")
            return users
        if users:
            user = users[0]
            return (
                user.get("id", ""),
                user.get("userPrincipalName", ""),
                user.get("onPremisesSamAccountName", ""),
            )
    log.error(f"Failed to resolve user ID and email for [{user_identifier}]")
    return "", "", ""

def get_exo_distributiongroup(log, group_name, exo_access_token, azure_domain):
    log.info(f"Checking if Exchange Online distribution group [{group_name}] exists")
    ps_command = f"""
    $ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false
    try {{
        Get-DistributionGroup -Identity '{group_name}' -ErrorAction Stop | Out-Null
        Write-Output "exists"
    }} catch {{
        Write-Output "not found"
    }}
    Disconnect-ExchangeOnline -Confirm:$false
    """
    success, output = execute_powershell(log, ps_command)
    if not success or "not found" in output.lower():
        log.warning(f"Group [{group_name}] could not be found in Exchange Online")
        return False
    return True

def add_exo_distributiongroup_permissions(log, group_name, user_upn, permission_type, exo_access_token, azure_domain):
    log.info(f"Adding user [{user_upn}] to group [{group_name}] as [{permission_type}]")
    ps_command = f"""
    $ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false
    try {{
        $group = Get-DistributionGroup -Identity '{group_name}' -ErrorAction Stop
        $groupName = $group.Identity

        $existingMembers = Get-DistributionGroupMember -Identity $groupName -ResultSize Unlimited | Where-Object {{ $_.PrimarySmtpAddress -eq '{user_upn}' }}
        $existingOwners = $group.ManagedBy

        if ('{permission_type}' -eq 'Member') {{
            if ($existingMembers) {{
                Write-Output "skipped : {user_upn}:$groupName:{permission_type} - Already a member"
            }} else {{
                Add-DistributionGroupMember -Identity $groupName -Member '{user_upn}' -BypassSecurityGroupManagerCheck
                Write-Output "success : {user_upn}:$groupName:{permission_type}"
            }}
        }} elseif ('{permission_type}' -eq 'Owner') {{
            if ($existingOwners -contains '{user_upn}') {{
                Write-Output "skipped : {user_upn}:$groupName:{permission_type} - Already an owner"
            }} else {{
                Set-DistributionGroup -Identity $groupName -ManagedBy @{{add = '{user_upn}' }} -BypassSecurityGroupManagerCheck
                Write-Output "success : {user_upn}:$groupName:{permission_type}"
            }}
        }} else {{
            Write-Output "failed : {user_upn}:$groupName:{permission_type} - Unsupported permission type"
        }}
    }} catch {{
        Write-Output "failed : {user_upn}:{group_name}:{permission_type} - $($_.Exception.Message)"
    }}
    Disconnect-ExchangeOnline -Confirm:$false
    """
    return execute_powershell(log, ps_command)

def add_exo_distributiongroup_permissions_contact(log, group_list, contact_email, exo_access_token, azure_domain):
    results = []
    for group_name in group_list:
        if not get_exo_distributiongroup(log, group_name, exo_access_token, azure_domain):
            results.append(f"failed : {contact_email}:{group_name} - Group not found")
            continue

        log.info(f"Adding mail contact [{contact_email}] to group [{group_name}]")
        ps_command = f"""
        $ErrorActionPreference = 'Stop'
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false
        try {{
            Add-DistributionGroupMember -Identity '{group_name}' -Member '{contact_email}' -BypassSecurityGroupManagerCheck
            Write-Output "success : {contact_email}:{group_name}"
        }} catch {{
            Write-Output "failed : {contact_email}:{group_name} - $($_.Exception.Message)"
        }}
        Disconnect-ExchangeOnline -Confirm:$false
        """
        success, output = execute_powershell(log, ps_command)
        results.append(output.strip() if output else f"failed : {contact_email}:{group_name} - No response")
    return results

def remove_exo_distributiongroup_permissions(log, group_name, user_upn, role, exo_access_token, azure_domain):
    log.info(f"Removing user [{user_upn}] from group [{group_name}] as [{role}]")
    ps_command = f"""
    $ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false
    try {{
        $group = Get-DistributionGroup -Identity '{group_name}' -ErrorAction Stop
        $groupName = $group.Identity
        $success = $false
        $skipped = $false

        if ('{role}' -eq 'Member') {{
            $isMember = Get-DistributionGroupMember -Identity $groupName -ResultSize Unlimited | Where-Object {{ $_.PrimarySmtpAddress -eq '{user_upn}' }}
            if ($isMember) {{
                Remove-DistributionGroupMember -Identity $groupName -Member '{user_upn}' -Confirm:$false -BypassSecurityGroupManagerCheck
                $success = $true
            }} else {{
                Write-Output "skipped : $($user_upn):$($groupName):{role} - Not a member"
                $skipped = $true
            }}
        }} elseif ('{role}' -eq 'Owner') {{
            $owners = $group.ManagedBy
            if ($owners -contains '{user_upn}') {{
                Set-DistributionGroup -Identity $groupName -ManagedBy @{{remove = '{user_upn}' }} -BypassSecurityGroupManagerCheck
                $success = $true
            }} else {{
                Write-Output "skipped : $($user_upn):$($groupName):{role} - Not an owner"
                $skipped = $true
            }}
        }}

        if ($success) {{
            Write-Output "success : $($user_upn):$($groupName):{role}"
        }} elseif (-not $skipped) {{
            Write-Output "failed : $($user_upn):$($groupName):{role} - Unknown failure"
        }}
    }} catch {{
        Write-Output "failed : $($user_upn):$($group_name):{role} - $($_.Exception.Message)"
    }}
    Disconnect-ExchangeOnline -Confirm:$false
    """
    return execute_powershell(log, ps_command)

def remove_exo_all_distributiongroup_permissions(log, user_upn, exo_access_token, azure_domain):
    log.info(f"Removing [{user_upn}] from all distribution groups (member and owner)")

    ps_command = f"""
    $ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false

    $UPN = '{user_upn}'
    $Success = @()
    $Failed = @()
    $Found = $false

    try {{
        $user = Get-User -Identity $UPN
        if (-not $user) {{
            Write-Output "User not found: $UPN"
            Disconnect-ExchangeOnline -Confirm:$false
            exit 1
        }}
        $dn = $user.DistinguishedName

        $groups = Get-DistributionGroup -ResultSize Unlimited

        foreach ($group in $groups) {{
            $groupName = $group.Identity
            $isMember = Get-DistributionGroupMember -Identity $groupName -ResultSize Unlimited | Where-Object {{ $_.PrimarySmtpAddress -eq $UPN }}
            $isOwner = $group.ManagedBy -contains $dn

            if ($isMember) {{
                try {{
                    Remove-DistributionGroupMember -Identity $groupName -Member $UPN -Confirm:$false -BypassSecurityGroupManagerCheck
                    $Success += "$($UPN):$($groupName):Member"
                    $Found = $true
                }} catch {{
                    $Failed += "$($UPN):$($groupName):Member"
                }}
            }}

            if ($isOwner) {{
                try {{
                    Set-DistributionGroup -Identity $groupName -ManagedBy @{{remove = $dn }} -BypassSecurityGroupManagerCheck
                    $Success += "$($UPN):$($groupName):Owner"
                    $Found = $true
                }} catch {{
                    $Failed += "$($UPN):$($groupName):Owner"
                }}
            }}
        }}
    }} catch {{
        Write-Output "PowerShell exception occurred: $_"
        Disconnect-ExchangeOnline -Confirm:$false
        exit 1
    }}

    Disconnect-ExchangeOnline -Confirm:$false

    if (-not $Found) {{
        Write-Output "NO_PERMISSIONS_FOUND"
    }}

    if ($Success.Count -gt 0) {{
        Write-Output "Removed from: $($Success -join ', ')"
    }}

    if ($Failed.Count -gt 0) {{
        Write-Output "Failed on: $($Failed -join ', ')"
    }}
    """

    success, output = execute_powershell(log, ps_command)
    if not success:
        return False, "PowerShell execution failed"

    found_removal = False
    for line in output.splitlines():
        if line == "NO_PERMISSIONS_FOUND":
            return True, f"No distribution group permissions found for [{user_upn}]"
        elif line.startswith("Removed from:"):
            for entry in line.replace("Removed from:", "").split(","):
                e = entry.strip()
                if e:
                    log.info(e)
                    data_to_log.setdefault("Success", "")
                    data_to_log["Success"] += ("," if data_to_log["Success"] else "") + e
                    found_removal = True
        elif line.startswith("Failed on:"):
            for e in line.replace("Failed on:", "").split(","):
                e = e.strip()
                if e:
                    log.warning(f"Remove Distribution Group Permissions - {e}")
                    data_to_log.setdefault("Failed", "")
                    data_to_log["Failed"] += ("," if data_to_log["Failed"] else "") + f"Remove Distribution Group Permissions - {e}"

    return True, "Removal complete" if found_removal else f"No distribution group permissions found for [{user_upn}]"

def main():
    try:
        try:
            operation = input.get_value("Operation_1749762984341")
            user_identifier = input.get_value("User_1749762990167")
            access = input.get_value("Access_1750035814390")
            group = input.get_value("DistributionGroup_1749762993185")
            ticket_number = input.get_value("TicketNumber_1749762981385")
            auth_code = input.get_value("AuthCode_1749762991858")
            provided_token = input.get_value("AccessToken_1749762994745")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        user_identifier = user_identifier.strip() if user_identifier else ""
        ticket_number = ticket_number.strip() if ticket_number else ""
        auth_code = auth_code.strip() if auth_code else ""
        provided_token = provided_token.strip() if provided_token else ""
        operation = operation.strip()
        access = access.strip() if access else ""
        group = group.strip() if group else ""

        log.info("Successfully fetched all input values")
        log.info(f"Raw inputs: operation=[{operation}], user_identifier=[{user_identifier}], access=[{access}], group=[{group}], ticket_number=[{ticket_number}]")

        if not user_identifier and operation != "Get Distribution Group":
            record_result(log, ResultLevel.WARNING, "User identifier is empty or invalid")
            return

        if operation not in (
            "Get Distribution Group",
            "Add Distribution Group Permissions",
            "Remove Distribution Group Permissions",
            "Remove All Distribution Group Permissions",
            "Add Distribution Group Permissions for Contact"
        ):
            record_result(log, ResultLevel.WARNING, f"Invalid operation: {operation}")
            return

        if operation in (
            "Add Distribution Group Permissions",
            "Remove Distribution Group Permissions"
        ) and (not access or not group):
            record_result(log, ResultLevel.WARNING, "Access and Distribution Group are required")
            return

        if operation == "Add Distribution Group Permissions for Contact" and not group:
            record_result(log, ResultLevel.WARNING, "Distribution Group is required")
            return

        if provided_token and "." in provided_token:
            access_token = provided_token
            log.info("Using provided access token")
        elif ticket_number:
            log.info(f"Authenticating using ticket [{ticket_number}]")
            company_identifier, _, _, _ = get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
            if not company_identifier:
                record_result(log, ResultLevel.WARNING, f"Failed to retrieve company identifier for ticket [{ticket_number}]")
                return

            log.info(f"Company identifier resolved as [{company_identifier}]")

            if company_identifier == "MIT":
                if not validate_mit_authentication(log, http_client, vault_name, auth_code):
                    return

            log.info("Retrieving secrets for Graph access token")
            client_id = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientID")
            client_secret = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientSecret")
            azure_domain = get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")

            if not all([client_id, client_secret, azure_domain]):
                record_result(log, ResultLevel.WARNING, "Failed to retrieve required secrets for Azure authentication")
                return

            tenant_id = get_tenant_id_from_domain(log, http_client, azure_domain)
            if not tenant_id:
                record_result(log, ResultLevel.WARNING, "Failed to resolve tenant ID from Azure domain")
                return

            log.info("Requesting Microsoft Graph access token")
            access_token = get_access_token(log, http_client, tenant_id, client_id, client_secret)
            if not access_token or "." not in access_token:
                record_result(log, ResultLevel.WARNING, "Access token is malformed (missing dots)")
                return
        else:
            record_result(log, ResultLevel.WARNING, "Either Access Token or Ticket Number is required")
            return

        log.info("Retrieving Exchange Online credentials")
        exo_client_id = get_secret_value(log, http_client, vault_name, f"{company_identifier}-ExchangeApp-ClientID")
        exo_client_secret = get_secret_value(log, http_client, vault_name, f"{company_identifier}-ExchangeApp-ClientSecret")
        azure_domain = azure_domain if 'azure_domain' in locals() else get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")

        if not all([exo_client_id, exo_client_secret, azure_domain]):
            record_result(log, ResultLevel.WARNING, "Failed to retrieve required secrets for Exchange Online authentication")
            return

        log.info("Requesting Exchange Online access token")
        exo_access_token = get_access_token(
            log, http_client, tenant_id, exo_client_id, exo_client_secret,
            scope="https://outlook.office365.com/.default"
        )
        if not exo_access_token or "." not in exo_access_token:
            record_result(log, ResultLevel.WARNING, "Exchange Online access token is malformed (missing dots)")
            return

        if operation == "Add Distribution Group Permissions for Contact":
            user_email = user_identifier
        else:
            aad_user_result = get_aad_user_data(log, http_client, msgraph_base_url, user_identifier, access_token)

            if isinstance(aad_user_result, list):
                details = "\n".join(
                    [f"- {u.get('displayName')} | {u.get('userPrincipalName')} | {u.get('id')}" for u in aad_user_result]
                )
                record_result(
                    log,
                    ResultLevel.WARNING,
                    f"Multiple users found for [{user_identifier}]\n{details}",
                )
                return

            _, user_email, _ = aad_user_result

        log.info(f"Executing operation [{operation}] for user [{user_email}]")

        log_entries = []
        successes = []
        failures = []

        if operation == "Get Distribution Group":
            group_list = [g.strip() for g in group.split(",") if g.strip()]
            if not group_list:
                record_result(log, ResultLevel.WARNING, "Distribution Group is required")
                return

            retrieved = []
            for group_name in group_list:
                exists = get_exo_distributiongroup(log, group_name, exo_access_token, azure_domain)
                if not exists:
                    record_result(log, ResultLevel.WARNING, f"Group [{group_name}] could not be resolved in Exchange Online")
                else:
                    retrieved.append(group_name)
                    record_result(log, ResultLevel.SUCCESS, f"Retrieved distribution group [{group_name}]")
            if retrieved:
                data_to_log["Groups"] = retrieved
            return

        elif operation == "Add Distribution Group Permissions":
            group_list = [g.strip() for g in group.split(",") if g.strip()]
            for group_name in group_list:
                if not get_exo_distributiongroup(log, group_name, exo_access_token, azure_domain):
                    failures.append(f"Group [{group_name}] could not be resolved")
                    continue

                success, entry = add_exo_distributiongroup_permissions(log, group_name, user_email, access, exo_access_token, azure_domain)
                log_entries.append(entry)
                if success and entry.startswith("success :"):
                    parts = entry.split("success : ")[1].split(":")
                    if len(parts) >= 3:
                        summary = f"{parts[0]} added as {parts[2]} to {parts[1]}"
                    else:
                        summary = entry.split("success : ")[1]
                    successes.append(summary)
                elif entry.startswith("failed :"):
                    parts = entry.split(" - ")[0].replace("failed :", "").strip().split(":")
                    user = parts[0] if len(parts) > 0 else user_email
                    grp = parts[1] if len(parts) > 1 else group_name
                    failures.append(f"Failed to add {user} to {grp}")

        elif operation == "Remove Distribution Group Permissions":
            group_list = [g.strip() for g in group.split(",") if g.strip()]
            for group_name in group_list:
                if not get_exo_distributiongroup(log, group_name, exo_access_token, azure_domain):
                    failures.append(f"Group [{group_name}] could not be resolved")
                    continue

                success, entry = remove_exo_distributiongroup_permissions(log, group_name, user_email, access, exo_access_token, azure_domain)
                log_entries.append(entry)
                if success and entry.startswith("success :"):
                    parts = entry.split("success : ")[1].split(":")
                    if len(parts) >= 3:
                        summary = f"{parts[0]} removed as {parts[2]} from {parts[1]}"
                    else:
                        summary = entry.split("success : ")[1]
                    successes.append(summary)
                elif entry.startswith("failed :"):
                    parts = entry.split(" - ")[0].replace("failed :", "").strip().split(":")
                    user = parts[0] if len(parts) > 0 else user_email
                    grp = parts[1] if len(parts) > 1 else group_name
                    failures.append(f"Failed to remove {user} from {grp}")

        elif operation == "Remove All Distribution Group Permissions":
            success, message = remove_exo_all_distributiongroup_permissions(log, user_email, exo_access_token, azure_domain)

            if not success:
                record_result(log, ResultLevel.WARNING, message)
            else:
                if "Success" in data_to_log and data_to_log["Success"]:
                    for entry in data_to_log["Success"].split(","):
                        record_result(log, ResultLevel.SUCCESS, f"{entry.strip()} removed")

                if "Failed" in data_to_log and data_to_log["Failed"]:
                    for entry in data_to_log["Failed"].split(","):
                        record_result(log, ResultLevel.WARNING, entry.strip())

                if "Success" not in data_to_log and "Failed" not in data_to_log:
                    record_result(log, ResultLevel.SUCCESS, message)

        elif operation == "Add Distribution Group Permissions for Contact":
            group_list = [g.strip() for g in group.split(",") if g.strip()]
            results = add_exo_distributiongroup_permissions_contact(log, group_list, user_email, exo_access_token, azure_domain)
            for result in results:
                if result.startswith("success :"):
                    record_result(log, ResultLevel.SUCCESS, result.replace("success :", "").strip())
                else:
                    record_result(log, ResultLevel.WARNING, result.replace("failed :", "").strip())

        if successes:
            data_to_log["Success"] = ", ".join(successes)

        if failures:
            data_to_log["Failed"] = ", ".join(failures)

        result_level = ResultLevel.SUCCESS if not failures else ResultLevel.WARNING
        for log_entry in log_entries:
            if not log_entry.startswith("skipped :"):
                cleaned_entry = log_entry.replace("success : ", "", 1) if log_entry.startswith("success :") else log_entry
                record_result(log, result_level, cleaned_entry)

        log.result_data(data_to_log)

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()