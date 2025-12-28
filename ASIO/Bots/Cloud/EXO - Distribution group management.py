import sys
import random
import re
import subprocess
import os
import time
import urllib.parse
import requests
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

cwpsa_base_url = "https://au.myconnectwise.net/v4_6_release/apis/3.0"
msgraph_base_url = "https://graph.microsoft.com/v1.0"
msgraph_base_url_beta = "https://graph.microsoft.com/beta"
vault_name = "PLACEHOLDER-akv1"

data_to_log = {}
bot_name = "EXO - Distribution group management"
log.info("Static variables set")

def record_result(log, level, message):
    log.result_message(level, f"[{bot_name}]: {message}")

    if level == ResultLevel.WARNING:
        data_to_log["status_result"] = "Fail"
    elif level == ResultLevel.SUCCESS:
        if "status_result" not in data_to_log or data_to_log["status_result"] != "Fail":
            data_to_log["status_result"] = "Success"

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

def execute_powershell(log, command, shell="pwsh", debug_mode=False, timeout=None, ignore_stderr_warnings=False, log_command=True, log_output=True):
    try:
        if shell not in ("pwsh", "powershell"):
            raise ValueError("Invalid shell specified. Use 'pwsh' or 'powershell'")
        if log_command:
            sanitized_cmd = command
            sanitized_cmd = re.sub(r"('|\")(eyJ[a-zA-Z0-9_-]{5,}?\.[a-zA-Z0-9_-]{5,}?\.([a-zA-Z0-9_-]{5,})?)\1", r"\1***TOKEN-MASKED***\1", sanitized_cmd)
            sanitized_cmd = re.sub(r"('|\")([a-zA-Z0-9]{8,}-(clientid|clientsecret|password))\1", r"\1***SECRET-MASKED***\1", sanitized_cmd, flags=re.IGNORECASE)
            log.info(f"Executing {shell} command: {sanitized_cmd[:100]}{'...' if len(sanitized_cmd) > 100 else ''}")
        if debug_mode:
            log.info(f"Debug mode enabled for PowerShell execution")
        result = subprocess.run([shell, "-Command", command], capture_output=True, text=True, timeout=timeout)
        stdout = result.stdout.strip()
        stderr = result.stderr.strip()
        if log_output and stdout and debug_mode:
            log.info(f"{shell} stdout: {stdout[:500]}{'...' if len(stdout) > 500 else ''}")
        if stderr and debug_mode:
            log.info(f"{shell} stderr: {stderr}")
        success = result.returncode == 0
        if success and stderr and ignore_stderr_warnings:
            error_patterns = ["error:", "exception:", "fatal:", "failed:"]
            if any(pattern in stderr.lower() for pattern in error_patterns):
                success = False
                log.warning(f"{shell} command returned warnings: {stderr}")
            else:
                log.info(f"{shell} command executed with warnings: {stderr}")
        if success:
            if not stderr:
                log.info(f"{shell} command executed successfully")
            return True, stdout
        else:
            error_message = stderr or stdout
            log.error(f"{shell} execution failed (returncode={result.returncode}): {error_message}")
            return False, error_message
    except subprocess.TimeoutExpired:
        log.error(f"{shell} command timed out after {timeout} seconds")
        return False, f"Command timed out after {timeout} seconds"
    except Exception as e:
        log.exception(e, f"Exception occurred during {shell} execution")
        return False, str(e)

def get_secret_value(log, http_client, vault_name, secret_name):
    log.info(f"Fetching secret [{secret_name}] from Key Vault [{vault_name}]")

    secret_url = f"https://{vault_name}.vault.azure.net/secrets/{secret_name}?api-version=7.3"
    response = execute_api_call(log, http_client, "get", secret_url, integration_name="custom_wf_oauth2_client_creds")

    if response:
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

        if response:
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

def get_access_token(log, http_client, tenant_id, client_id, client_secret, scope="https://graph.microsoft.com/.default", log_prefix="Token"):
    log.info(f"[{log_prefix}] Requesting access token for scope [{scope}]")
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    payload = urllib.parse.urlencode({
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": scope
    })
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    response = execute_api_call(log, http_client, "post", token_url, data=payload, retries=3, headers=headers)
    if response:
        token_data = response.json()
        access_token = str(token_data.get("access_token", "")).strip()
        log.info(f"[{log_prefix}] Access token length: {len(access_token)}")
        log.info(f"[{log_prefix}] Access token preview: {access_token[:30]}...")
        if not isinstance(access_token, str) or "." not in access_token:
            log.error(f"[{log_prefix}] Access token is invalid or malformed")
            return ""
        log.info(f"[{log_prefix}] Successfully retrieved access token")
        return access_token
    log.error(f"[{log_prefix}] Failed to retrieve access token Status code: {response.status_code if response else 'N/A'}")
    return ""

def get_graph_token(log, http_client, vault_name, company_identifier):
    client_id = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientID")
    client_secret = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientSecret")
    azure_domain = get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")
    if not all([client_id, client_secret, azure_domain]):
        log.error("Failed to retrieve required secrets for MS Graph")
        return "", ""
    tenant_id = get_tenant_id_from_domain(log, http_client, azure_domain)
    if not tenant_id:
        log.error("Failed to resolve tenant ID for MS Graph")
        return "", ""
    token = get_access_token(log, http_client, tenant_id, client_id, client_secret, scope="https://graph.microsoft.com/.default", log_prefix="Graph")
    if not isinstance(token, str) or "." not in token:
        log.error("MS Graph access token is malformed (missing dots)")
        return "", ""
    return tenant_id, token

def get_exo_token(log, http_client, vault_name, company_identifier):
    client_id = get_secret_value(log, http_client, vault_name, f"{company_identifier}-ExchangeApp-ClientID")
    client_secret = get_secret_value(log, http_client, vault_name, f"{company_identifier}-ExchangeApp-ClientSecret")
    azure_domain = get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")
    if not all([client_id, client_secret, azure_domain]):
        log.error("Failed to retrieve required secrets for EXO")
        return "", ""
    tenant_id = get_tenant_id_from_domain(log, http_client, azure_domain)
    if not tenant_id:
        log.error("Failed to resolve tenant ID for EXO")
        return "", ""
    token = get_access_token(log, http_client, tenant_id, client_id, client_secret, scope="https://outlook.office365.com/.default", log_prefix="EXO")
    if not isinstance(token, str) or "." not in token:
        log.error("EXO access token is malformed (missing dots)")
        return "", ""
    return tenant_id, token

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

        company_types = []
        if company_response and company_response.status_code == 200:
            company_data = company_response.json()
            types = company_data.get("types", [])
            company_types = [t.get("name", "") for t in types if "name" in t]
            log.info(f"Company types for ID [{company_id}]: {company_types}")
        else:
            log.warning(f"Unable to retrieve company types for ID [{company_id}]")

        return company_identifier, company_name, company_id, company_types

    elif ticket_response:
        log.error(
            f"Failed to retrieve ticket [{ticket_number}] "
            f"Status: {ticket_response.status_code}, Body: {ticket_response.text}"
        )
    return "", "", 0, []

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
    if re.fullmatch(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}", user_identifier):
        endpoint = f"{msgraph_base_url}/users/{user_identifier}"
        response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
        if response:
            user = response.json()
            return user.get("id", ""), user.get("userPrincipalName", ""), user.get("onPremisesSamAccountName", ""), user.get("onPremisesSyncEnabled", False)
        return "", "", "", False
    filters = [
        f"startswith(displayName,'{user_identifier}')",
        f"startswith(userPrincipalName,'{user_identifier}')",
        f"startswith(mail,'{user_identifier}')"
    ]
    filter_query = " or ".join(filters)
    endpoint = f"{msgraph_base_url}/users?$filter={urllib.parse.quote(filter_query)}"
    response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
    if response:
        users = response.json().get("value", [])
        if len(users) > 1:
            log.error(f"Multiple users found for [{user_identifier}]")
            return users
        if users:
            user = users[0]
            return user.get("id", ""), user.get("userPrincipalName", ""), user.get("onPremisesSamAccountName", ""), user.get("onPremisesSyncEnabled", False)
    log.error(f"Failed to resolve user ID and email for [{user_identifier}]")
    return "", "", "", False

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
                Write-Output "skipped : {user_upn}:{group_name}:{permission_type}:already a [Member]"
            }} else {{
                Add-DistributionGroupMember -Identity $groupName -Member '{user_upn}' -BypassSecurityGroupManagerCheck
                Write-Output "success : {user_upn}:{group_name}:{permission_type}"
            }}
        }} elseif ('{permission_type}' -eq 'Owner') {{
            if ($existingOwners -contains '{user_upn}') {{
                Write-Output "skipped : {user_upn}:{group_name}:{permission_type}:already an [Owner]"
            }} else {{
                Set-DistributionGroup -Identity $groupName -ManagedBy @{{add = '{user_upn}' }} -BypassSecurityGroupManagerCheck
                Write-Output "success : {user_upn}:{group_name}:{permission_type}"
            }}
        }} else {{
            Write-Output "failed : {user_upn}:{group_name}:{permission_type}:unsupported permission type"
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
            results.append(f"failed : {contact_email}:{group_name}:group not found")
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
            Write-Output "failed : {contact_email}:{group_name}:$($_.Exception.Message)"
        }}
        Disconnect-ExchangeOnline -Confirm:$false
        """
        success, output = execute_powershell(log, ps_command)
        results.append(output.strip() if output else f"failed : {contact_email}:{group_name}:no response")
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
            $isMember = Get-DistributionGroupMember -Identity {group_name} -ResultSize Unlimited | Where-Object {{ $_.PrimarySmtpAddress -eq '{user_upn}' }}
            if ($isMember) {{
                Remove-DistributionGroupMember -Identity {group_name} -Member '{user_upn}' -Confirm:$false -BypassSecurityGroupManagerCheck
                $success = $true
            }} else {{
                Write-Output "skipped : {user_upn}:{group_name}:{role}:not a member"
                $skipped = $true
            }}
        }} elseif ('{role}' -eq 'Owner') {{
            $owners = $group.ManagedBy
            if ($owners -contains '{user_upn}') {{
                Set-DistributionGroup -Identity {group_name} -ManagedBy @{{remove = '{user_upn}' }} -BypassSecurityGroupManagerCheck
                $success = $true
            }} else {{
                Write-Output "skipped : {user_upn}:{group_name}:{role}:not an owner"
                $skipped = $true
            }}
        }}

        if ($success) {{
            Write-Output "success : {user_upn}:{group_name}:{role}"
        }} elseif (-not $skipped) {{
            Write-Output "failed : {user_upn}:{group_name}:{role}:unknown failure"
        }}
    }} catch {{
        Write-Output "failed : {user_upn}:{group_name}:{role}:$($_.Exception.Message)"
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
            ticket_number = input.get_value("TicketNumber_1757550681738")
            operation = input.get_value("Operation_1757550683556")
            user_identifier = input.get_value("User_1757550689276")
            access = input.get_value("Access_1757550833787")
            group = input.get_value("DistributionGroups_1757550870349")
            auth_code = input.get_value("AuthenticationCode_1757550931344")
            graph_token = input.get_value("GraphToken_1757550929082")
            exo_token = input.get_value("ExchangeOnlineToken_1757550932692")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        operation = operation.strip() if operation else ""
        user_identifier = user_identifier.strip() if user_identifier else ""
        access = access.strip() if access else ""
        group = group.strip() if group else ""
        auth_code = auth_code.strip() if auth_code else ""
        graph_token = graph_token.strip() if graph_token else ""
        exo_token = exo_token.strip() if exo_token else ""

        log.info(f"Ticket Number = [{ticket_number}]")
        log.info(f"Requested operation = [{operation}]")
        log.info(f"Received input user = [{user_identifier}]")

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number is required but missing")
            return
        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation value is missing or invalid")
            return
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
            record_result(log, ResultLevel.WARNING, "Please provide an access level and distribution group to add or remove permissions")
            return

        if operation == "Add Distribution Group Permissions for Contact" and not group:
            record_result(log, ResultLevel.WARNING, "Please provide a distribution group to add permissions for a contact")
            return

        log.info(f"Retrieving company data for ticket [{ticket_number}]")
        company_identifier, company_name, company_id, company_types = get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
        if not company_identifier:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve company identifier from ticket [{ticket_number}]")
            return

        if company_identifier == "MIT":
            if not validate_mit_authentication(log, http_client, vault_name, auth_code):
                return

        if graph_token:
            log.info("Using provided MS Graph token")
            graph_access_token = graph_token
            graph_tenant_id = ""
        else:
            graph_tenant_id, graph_access_token = get_graph_token(log, http_client, vault_name, company_identifier)
            if not graph_access_token:
                record_result(log, ResultLevel.WARNING, "Failed to obtain MS Graph access token")
                return

        if exo_token:
            log.info("Using provided Exchange Online token")
            exo_access_token = exo_token
            exo_tenant_id = ""
        else:
            exo_tenant_id, exo_access_token = get_exo_token(log, http_client, vault_name, company_identifier)
            if not exo_access_token:
                record_result(log, ResultLevel.WARNING, "Failed to obtain Exchange Online access token")
                return

        azure_domain = get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")
        if not azure_domain:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve Azure domain for [{company_identifier}]")
            return

        if operation == "Add Distribution Group Permissions for Contact":
            user_email = user_identifier
        else:
            aad_user_result = get_aad_user_data(log, http_client, msgraph_base_url, user_identifier, graph_access_token)

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

            user_id, user_email, user_sam, user_onpremisessyncenabled = aad_user_result

        log.info(f"Executing operation [{operation}] for user [{user_email}]")

        log_entries = []
        successes = []
        failures = []
        skipped = []

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
                data_to_log["groups"] = retrieved
            return

        elif operation == "Add Distribution Group Permissions":
            group_list = [g.strip() for g in group.split(",") if g.strip()]
            for group_name in group_list:
                if not get_exo_distributiongroup(log, group_name, exo_access_token, azure_domain):
                    failures.append(f"Group [{group_name}] could not be resolved")
                    continue

                success, entry = add_exo_distributiongroup_permissions(log, group_name, user_email, access, exo_access_token, azure_domain)
                log_entries.append(entry)
                log.info(f"Entry: {entry}")
                if success and entry.startswith("success :"):
                    parts = entry.split("success : ")[1].split(":")
                    if len(parts) >= 3:
                        summary = f"[{parts[0]}] added to [{parts[1]}] as [{parts[2]}]"
                    else:
                        summary = entry.split("success : ")[1]
                    successes.append(summary)
                elif entry.startswith("failed :"):
                    parts = entry.split(" - ")[0].replace("failed :", "").strip().split(":")
                    user = parts[0] if len(parts) > 0 else user_email
                    grp = parts[1] if len(parts) > 1 else group_name
                    failures.append(f"Failed to add [{user}] to [{grp}]")
                elif entry.startswith("skipped :"):
                    parts = entry.split(" - ")[0].replace("skipped :", "").strip().split(":")
                    user = parts[0] if len(parts) > 0 else user_email
                    grp = parts[1] if len(parts) > 1 else group_name
                    if len(parts) > 2:
                        reason = f", {parts[3]}"
                    else:
                        reason = ""
                    skipped.append(f"Skipped adding [{user}] to [{grp}]{reason}")

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
                        summary = f"[{parts[0]}] removed as [{parts[2]}] from [{parts[1]}]"
                    else:
                        summary = entry.split("success : ")[1]
                    successes.append(summary)
                elif entry.startswith("failed :"):
                    parts = entry.split(" - ")[0].replace("failed :", "").strip().split(":")
                    user = parts[0] if len(parts) > 0 else user_email
                    grp = parts[1] if len(parts) > 1 else group_name
                    failures.append(f"Failed to remove [{user}] from [{grp}]")

        elif operation == "Remove All Distribution Group Permissions":
            success, message = remove_exo_all_distributiongroup_permissions(log, user_email, exo_access_token, azure_domain)

            if not success:
                record_result(log, ResultLevel.WARNING, message)
            else:
                if "Success" in data_to_log and data_to_log["Success"]:
                    for entry in data_to_log["Success"].split(","):
                        parts = entry.split(":")
                        summary = f"[{parts[0]}] removed as [{parts[2]}] from [{parts[1]}]"

                if "Failed" in data_to_log and data_to_log["Failed"]:
                    for entry in data_to_log["Failed"].split(","):
                        parts = entry.split(":")
                        summary = f"Failed to remove [{parts[0]}] from [{parts[1]}] as [{parts[2]}]"
                        failures.append(summary)

                if "Success" not in data_to_log and "Failed" not in data_to_log:
                    record_result(log, ResultLevel.SUCCESS, message)

        elif operation == "Add Distribution Group Permissions for Contact":
            group_list = [g.strip() for g in group.split(",") if g.strip()]
            results = add_exo_distributiongroup_permissions_contact(log, group_list, user_email, exo_access_token, azure_domain)
            for result in results:
                if result.startswith("success :"):
                    parts = result.split(":")
                    summary = f"Contact[{parts[0]}] added to [{parts[1]}]"
                    successes.append(summary)
                else:
                    parts = result.split(":")
                    summary = f"Could not add [{parts[0]}] to [{parts[1]}]"
                    failures.append(summary)

        if successes:
            data_to_log["Success"] = ", ".join(successes)
        if failures:
            data_to_log["Failed"] = ", ".join(failures)
        if skipped:
            data_to_log["Skipped"] = ", ".join(skipped)

        for success in successes:
            record_result(log, ResultLevel.SUCCESS, success)
        for failure in failures:
            record_result(log, ResultLevel.WARNING, failure)
        for skipped in skipped:
            record_result(log, ResultLevel.SUCCESS, skipped)

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
