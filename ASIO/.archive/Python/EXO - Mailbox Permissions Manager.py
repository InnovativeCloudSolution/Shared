import sys
import random
import re
import subprocess
import os
import time
import urllib.parse
import requests
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
bot_name = "[EXO Mailbox Permissions Manager]"
log.info("Static variables set")

def record_result(log, level, message):
    log.result_message(level, f"{bot_name}: {message}")

    if level == ResultLevel.WARNING:
        data_to_log["Result"] = "Fail"
    elif level == ResultLevel.SUCCESS:
        if "Result" not in data_to_log or data_to_log["Result"] != "Fail":
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

    if response and response.status_code == 200:
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

    token = get_access_token(
        log, http_client, tenant_id, client_id, client_secret,
        scope="https://graph.microsoft.com/.default", log_prefix="Graph"
    )
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

    token = get_access_token(
        log, http_client, tenant_id, client_id, client_secret,
        scope="https://outlook.office365.com/.default", log_prefix="EXO"
    )
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
    else:
        log.error(f"No response received when retrieving ticket [{ticket_number}]")

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
                user.get("onPremisesSyncEnabled", False)
            )
        return "", "", "", False

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
                user.get("onPremisesSyncEnabled", False)
            )
    log.error(f"Failed to resolve user ID and email for [{user_identifier}]")
    return "", "", "", False

def add_exo_sendonbehalfof_permissions(log, azure_domain, user_email, user_id, mailbox, exo_access_token):
    log.info(f"Adding SendOnBehalfOf permission for [{user_email}] on mailbox [{mailbox}]")
    ps_command = f"""$ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false
    $resolved = try {{ (Get-Mailbox -Identity "{mailbox}" -ErrorAction Stop).UserPrincipalName }} catch {{ "{mailbox}" }}
    try {{
        Set-Mailbox -Identity "{mailbox}" -GrantSendOnBehalfTo "{user_id}" -Confirm:$false -ErrorAction Stop
        Write-Output "success : {user_email} => $resolved:SendOnBehalfOf"
    }} catch {{
        Write-Output "failed : {user_email} => $resolved:SendOnBehalfOf"
    }}
    Disconnect-ExchangeOnline -Confirm:$false"""

    success, output = execute_powershell(log, ps_command, ignore_stderr_warnings=True)
    
    if success:
        msg = f"{user_email} given SendOnBehalfOf to {mailbox}"
        log.info(msg)
        return success, [(ResultLevel.SUCCESS, msg)]
    else:
        if user_email and mailbox:
            msg = f"Failed to give SendOnBehalfOf to {user_email} on {mailbox}"
            log.warning(msg)
            return success, [(ResultLevel.WARNING, msg)]
    return success, []

def add_exo_sendas_permissions(log, azure_domain, user_email, mailbox, exo_access_token):
    log.info(f"Adding SendAs permission for [{user_email}] on mailbox [{mailbox}]")
    ps_command = f"""$ErrorActionPreference = 'Continue'
        $WarningPreference = 'Continue'
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false
        
        try {{
            $result = Add-RecipientPermission -Identity '{mailbox}' -Trustee '{user_email}' -AccessRights SendAs -Confirm:$false -ErrorAction Stop -WarningAction Continue -WarningVariable warnings 3>&1
            if ($warnings -and $warnings -like "*appropriate access control entry is already present*") {{
                Write-Output 'SUCCESS_ALREADY_EXISTS'
            }} else {{
                Write-Output 'SUCCESS'
            }}
        }} catch {{
            $errorMsg = $_.Exception.Message
            if ($errorMsg -like "*already has*" -or $errorMsg -like "*already exists*") {{
                Write-Output 'SUCCESS_ALREADY_EXISTS'
            }} else {{
                Write-Output "FAILED:$errorMsg"
            }}
        }}
        Disconnect-ExchangeOnline -Confirm:$false"""
    success, output = execute_powershell(log, ps_command, ignore_stderr_warnings=True)
    
    if 'SUCCESS_ALREADY_EXISTS' in (output or ''):
        msg = f"{user_email} already has SendAs on {mailbox}"
        log.info(msg)
        return True, [(ResultLevel.SUCCESS, msg)]
    elif 'SUCCESS' in (output or ''):
        msg = f"{user_email} given SendAs to {mailbox}"
        log.info(msg)
        return True, [(ResultLevel.SUCCESS, msg)]
    elif 'FAILED:' in (output or ''):
        error_msg = output.split('FAILED:')[1].split('\n')[0].strip() if output else "Unknown error"
        msg = f"Failed to give SendAs to {user_email} on {mailbox}: {error_msg}"
        log.warning(msg)
        return False, [(ResultLevel.WARNING, msg)]
    else:
        msg = f"Failed to give SendAs to {user_email} on {mailbox}"
        log.warning(msg)
        return False, [(ResultLevel.WARNING, msg)]

def add_exo_mailbox_permissions(log, azure_domain, user_email, mailbox, access, exo_access_token, auto_mapping):
    log.info(f"Adding mailbox permissions for user [{user_email}] on mailbox [{mailbox}] with access level [{access}], AutoMapping = [{auto_mapping}]")
    auto_mapping_flag = "$true" if auto_mapping else "$false"
    ps_command = f"""$ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false
    $resolved = try {{ (Get-Mailbox -Identity "{mailbox}" -ErrorAction Stop).UserPrincipalName }} catch {{ "{mailbox}" }}
    try {{
        Add-MailboxPermission -Identity "{mailbox}" -User "{user_email}" -AccessRights {access} -AutoMapping {auto_mapping_flag} -Confirm:$false -ErrorAction Stop
        Write-Output "success : {user_email} => $resolved:{access}"
    }} catch {{
        Write-Output "failed : {user_email} => $resolved:{access}"
    }}
    Disconnect-ExchangeOnline -Confirm:$false"""

    success, output = execute_powershell(log, ps_command)
    if success:
        msg = f"{user_email} given {access} to {mailbox}"
        log.info(msg)
        return success, [(ResultLevel.SUCCESS, msg)]
    else:
        if user_email and mailbox and access:
            msg = f"Failed to give {access} to {user_email} on {mailbox}"
            log.warning(msg)
            return success, [(ResultLevel.WARNING, msg)]
    return success, []

def remove_exo_mailbox_permissions(log, azure_domain, user_email, mailbox, access, exo_access_token):
    log.info(f"Removing [{access}] permission for user [{user_email}] on mailbox [{mailbox}]")
    if access in ["SendAs", "SendOnBehalfOf"]:
        if access == "SendAs":
            cmd = f'Remove-RecipientPermission -Identity "{mailbox}" -Trustee "{user_email}" -AccessRights SendAs -Confirm:$false -ErrorAction Stop -SkipDomainValidationForMailContact -SkipDomainValidationForMailUser -SkipDomainValidationForSharedMailbox'
        else:
            cmd = f'Set-Mailbox -Identity "{mailbox}" -GrantSendOnBehalfTo @{{Remove="{user_email}"}} -ErrorAction Stop'
    else:
        cmd = f'Remove-MailboxPermission -Identity "{mailbox}" -User "{user_email}" -AccessRights {access} -Confirm:$false -ErrorAction Stop'
    
    ps_command = f"""$ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false
    $hasPerm = $false
    try {{
        if ("{access}" -eq "SendAs") {{
            $perm = Get-RecipientPermission -Identity "{mailbox}" | Where-Object {{ $_.Trustee -eq "{user_email}" -and $_.AccessRights -contains "SendAs" }}
            if ($perm) {{
                Write-Output "DEBUG: {user_email} has access to {mailbox}:{access}"
                $hasPerm = $true
            }}
        }} elseif ("{access}" -eq "SendOnBehalfOf") {{
            $mbx = Get-Mailbox -Identity "{mailbox}"
            if ($mbx.GrantSendOnBehalfTo -contains "{user_email}") {{
                Write-Output "DEBUG: {user_email} has access to {mailbox}:{access}"
                $hasPerm = $true
            }}
        }} else {{
            $perm = Get-MailboxPermission -Identity "{mailbox}" -User "{user_email}" -ErrorAction SilentlyContinue | Where-Object {{ $_.IsInherited -eq $false -and $_.AccessRights -contains "{access}" }}
            if ($perm) {{
                Write-Output "DEBUG: {user_email} has access to {mailbox}:{access}"
                $hasPerm = $true
            }}
        }}
        if ($hasPerm) {{
            try {{
                {cmd}
                Write-Output "success : {user_email} <= {mailbox}:{access}"
            }} catch {{
                Write-Output "failed : {user_email} <= {mailbox}:{access}"
            }}
        }}
    }} catch {{
        Write-Output "Error checking/removing {access} for {user_email} on {mailbox}: $_"
    }}
    Disconnect-ExchangeOnline -Confirm:$false"""

    success, output = execute_powershell(log, ps_command)
    results = []
    if not success:
        return False, [(ResultLevel.WARNING, f"PowerShell execution failed for {user_email} <= {mailbox}:{access}")]

    found = False
    msg = None
    for line in output.splitlines():
        if line.startswith("DEBUG:"):
            log.info(line)
        elif line.startswith("success :"):
            found = True
            entry = line.replace("success : ", "").strip()
            msg = f"{user_email} had {access} removed from {mailbox}"
            log.info(msg)
            results.append((ResultLevel.SUCCESS, msg))
        elif line.startswith("failed :"):
            found = True
            entry = line.replace("failed : ", "").strip()
            msg = f"Failed to remove {access} from {user_email} on {mailbox}"
            log.warning(msg)
            results.append((ResultLevel.WARNING, msg))

    if not found:
        return True, [(ResultLevel.SUCCESS, f"No {access} permissions found to remove for {user_email} on {mailbox}")]

    return True, results

def remove_exo_all_mailbox_permissions(log, azure_domain, user_email, exo_access_token):
    log.info(f"Removing FullAccess permissions for user [{user_email}]")
    ps_command = f'''$ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false
    $UPN = "{user_email}"
    $Success = @()
    $Failed = @()
    $Found = $false
    try {{
        $mailboxes = Get-Mailbox -ResultSize Unlimited
        foreach ($mailbox in $mailboxes) {{
            $permissions = Get-MailboxPermission -Identity $mailbox.UserPrincipalName | Where-Object {{ $_.User -eq $UPN -and $_.IsInherited -eq $false -and $_.AccessRights.Count -gt 0 }}
            foreach ($perm in $permissions) {{
                $Found = $true
                Write-Output "DEBUG: Has access to: $($mailbox.UserPrincipalName) - $($perm.AccessRights -join ',')"
                try {{
                    Remove-MailboxPermission -Identity $mailbox.UserPrincipalName -User $perm.User -AccessRights $perm.AccessRights -Confirm:$false -ErrorAction Stop
                    $Success += "$($mailbox.UserPrincipalName):$($perm.AccessRights -join ',')"
                }} catch {{
                    $Failed += "$($mailbox.UserPrincipalName):$($perm.AccessRights -join ',')"
                }}
            }}
        }}
    }} catch {{
        Write-Output "Error removing mailbox access permissions for $($UPN): $_"
    }}
    Disconnect-ExchangeOnline -Confirm:$false
    if (-not $Found) {{ Write-Output "NO_PERMISSIONS_FOUND" }}
    if ($Success.Count -gt 0) {{ Write-Output "Removed from: $($Success -join ', ')" }}
    if ($Failed.Count -gt 0) {{ Write-Output "Failed on: $($Failed -join ', ')" }}'''

    success, output = execute_powershell(log, ps_command)
    if not success:
        return False, "PowerShell execution failed"

    for line in output.splitlines():
        if line.startswith("DEBUG: Has access to:"):
            log.info(line)
        elif line == "NO_PERMISSIONS_FOUND":
            return True, f"No FullAccess permissions found to remove for {user_email}"
        elif line.startswith("Removed from:"):
            for entry in line.replace("Removed from: ", "").split(","):
                e = entry.strip()
                if e:
                    mailbox_name = e.split(":")[0] if ":" in e else e
                    msg = f"{user_email} had FullAccess removed from {mailbox_name}"
                    log.info(msg)
                    record_result(log, ResultLevel.SUCCESS, msg)
                    data_to_log.setdefault("Success", "")
                    data_to_log["Success"] += ("\n" if data_to_log["Success"] else "") + msg
        elif line.startswith("Failed on:"):
            content_after_failed = line.replace("Failed on: ", "").strip()
            if content_after_failed:
                entries = [e.strip() for e in content_after_failed.split(",") if e.strip()]
                if entries:
                    for e in entries:
                        mailbox_name = e.split(":")[0] if ":" in e else e
                        msg = f"Failed to remove FullAccess from {user_email} on {mailbox_name}"
                        log.warning(msg)
                        record_result(log, ResultLevel.WARNING, msg)
                    data_to_log.setdefault("Failed", "")
                    data_to_log["Failed"] += ("\n" if data_to_log["Failed"] else "") + msg

    return True, f"FullAccess permissions removal complete for {user_email}"

def remove_exo_all_sendonbehalfof_permissions(log, azure_domain, user_email, exo_access_token):
    log.info(f"Removing SendOnBehalfOf permissions for user [{user_email}]")
    ps_command = f'''$ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false
    $UPN = "{user_email}"
    $Success = @()
    $Failed = @()
    $Found = $false
    try {{
        $userGuid = (Get-Mailbox -Identity $UPN -ErrorAction Stop).ExternalDirectoryObjectId
        $mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {{ $_.GrantSendOnBehalfTo -contains $userGuid }}
        foreach ($mailbox in $mailboxes) {{
            $Found = $true
            Write-Output "DEBUG: Has access to: $($mailbox.UserPrincipalName)"
            try {{
                Set-Mailbox -Identity $mailbox.UserPrincipalName -GrantSendOnBehalfTo @{{Remove=$UPN}} -Confirm:$false -ErrorAction Stop
                $Success += "$($mailbox.UserPrincipalName):SendOnBehalfOf"
            }} catch {{
                $Failed += "$($mailbox.UserPrincipalName):SendOnBehalfOf"
            }}
        }}
    }} catch {{
        Write-Output "Error removing Send On Behalf permissions for $($UPN): $_"
    }}
    Disconnect-ExchangeOnline -Confirm:$false
    if (-not $Found) {{ Write-Output "NO_PERMISSIONS_FOUND" }}
    if ($Success.Count -gt 0) {{ Write-Output "Removed from: $($Success -join ', ')" }}
    if ($Failed.Count -gt 0) {{ Write-Output "Failed on: $($Failed -join ', ')" }}'''

    success, output = execute_powershell(log, ps_command)
    if not success:
        return False, "PowerShell execution failed"

    for line in output.splitlines():
        if line.startswith("DEBUG: Has access to:"):
            log.info(line)
        elif line == "NO_PERMISSIONS_FOUND":
            return True, f"No SendOnBehalfOf permissions found to remove for {user_email}"
        elif line.startswith("Removed from:"):
            for entry in line.replace("Removed from: ", "").split(","):
                e = entry.strip()
                if e:
                    mailbox_name = e.split(":")[0] if ":" in e else e
                    msg = f"{user_email} had SendOnBehalfOf removed from {mailbox_name}"
                    log.info(msg)
                    record_result(log, ResultLevel.SUCCESS, msg)
                    data_to_log.setdefault("Success", "")
                    data_to_log["Success"] += ("\n" if data_to_log["Success"] else "") + msg
        elif line.startswith("Failed on:"):
            content_after_failed = line.replace("Failed on: ", "").strip()
            if content_after_failed:
                entries = [e.strip() for e in content_after_failed.split(",") if e.strip()]
                if entries:
                    for e in entries:
                        mailbox_name = e.split(":")[0] if ":" in e else e
                        msg = f"Failed to remove SendOnBehalfOf from {user_email} on {mailbox_name}"
                        log.warning(msg)
                        record_result(log, ResultLevel.WARNING, msg)
                        data_to_log.setdefault("Failed", "")
                        data_to_log["Failed"] += ("\n" if data_to_log["Failed"] else "") + msg

    return True, f"SendOnBehalfOf permissions removal complete for {user_email}"

def remove_exo_all_sendas_permissions(log, azure_domain, user_email, exo_access_token):
    log.info(f"Removing SendAs permissions for user [{user_email}]")
    ps_command = f"""$ErrorActionPreference = 'Continue'
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false
        $UPN = "{user_email}"
        $Success = @()
        $Failed = @()
        $Found = $false
        try {{
            $permissions = Get-RecipientPermission -Trustee '{user_email}' -ErrorAction SilentlyContinue
            if ($permissions) {{
                $Found = $true
                foreach ($perm in $permissions) {{
                    $resolved = try {{ (Get-Mailbox -Identity $perm.Identity -ErrorAction Stop).UserPrincipalName }} catch {{ $perm.Identity }}
                    Write-Output "DEBUG: Has access to: $resolved"
                    try {{
                        Remove-RecipientPermission -Identity $resolved -Trustee $UPN -AccessRights SendAs -Confirm:$false -SkipDomainValidationForMailContact -SkipDomainValidationForMailUser -SkipDomainValidationForSharedMailbox -ErrorAction Stop
                        $Success += "$($resolved):SendAs"
                    }} catch {{
                        $Failed += "$($resolved):SendAs"
                    }}
                }}
            }}
        }} catch {{
            Write-Output "Error removing Send As permissions for $($UPN): $_"
        }}
        Disconnect-ExchangeOnline -Confirm:$false
        if (-not $Found) {{ Write-Output "NO_PERMISSIONS_FOUND" }}
        if ($Success.Count -gt 0) {{ Write-Output "Removed from: $($Success -join ', ')" }}
        if ($Failed.Count -gt 0) {{ Write-Output "Failed on: $($Failed -join ', ')" }}
    """
    success, output = execute_powershell(log, ps_command)
    if not success:
        return False, "PowerShell execution failed"

    for line in output.splitlines():
        if line.startswith("DEBUG: Has access to:"):
            log.info(line)
        elif line == "NO_PERMISSIONS_FOUND":
            return True, f"No SendAs permissions found to remove for {user_email}"
        elif line.startswith("Removed from:"):
            for entry in line.replace("Removed from: ", "").split(","):
                e = entry.strip()
                if e:
                    mailbox_name = e.split(":")[0] if ":" in e else e
                    msg = f"{user_email} had SendAs removed from {mailbox_name}"
                    log.info(msg)
                    record_result(log, ResultLevel.SUCCESS, msg)
                    data_to_log.setdefault("Success", "")
                    data_to_log["Success"] += ("\n" if data_to_log["Success"] else "") + msg
        elif line.startswith("Failed on:"):
            content_after_failed = line.replace("Failed on: ", "").strip()
            if content_after_failed:
                entries = [e.strip() for e in content_after_failed.split(",") if e.strip()]
                if entries:
                    for e in entries:
                        mailbox_name = e.split(":")[0] if ":" in e else e
                        msg = f"Failed to remove SendAs from {user_email} on {mailbox_name}"
                        log.warning(msg)
                        record_result(log, ResultLevel.WARNING, msg)
                        data_to_log.setdefault("Failed", "")
                        data_to_log["Failed"] += ("\n" if data_to_log["Failed"] else "") + msg

    return True, f"SendAs permissions removal complete for {user_email}"

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1747979988829")
            operation = input.get_value("Operation_1747891002014")
            user_identifier = input.get_value("User_1747891758785")
            mailboxes_raw = input.get_value("Mailboxes_1747891800428")
            access_level = input.get_value("AccessLevel_1750203013306")
            automapping_input = input.get_value("Automapping_1751512618440")
            auth_code = input.get_value("AuthenticationCode_1747979998154")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        operation = operation.strip() if operation else ""
        user_identifier = user_identifier.strip() if user_identifier else ""
        mailboxes_raw = mailboxes_raw.strip() if mailboxes_raw else ""
        access_level = access_level.strip() if access_level else ""
        automapping_input = automapping_input.strip() if automapping_input else "Yes"
        auth_code = auth_code.strip() if auth_code else ""

        log.info(f"Ticket Number = [{ticket_number}]")
        log.info(f"Requested operation = [{operation}]")
        log.info(f"Received input user = [{user_identifier}]")

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number is required but missing")
            return
        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation value is missing or invalid")
            return
        if not user_identifier:
            record_result(log, ResultLevel.WARNING, "User identifier is empty or invalid")
            return

        auto_mapping = True if automapping_input == "Yes" else False

        company_identifier, company_name, company_id, company_types = get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
        if not company_identifier:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve company identifier for ticket [{ticket_number}]")
            return

        if company_identifier == "MIT":
            if not validate_mit_authentication(log, http_client, vault_name, auth_code):
                return

        graph_tenant_id, graph_access_token = get_graph_token(log, http_client, vault_name, company_identifier)
        if not graph_access_token:
            record_result(log, ResultLevel.WARNING, "Failed to obtain MS Graph access token")
            return

        exo_tenant_id, exo_access_token = get_exo_token(log, http_client, vault_name, company_identifier)
        if not exo_access_token:
            record_result(log, ResultLevel.WARNING, "Failed to obtain Exchange Online access token")
            return

        azure_domain = get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")
        if not azure_domain:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve Azure domain for [{company_identifier}]")
            return

        aad_user_result = get_aad_user_data(log, http_client, msgraph_base_url, user_identifier, graph_access_token)
        if isinstance(aad_user_result, list):
            details = "\n".join([f"- {u.get('displayName')} | {u.get('userPrincipalName')} | {u.get('id')}" for u in aad_user_result])
            record_result(log, ResultLevel.WARNING, f"Multiple users found for [{user_identifier}]\n{details}")
            return
        user_id, user_email, user_sam, user_onpremisessyncenabled = aad_user_result

        if not user_id:
            record_result(log, ResultLevel.WARNING, f"Failed to resolve user ID for [{user_identifier}]")
            return
        if not user_email:
            record_result(log, ResultLevel.WARNING, f"Unable to resolve user email for [{user_identifier}]")
            return

        log.info(f"Operation: {operation}, User: {user_email}, Mailboxes: {mailboxes_raw}, Access: {access_level}")

        if operation == "Add User Mailbox Permissions":
            mailboxes = [m.strip() for m in mailboxes_raw.split(",") if m.strip()]
            if not mailboxes or not access_level:
                record_result(log, ResultLevel.WARNING, "No valid mailbox or access provided")
                return

            for mailbox in mailboxes:
                results = []
                if access_level == "Send on Behalf":
                    success, res = add_exo_sendonbehalfof_permissions(log, azure_domain, user_email, user_id, mailbox, exo_access_token)
                    results.extend(res)
                elif access_level == "Send As":
                    success, res = add_exo_sendas_permissions(log, azure_domain, user_email, mailbox, exo_access_token)
                    results.extend(res)
                else:
                    mapped_access = access_level.replace(" ", "")
                    success, res = add_exo_mailbox_permissions(log, azure_domain, user_email, mailbox, mapped_access, exo_access_token, auto_mapping)
                    results.extend(res)
                for level, msg in results:
                    record_result(log, level, msg)

        elif operation == "Remove Mailbox Permissions":
            mailboxes = [m.strip() for m in mailboxes_raw.split(",") if m.strip()]
            if not mailboxes or not access_level:
                record_result(log, ResultLevel.WARNING, "No valid mailbox or access provided to remove")
                return

            for mailbox in mailboxes:
                mapped_access = access_level.replace(" ", "")
                success, results = remove_exo_mailbox_permissions(log, azure_domain, user_email, mailbox, mapped_access, exo_access_token)
                for level, msg in results:
                    record_result(log, level, msg)

        elif operation == "Remove All Mailbox Permissions":
            log.info("Starting: Remove All Explicit Mailbox Permissions")
            success, message = remove_exo_all_mailbox_permissions(log, azure_domain, user_email, exo_access_token)
            if "No " in message:
                if user_email in message:
                    record_result(log, ResultLevel.SUCCESS, message)
                else:
                    record_result(log, ResultLevel.SUCCESS, f"{message} for [{user_email}]")
            else:
                record_result(log, ResultLevel.SUCCESS if success else ResultLevel.WARNING, message)
            log.info("Completed: Remove All Mailbox Permissions")

        elif operation == "Remove All SendAs Permissions":
            log.info("Starting: Remove All SendAs Permissions")
            success, message = remove_exo_all_sendas_permissions(log, azure_domain, user_email, exo_access_token)
            if "No " in message:
                if user_email in message:
                    record_result(log, ResultLevel.SUCCESS, message)
                else:
                    record_result(log, ResultLevel.SUCCESS, f"{message} for [{user_email}]")
            else:
                record_result(log, ResultLevel.SUCCESS if success else ResultLevel.WARNING, message)
            log.info("Completed: Remove All SendAs Permissions")

        elif operation == "Remove All SendOnBehalfOf Permissions":
            log.info("Starting: Remove All SendOnBehalfOf Permissions")
            success, message = remove_exo_all_sendonbehalfof_permissions(log, azure_domain, user_email, exo_access_token)
            if "No " in message:
                if user_email in message:
                    record_result(log, ResultLevel.SUCCESS, message)
                else:
                    record_result(log, ResultLevel.SUCCESS, f"{message} for [{user_email}]")
            else:
                record_result(log, ResultLevel.SUCCESS if success else ResultLevel.WARNING, message)
            log.info("Completed: Remove All SendOnBehalfOf Permissions")
        
        if "Failed" in data_to_log and data_to_log["Failed"].strip():
            for line in data_to_log["Failed"].split("\n"):
                line = line.strip()
                if line and line != "Failed on:":
                    record_result(log, ResultLevel.WARNING, line)

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()