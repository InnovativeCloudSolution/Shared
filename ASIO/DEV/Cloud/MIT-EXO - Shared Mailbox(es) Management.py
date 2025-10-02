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
vault_name = "mit-azu1-prod1-akv1"

data_to_log = {}
bot_name = "MIT-EXO - Shared Mailbox(es) Creation"
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
        log.error(f"Failed to retrieve ticket [{ticket_number}] Status: {ticket_response.status_code}, Body: {ticket_response.text}")
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
        if response:
            user = response.json()
            return (
                user.get("id", ""),
                user.get("userPrincipalName", ""),
                user.get("onPremisesSamAccountName", ""),
                user.get("onPremisesSyncEnabled", False),
                user.get("displayName", "")
            )
        return "", "", "", False, ""

    filters = [
        f"startswith(displayName,'{user_identifier}')",
        f"startswith(userPrincipalName,'{user_identifier}')",
        f"startswith(mail,'{user_identifier}')",
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
            return (
                user.get("id", ""),
                user.get("userPrincipalName", ""),
                user.get("onPremisesSamAccountName", ""),
                user.get("onPremisesSyncEnabled", False),
                user.get("displayName", "")
            )
    log.error(f"Failed to resolve user ID and email for [{user_identifier}]")
    return "", "", "", False, ""

def get_user_data(log, http_client, msgraph_base_url, user_identifier, token):
    user_data = get_aad_user_data(log, http_client, msgraph_base_url, user_identifier, token)
    if isinstance(user_data, list):
        return user_data
    user_id, user_email, user_sam, _, user_display_name = user_data
    return user_id, user_email, user_sam, user_display_name
    log.info(f"Parsing permissions string: {permissions_string}")
    if not permissions_string:
        return []

    permissions = []
    items = permissions_string.split(",")
    for item in items:
        item = item.strip()
        if ":" in item:
            mailbox, access = item.split(":", 1)
            mailbox = mailbox.strip()
            access = access.strip()
            if mailbox and access:
                permissions.append({"mailbox": mailbox, "access": access})
            else:
                log.warning(f"Invalid permission format: {item}")
        else:
            log.warning(f"Invalid permission format (missing colon): {item}")

    return permissions

def create_shared_mailbox(log, azure_domain, username, user_displayname, mailbox_domain, access_token, exo_access_token):
    log.info(f"Creating shared mailbox [{username}@{mailbox_domain}]")
    mailbox = f"{username}@{mailbox_domain}"

    ps_command = f"""
    $ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false
    Write-Output "Processing SendOnBehalf for mailbox $mailbox."

    $mailbox = "{mailbox}"
    $mailboxDomain = "{mailbox_domain}"
    $userDisplayName = "{user_displayname}"
    $displayNameFromDomain = ($mailboxDomain.split('.')[0].substring(0, 1).ToUpper() + $mailboxDomain.split('.')[0].substring(1).ToLower())
    $mailboxDisplayName = "$($userDisplayName) - $($displayNameFromDomain)"

    # Create the shared mailbox
    try {{
        New-Mailbox -Shared -Name $mailboxDisplayName -DisplayName $mailboxDisplayName -PrimarySmtpAddress $mailbox -Confirm:$false -ErrorAction Stop
        Write-Output "Created shared mailbox $mailbox with display name $mailboxDisplayName."
    }} catch {{
        Write-Output "Failed to create Shared Mailbox {mailbox} Error: $_"
        throw $_
    }}

    Disconnect-ExchangeOnline -Confirm:$false
    """
    success, output = execute_powershell(log, ps_command)
    if not success:
        log.warning(f"Failed to add mailbox permissions: {output}")
        return False, output
    return True, output

def set_shared_mailbox_save_sent_items(log, azure_domain, shared_mailbox, access_token, exo_access_token):
    log.info(f"Setting 'SendAsItemsCopiedTo' for shared mailbox [{shared_mailbox}] to 'SenderAndFrom'")

    ps_command = f"""
    $ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false
    Write-Output "Processing SendOnBehalf for mailbox $mailbox."

    $mailbox = "{shared_mailbox}"

    # Create the shared mailbox
    try {{
        Set-Mailbox -Identity $mailbox -MessageCopyForSentAsEnabled $True -Confirm:$false -ErrorAction Stop
        Write-Output "Set 'SendAsItemsCopiedTo' for shared mailbox [$($mailbox)] to 'SenderAndFrom'"
    }} catch {{
        Write-Output "Failed to set 'SendAsItemsCopiedTo' for shared mailbox [$($mailbox)] to 'SenderAndFrom'. Error: $_"
        throw $_
    }}

    Disconnect-ExchangeOnline -Confirm:$false
    """
    success, output = execute_powershell(log, ps_command)
    if not success:
        log.warning(f"Failed to add mailbox permissions: {output}")
        return False, output
    return True, output

def add_sendas_permissions(log, azure_domain, user_email, mailbox, exo_access_token):
    log.info(f"Adding mailbox permissions for user [{user_email}] on mailbox [{mailbox}] with access level [SendAs]")

    ps_command = f"""
    $ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false
    Write-Output "Processing SendOnBehalf for mailbox $mailbox."

    $mailbox = "{mailbox}"
    $userEmail = "{user_email}"

    # Add the mailbox permission
    try {{
        Add-RecipientPermission -Identity $mailbox -Trustee $userEmail -AccessRights SendAs -Confirm:$false -ErrorAction Stop
        Write-Output "Added $accessLevel permission for $userEmail on mailbox $mailbox."
    }} catch {{
        Write-Output "Failed to add $accessLevel permission for $userEmail on mailbox $mailbox. Error: $_"
        throw $_            
    }}

    Disconnect-ExchangeOnline -Confirm:$false
    """

    success, output = execute_powershell(log, ps_command)
    if not success:
        log.error(f"Failed to add mailbox permissions: {output}")
        return False, output
    return True, output

def add_mailbox_permissions(log, azure_domain, user_email, mailbox, access, exo_access_token):
    log.info(f"Adding mailbox permissions for user [{user_email}] on mailbox [{mailbox}] with access level [{access}]")

    ps_command = f"""
    $ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false
    Write-Output "Processing SendOnBehalf for mailbox $mailbox."

    $mailbox = "{mailbox}"
    $userEmail = "{user_email}"
    $accessLevel = "{access}"

    # Add the mailbox permission
    try {{
        Add-MailboxPermission -Identity $mailbox -User $userEmail -AccessRights $accessLevel -Confirm:$false -ErrorAction Stop
        Write-Output "Added $accessLevel permission for $userEmail on mailbox $mailbox."
    }} catch {{
        Write-Output "Failed to add $accessLevel permission for $userEmail on mailbox $mailbox. Error: $_"
        throw $_
    }}

    Disconnect-ExchangeOnline -Confirm:$false
    """

    success, output = execute_powershell(log, ps_command)
    if not success:
        log.error(f"Failed to add mailbox permissions: {output}")
        return False, output
    return True, output

def create_shared_mailbox_for_user(log, http_client, msgraph_base_url, user_identifier, mailbox_domains, azure_domain, graph_access_token, exo_access_token):
    user_result = get_user_data(log, http_client, msgraph_base_url, user_identifier, graph_access_token)

    if isinstance(user_result, list):
        details = "\n".join([f"- {u.get('displayName')} | {u.get('userPrincipalName')} | {u.get('id')}" for u in user_result])
        record_result(log, ResultLevel.WARNING, f"Multiple users found for [{user_identifier}]\n{details}")
        return

    user_id, user_email, user_sam, user_displayname = user_result
    if not user_id:
        record_result(log, ResultLevel.WARNING, f"Failed to resolve user ID for [{user_identifier}]")
        return
    if not user_email:
        record_result(log, ResultLevel.WARNING, f"Unable to resolve user email for [{user_identifier}]")
        return

    mailboxes_created = []
    if not mailbox_domains:
        record_result(log, ResultLevel.WARNING, "No mailboxes provided for permission assignment")
        return

    username = user_email.split("@")[0] if "@" in user_email else ""
    mailbox_domains = mailbox_domains.split(",")
    mailbox_domains = [mailbox.strip() for mailbox in mailbox_domains if mailbox.strip()]

    for mailbox_domain in mailbox_domains:
        success, output = create_shared_mailbox(log, azure_domain, username, user_displayname, mailbox_domain, graph_access_token, exo_access_token)
        if success:
            log.info(f"Successfully created shared mailbox {username}@{mailbox_domain}")
            mailboxes_created.append({'mailbox': f"{username}@{mailbox_domain}", 'displayName': user_displayname})
        else:
            record_result(log, ResultLevel.WARNING, f"Failed to create shared mailbox for {username} at {mailbox_domain}")
            continue

        success, output = set_shared_mailbox_save_sent_items(log, azure_domain, f"{username}@{mailbox_domain}", graph_access_token, exo_access_token)
        if success:
            log.info(f"Successfully set 'SendAsItemsCopiedTo' for shared mailbox {username}@{mailbox_domain}")
            mailboxes_created[-1]['saveSentItems'] = True
        else:
            record_result(log, ResultLevel.WARNING, f"Failed to set 'SendAsItemsCopiedTo' for shared mailbox {username}@{mailbox_domain}")

        success, output = add_mailbox_permissions(log, azure_domain, user_email, f"{username}@{mailbox_domain}", "FullAccess", exo_access_token)
        if success:
            log.info(f"Successfully added FullAccess permission for {user_email} on mailbox {username}@{mailbox_domain}")
            mailboxes_created[-1]['fullAccess'] = True
        else:
            record_result(log, ResultLevel.WARNING, f"Failed to add FullAccess permission for {user_email} on mailbox {username}@{mailbox_domain}")

        success, output = add_sendas_permissions(log, azure_domain, user_email, f"{username}@{mailbox_domain}", exo_access_token)
        if success:
            log.info(f"Successfully added SendAs permission for {user_email} on mailbox {username}@{mailbox_domain}")
            mailboxes_created[-1]['sendAs'] = True
        else:
            record_result(log, ResultLevel.WARNING, f"Failed to add SendAs permission for {user_email} on mailbox {username}@{mailbox_domain}")

    if mailboxes_created:
        record_result(log, ResultLevel.SUCCESS, f"Successfully created shared mailboxes: {mailboxes_created}")
        data_to_log.update({"mailboxes_created": mailboxes_created})
    else:
        record_result(log, ResultLevel.WARNING, "No shared mailboxes were created successfully")

def create_generic_shared_mailbox(log, mailbox_name, display_name, azure_domain, exo_access_token):
    log.info(f"Creating generic shared mailbox [{mailbox_name}] with display name [{display_name}]")
    
    ps_command = f"""
      $ErrorActionPreference = 'Stop'
      Import-Module ExchangeOnlineManagement
      Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false
      
      $mailbox = "{mailbox_name}"
      $displayName = "{display_name}"
      
      try {{
          New-Mailbox -Shared -Name $displayName -DisplayName $displayName -PrimarySmtpAddress $mailbox -Confirm:$false -ErrorAction Stop
          Write-Output "Created shared mailbox $mailbox with display name $displayName."
      }} catch {{
          Write-Output "Failed to create Shared Mailbox $mailbox Error: $_"
          throw $_
      }}
      
      try {{
          Set-Mailbox -Identity $mailbox -MessageCopyForSentAsEnabled $True -Confirm:$false -ErrorAction Stop
          Write-Output "Set 'SendAsItemsCopiedTo' for shared mailbox $mailbox to 'SenderAndFrom'"
      }} catch {{
          Write-Output "Failed to set 'SendAsItemsCopiedTo' for shared mailbox $mailbox to 'SenderAndFrom'. Error: $_"
      }}
      
      Disconnect-ExchangeOnline -Confirm:$false
    """
    
    success, output = execute_powershell(log, ps_command)
    if success:
        record_result(log, ResultLevel.SUCCESS, f"Successfully created shared mailbox [{mailbox_name}]")
        data_to_log.update({"mailbox_created": {"mailbox": mailbox_name, "displayName": display_name}})
    else:
        record_result(log, ResultLevel.WARNING, f"Failed to create shared mailbox [{mailbox_name}]: {output}")

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1758573299776")
            operation = input.get_value("Operation_1758573324207")
            user_identifier = input.get_value("User_1748577269611")
            mailbox_domains = input.get_value("MailboxesDomain_1748577270369")
            mailbox_name = input.get_value("MailboxName_1758573637996")
            display_name = input.get_value("DisplayName_1748577271135")
            auth_code = input.get_value("AuthenticationCode_1748577271943")
            graph_token = input.get_value("GraphToken_1748577323864")
            exo_token = input.get_value("EXOToken_1758573814965")
        except Exception as e:
            record_result(log, ResultLevel.WARNING, f"Failed to fetch input values: {e}")
            return
        
        ticket_number = ticket_number.strip() if ticket_number else ""
        operation = operation.strip() if operation else ""
        user_identifier = user_identifier.strip() if user_identifier else ""
        mailbox_domains = mailbox_domains.strip() if mailbox_domains else ""
        mailbox_name = mailbox_name.strip() if mailbox_name else ""
        display_name = display_name.strip() if display_name else ""
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

        if operation == "Create Shared Mailbox for a User":
            if not user_identifier:
                record_result(log, ResultLevel.WARNING, "User identifier is required for user-based mailbox creation")
                return
            create_shared_mailbox_for_user(log, http_client, msgraph_base_url, user_identifier, mailbox_domains, azure_domain, graph_access_token, exo_access_token)
            
        elif operation == "Create a Shared Mailbox":
            if not mailbox_name:
                record_result(log, ResultLevel.WARNING, "Mailbox name is required for generic mailbox creation")
                return
            if not display_name:
                record_result(log, ResultLevel.WARNING, "Display name is required for generic mailbox creation")
                return
            create_generic_shared_mailbox(log, mailbox_name, display_name, azure_domain, exo_access_token)

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
