import sys
import traceback
import json
import random
import re
import subprocess
import os
import io
import base64
import hashlib
import time
import urllib.parse
import string
import requests
import pandas as pd
from datetime import datetime, timedelta, timezone
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
                response = getattr(http_client.third_party_integration(integration_name), method)(url=endpoint, json=data) if data else getattr(http_client.third_party_integration(integration_name), method)(url=endpoint)
            else:
                request_args = {"url": endpoint}
                if params:
                    request_args["params"] = params
                if headers:
                    request_args["headers"] = headers
                if data:
                    if headers and headers.get("Content-Type") == "application/x-www-form-urlencoded":
                        request_args["data"] = data
                    else:
                        request_args["json"] = data
                response = getattr(requests, method)(**request_args)

            if response.status_code in [200, 204]:
                return response
            if response.status_code in [429, 503]:
                retry_after = response.headers.get("Retry-After")
                wait_time = int(retry_after) if retry_after else base_delay * (2 ** attempt) + random.uniform(0, 3)
                log.warning(f"Rate limit exceeded. Retrying in {wait_time:.2f} seconds")
                time.sleep(wait_time)
            elif response.status_code == 404:
                log.warning(f"Skipping non-existent resource [{endpoint}]")
                return None
            else:
                log.error(f"API request failed Status: {response.status_code}, Response: {response.text}")
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

def get_company_identifier_from_ticket(log, http_client, cwpsa_base_url, ticket_number):
    log.info(f"Retrieving company identifier for ticket [{ticket_number}]")
    endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}"

    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")

    if response:
        if response.status_code == 200:
            data = response.json()
            company_identifier = data.get("company", {}).get("identifier", "")
            if company_identifier:
                log.info(f"Company identifier for ticket [{ticket_number}] is [{company_identifier}]")
                return company_identifier
            else:
                log.error(f"Company identifier not found in response for ticket [{ticket_number}]")
        else:
            log.error(f"Failed to retrieve company identifier for ticket [{ticket_number}] Status: {response.status_code}, Body: {response.text}")
    else:
        log.error(f"Failed to retrieve company identifier for ticket [{ticket_number}]: No response received")

    return ""

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

def get_user_data(log, http_client, msgraph_base_url, user_identifier, token):
    log.info(f"Resolving user ID and email for [{user_identifier}]")
    headers = {"Authorization": f"Bearer {token}"}

    if re.fullmatch(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}", user_identifier):
        endpoint = f"{msgraph_base_url}/users/{user_identifier}"
        response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
        if response and response.status_code == 200:
            user = response.json()
            user_id = user.get("id")
            user_email = user.get("userPrincipalName", "")
            user_sam = user.get("onPremisesSamAccountName", "")
            log.info(f"User found by object ID [{user_id}] with email [{user_email}] and SAM [{user_sam}]")
            return user_id, user_email, user_sam
        log.error(f"No user found with object ID [{user_identifier}]")
        return "", "", ""

    filters = [
        f"startswith(displayName,'{user_identifier}')",
        f"startswith(userPrincipalName,'{user_identifier}')",
        f"startswith(mail,'{user_identifier}')"
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
            user_id = user.get("id")
            user_email = user.get("userPrincipalName", "")
            user_sam = user.get("onPremisesSamAccountName", "")
            log.info(f"User found for [{user_identifier}] - ID: {user_id}, Email: {user_email}, SAM: {user_sam}")
            return user_id, user_email, user_sam

    log.error(f"Failed to resolve user ID and email for [{user_identifier}]")
    return "", "", ""

def add_sendonbehalfof_permissions(log, azure_domain, user_email, user_id, mailbox, exo_access_token):
    log.info(f"Adding SendOnBehalfOf permission for [{user_email}] on mailbox [{mailbox}]")

    ps_command = f"""
    $ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false

    try {{
        Set-Mailbox -Identity "{mailbox}" -GrantSendOnBehalfTo "{user_id}" -Confirm:$false -ErrorAction Stop
        Write-Output "Added SendOnBehalfOf permission for {user_email} on {mailbox}"
    }} catch {{
        Write-Output "Failed to add SendOnBehalfOf permission for {user_email} on {mailbox}. Error: $_"
        throw $_
    }}

    Disconnect-ExchangeOnline -Confirm:$false
    """

    success, _ = execute_powershell(log, ps_command)
    log_entry = f"{'success' if success else 'failed'} : {mailbox}:SendOnBehalfOf"
    return success, log_entry

def add_sendas_permissions(log, azure_domain, user_email, mailbox, exo_access_token):
    log.info(f"Adding SendAs permission for [{user_email}] on mailbox [{mailbox}]")

    ps_command = f"""
    $ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false

    try {{
        Add-RecipientPermission -Identity "{mailbox}" -Trustee "{user_email}" -AccessRights SendAs -Confirm:$false -ErrorAction Stop
        Write-Output "Added SendAs permission for {user_email} on {mailbox}"
    }} catch {{
        Write-Output "Failed to add SendAs permission for {user_email} on {mailbox}. Error: $_"
        throw $_
    }}

    Disconnect-ExchangeOnline -Confirm:$false
    """

    success, _ = execute_powershell(log, ps_command)
    log_entry = f"{'success' if success else 'failed'} : {mailbox}:SendAs"
    return success, log_entry

def add_mailbox_permissions(log, azure_domain, user_email, mailbox, access, exo_access_token):
    log.info(f"Adding mailbox permissions for user [{user_email}] on mailbox [{mailbox}] with access level [{access}]")

    ps_command = f"""
    $ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false

    try {{
        Add-MailboxPermission -Identity "{mailbox}" -User "{user_email}" -AccessRights {access} -Confirm:$false -ErrorAction Stop
        Write-Output "Added {access} permission for {user_email} on mailbox {mailbox}"
    }} catch {{
        Write-Output "Failed to add {access} permission for {user_email} on mailbox {mailbox}. Error: $_"
        throw $_
    }}

    Disconnect-ExchangeOnline -Confirm:$false
    """

    success, _ = execute_powershell(log, ps_command)
    log_entry = f"{'success' if success else 'failed'} : {mailbox}:{access}"
    return success, log_entry

def remove_user_mailbox_permissions(log, azure_domain, user_email, exo_access_token):
    log.info(f"Removing all mailbox permissions for user [{user_email}]")

    ps_command = f"""
    $ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false

    $UserToRemove = "{user_email}"
    $Mailboxes = Get-Mailbox -ResultSize Unlimited

    foreach ($Mailbox in $Mailboxes) {{
        try {{
            Remove-MailboxPermission -Identity $Mailbox.PrimarySmtpAddress -User $UserToRemove -AccessRights FullAccess -Confirm:$false -ErrorAction SilentlyContinue
            Remove-RecipientPermission -Identity $Mailbox.PrimarySmtpAddress -Trustee $UserToRemove -AccessRights SendAs -Confirm:$false -ErrorAction SilentlyContinue
            Set-Mailbox -Identity $Mailbox.PrimarySmtpAddress -GrantSendOnBehalfTo @{{Remove=$UserToRemove}} -ErrorAction SilentlyContinue
        }} catch {{
            Write-Host "Error removing permissions for $UserToRemove on $($Mailbox.PrimarySmtpAddress): $_"
        }}
    }}

    Disconnect-ExchangeOnline -Confirm:$false
    """

    success, _ = execute_powershell(log, ps_command)
    log_entry = f"{'success' if success else 'failed'} : removed all mailbox permissions for {user_email}"
    return success, log_entry

def remove_specific_mailbox_permission(log, azure_domain, user_email, mailbox, access, exo_access_token):
    log.info(f"Removing [{access}] permission for user [{user_email}] on mailbox [{mailbox}]")

    if access == "FullAccess":
        cmd = f'Remove-MailboxPermission -Identity "{mailbox}" -User "{user_email}" -AccessRights FullAccess -Confirm:$false -ErrorAction Stop'
    elif access == "SendAs":
        cmd = f'Remove-RecipientPermission -Identity "{mailbox}" -Trustee "{user_email}" -AccessRights SendAs -Confirm:$false -ErrorAction Stop'
    elif access == "SendOnBehalfOf":
        cmd = f'Set-Mailbox -Identity "{mailbox}" -GrantSendOnBehalfTo @{{Remove="{user_email}"}} -ErrorAction Stop'
    else:
        log_entry = f"failed : {mailbox}:{access}"
        return False, log_entry

    ps_command = f"""
    $ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false

    try {{
        {cmd}
        Write-Output "Removed {access} permission for {user_email} on {mailbox}"
    }} catch {{
        Write-Output "Failed to remove {access} permission for {user_email} on {mailbox}. Error: $_"
        throw $_
    }}

    Disconnect-ExchangeOnline -Confirm:$false
    """

    success, _ = execute_powershell(log, ps_command)
    log_entry = f"{'success' if success else 'failed'} : {mailbox}:{access}"
    return success, log_entry

def parse_permissions(log, permissions_string):
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

def main():
    try:
        try:
            operation = input.get_value("Operation_1747891002014")
            user_identifier = input.get_value("User_1747891758785")
            mailboxes_and_access = input.get_value("MailboxesandAccessLevel_1747891800428")
            ticket_number = input.get_value("TicketNumber_1747979988829")
            auth_code = input.get_value("AuthCode_1747979998154")
            provided_token = input.get_value("AccessToken_1747980009090")
        except Exception as e:
            record_result(log, ResultLevel.WARNING, f"Failed to fetch input values: {e}")
            return
        
        user_identifier = user_identifier.strip() if user_identifier else ""
        mailboxes_and_access = mailboxes_and_access.strip() if mailboxes_and_access else ""
        ticket_number = ticket_number.strip() if ticket_number else ""
        auth_code = auth_code.strip() if auth_code else ""
        provided_token = provided_token.strip() if provided_token else ""

        log.info(f"Received input user = [{user_identifier}], Ticket = [{ticket_number}]")

        if not user_identifier or not user_identifier.strip():
            record_result(log, ResultLevel.WARNING, "User identifier is empty or invalid")
            return

        if provided_token and provided_token.strip():
            access_token = provided_token.strip()
            log.info("Using provided access token")
            if not isinstance(access_token, str) or "." not in access_token:
                record_result(log, ResultLevel.WARNING, "Provided access token is malformed (missing dots)")
                return

        elif ticket_number:
            company_identifier = get_company_identifier_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
            if not company_identifier:
                record_result(log, ResultLevel.WARNING, f"Failed to retrieve company identifier for ticket [{ticket_number}]")
                return

            if company_identifier == "MIT":
                if not validate_mit_authentication(log, http_client, vault_name, auth_code):
                    return

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

            access_token = get_access_token(log, http_client, tenant_id, client_id, client_secret)
            if not isinstance(access_token, str) or "." not in access_token:
                record_result(log, ResultLevel.WARNING, "Access token is malformed (missing dots)")
                return

        else:
            record_result(log, ResultLevel.WARNING, "Either Access Token or Ticket Number is required")
            return
        
        exo_client_id = get_secret_value(log, http_client, vault_name, f"{company_identifier}-ExchangeApp-ClientID")
        exo_client_secret = get_secret_value(log, http_client, vault_name, f"{company_identifier}-ExchangeApp-ClientSecret")

        if not all([exo_client_id, exo_client_secret, azure_domain]):
            record_result(log, ResultLevel.WARNING, "Failed to retrieve required internal secrets")
            return

        exo_access_token = get_access_token(log, http_client, tenant_id, exo_client_id, exo_client_secret, scope="https://outlook.office365.com/.default")
        if not isinstance(exo_access_token, str) or "." not in exo_access_token:
            record_result(log, ResultLevel.WARNING, "Exchange Online access token is malformed (missing dots)")
            return

        user_result = get_user_data(log, http_client, msgraph_base_url, user_identifier, access_token)

        if isinstance(user_result, list):
            details = "\n".join([f"- {u.get('displayName')} | {u.get('userPrincipalName')} | {u.get('id')}" for u in user_result])
            record_result(log, ResultLevel.WARNING, f"Multiple users found for [{user_identifier}]\n{details}")
            return
        
        user_id, user_email, user_sam = user_result
        if not user_id:
            record_result(log, ResultLevel.WARNING, f"Failed to resolve user ID for [{user_identifier}]")
            return
        if not user_email:
            record_result(log, ResultLevel.WARNING, f"Unable to resolve user email for [{user_identifier}]")
            return
        
        permissions_added = []
        log.info(f"Operation: {operation}, User: {user_email}, Mailboxes: {mailboxes_and_access}")

        if operation == "Add Mailbox Permissions":
            log.info(f"Adding mailbox permissions for user [{user_email}] on mailboxes [{mailboxes_and_access}]")
            permissions = parse_permissions(log, mailboxes_and_access)
            log.info(f"Parsed permissions: {permissions}")
            if not permissions:
                record_result(log, ResultLevel.WARNING, "No valid mailbox permissions provided")
                return

            log_entries = []
            overall_success = False

            for permission in permissions:
                mailbox = permission.get("mailbox")
                access = permission.get("access")
                if not mailbox or not access:
                    log_entries.append(f"failed : {mailbox or 'unknown'}:{access or 'unknown'}")
                    continue
                if access not in ["FullAccess", "ReadPermission", "SendOnBehalfOf", "SendAs"]:
                    log_entries.append(f"failed : {mailbox}:{access}")
                    continue
                if access == "SendOnBehalfOf":
                    success, entry = add_sendonbehalfof_permissions(log, azure_domain, user_email, user_id, mailbox, exo_access_token)
                elif access == "SendAs":
                    success, entry = add_sendas_permissions(log, azure_domain, user_email, mailbox, exo_access_token)
                else:
                    success, entry = add_mailbox_permissions(log, azure_domain, user_email, mailbox, access, exo_access_token)

                log_entries.append(entry)
                if success:
                    overall_success = True

            result_level = ResultLevel.SUCCESS if overall_success else ResultLevel.WARNING
            record_result(log, result_level, ", ".join(log_entries))
            data_to_log["Result"] = "Success" if overall_success else "Fail"
            data_to_log["Details"] = ", ".join(log_entries)
            log.result_data(data_to_log)

        elif operation == "Remove All Mailbox Permissions":
            if not user_email:
                record_result(log, ResultLevel.WARNING, "User email is required for removing mailbox permissions")
                return
            success, entry = remove_user_mailbox_permissions(log, azure_domain, user_email, exo_access_token)
            result_level = ResultLevel.SUCCESS if success else ResultLevel.WARNING
            record_result(log, result_level, entry)
            data_to_log["Result"] = "Success" if success else "Fail"
            data_to_log["Details"] = entry
            log.result_data(data_to_log)

        elif operation == "Remove Mailbox Permission":
            permissions = parse_permissions(log, mailboxes_and_access)
            log.info(f"Parsed permissions to remove: {permissions}")
            if not permissions:
                record_result(log, ResultLevel.WARNING, "No valid mailbox permissions provided to remove")
                return

            log_entries = []
            overall_success = False

            for permission in permissions:
                mailbox = permission.get("mailbox")
                access = permission.get("access")
                if not mailbox or not access:
                    log_entries.append(f"failed : {mailbox or 'unknown'}:{access or 'unknown'}")
                    continue
                success, entry = remove_specific_mailbox_permission(log, azure_domain, user_email, mailbox, access, exo_access_token)
                log_entries.append(entry)
                if success:
                    overall_success = True

            result_level = ResultLevel.SUCCESS if overall_success else ResultLevel.WARNING
            record_result(log, result_level, ", ".join(log_entries))
            data_to_log["Result"] = "Success" if overall_success else "Fail"
            data_to_log["Details"] = ", ".join(log_entries)
            log.result_data(data_to_log)

            record_result(log, ResultLevel.WARNING, "Operation [Remove Mailbox Permission] is not implemented yet")
            return

    except Exception:
        log.exception("An error occurred while processing")
        log.result_message(ResultLevel.WARNING, "Process failed")

if __name__ == "__main__":
    main()