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
log.info("Static variables set")

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

def get_group(log, http_client, msgraph_base_url, group_identifier, token):
    log.info(f"Resolving group email for [{group_identifier}]")
    headers = {"Authorization": f"Bearer {token}"}

    filters = [
        f"startswith(displayName,'{group_identifier}')",
        f"startswith(mail,'{group_identifier}')"
    ]
    filter_query = " or ".join(filters)
    endpoint = f"{msgraph_base_url}/groups?$filter={urllib.parse.quote(filter_query)}"

    response = execute_api_call(log, http_client, "get", endpoint, headers=headers)

    if response and response.status_code == 200:
        groups = response.json().get("value", [])
        if len(groups) > 1:
            log.error(f"Multiple groups found for [{group_identifier}]")
            return groups
        if groups:
            group = groups[0]
            group_email = group.get("mail", "")
            group_id = group.get("id", "")
            log.info(f"Group found for [{group_identifier}] - Email: {group_email}, ID: {group_id}")
            return group_email, group_id

    log.error(f"Failed to resolve group for [{group_identifier}]")
    return "", ""

def get_user_groups(log, http_client, msgraph_base_url, user_id, access_token):
    log.info(f"Fetching distribution/mail-enabled groups for user ID [{user_id}]")
    headers = {"Authorization": f"Bearer {access_token}"}
    endpoint = f"{msgraph_base_url}/users/{user_id}/memberOf?$select=id,displayName,mail,groupTypes,securityEnabled,mailEnabled"

    groups = []
    while endpoint:
        response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
        if not response or response.status_code != 200:
            log.error(f"Failed to retrieve group memberships for user ID [{user_id}]")
            break

        data = response.json()
        for group in data.get("value", []):
            if group.get("@odata.type") != "#microsoft.graph.group":
                continue
            is_dynamic = "DynamicMembership" in group.get("groupTypes", [])
            is_mail_enabled = group.get("mailEnabled", False)
            is_security_enabled = group.get("securityEnabled", False)
            is_distribution = is_mail_enabled and not is_security_enabled

            if (is_mail_enabled or is_distribution) and not is_dynamic:
                group_email = group.get("mail", "")
                group_id = group.get("id", "")
                if group_email and group_id:
                    groups.append((group_email, group_id))

        next_link = data.get("@odata.nextLink", "")
        endpoint = next_link if next_link else None

    log.info(f"Total distribution/mail-enabled groups found: {len(groups)}")
    return groups

def add_user_to_groups(log, azure_domain, user_email, group_tuples, exo_access_token):
    group_array = "@(" + ",".join([f"'{g[0]}'" for g in group_tuples]) + ")"

    ps_command = f"""
    $ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false
    $groups = {group_array}
    foreach ($group in $groups) {{
        try {{
            Add-DistributionGroupMember -Identity $group -Member '{user_email}'
            Write-Output "Success: Added {user_email} to $group"
        }} catch {{
            Write-Output "Error: Failed to add {user_email} to $group"
        }}
    }}
    Disconnect-ExchangeOnline -Confirm:$false
    """

    success, output = execute_powershell(log, ps_command)
    log.info("Add group PowerShell executed")
    return success, output

def remove_user_from_groups(log, azure_domain, user_email, group_tuples, exo_access_token):
    group_array = "@(" + ",".join([f"'{g[0]}'" for g in group_tuples]) + ")"

    ps_command = f"""
    $ErrorActionPreference = 'Stop'
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false
    $groups = {group_array}
    foreach ($group in $groups) {{
        try {{
            Remove-DistributionGroupMember -Identity $group -Member '{user_email}' -Confirm:$false
            Write-Output "Success: Removed $group"
        }} catch {{
            Write-Output "Error: Failed to remove $group"
        }}
    }}
    Disconnect-ExchangeOnline -Confirm:$false
    """

    success, output = execute_powershell(log, ps_command)
    log.info("Remove group PowerShell executed")
    return success, output

def remove_user_from_all_groups(log, http_client, msgraph_base_url, azure_domain, user_id, user_email, exo_access_token, access_token):
    log.info(f"Removing user [{user_email}] from all distribution/mail-enabled groups")
    groups = get_user_groups(log, http_client, msgraph_base_url, user_id, access_token)

    if not groups:
        log.info(f"No distribution groups found for user [{user_email}]")
        return True, ""

    success, output = remove_user_from_groups(log, azure_domain, user_email, groups, exo_access_token)
    return success, output

def main():
    try:
        try:
            user_identifier = input.get_value("User_1744665516701")
            ticket_number = input.get_value("TicketNumber_1744665526566")
            auth_code = input.get_value("AuthCode_1744665534118")
            provided_token = input.get_value("AccessToken_1744665535547")
            user_groups = input.get_value("Groups_1744665522057")
            action_input = input.get_value("Action_1744668429969")
        except Exception as e:
            log.exception(e, "Failed to fetch input values")
            log.result_message(ResultLevel.FAILED, "Failed to fetch input values")
            return

        user_identifier = user_identifier.strip() if user_identifier else ""
        ticket_number = ticket_number.strip() if ticket_number else ""
        auth_code = auth_code.strip() if auth_code else ""
        provided_token = provided_token.strip() if provided_token else ""
        user_groups = user_groups.strip() if user_groups else ""
        action_input = action_input.strip().lower() if action_input else ""

        log.info(f"Received input user = [{user_identifier}], action = [{action_input}]")

        if action_input in ("remove from all groups", "remove_all"):
            action_input = "remove_all"
        elif action_input in ("add", "remove"):
            pass
        else:
            log.error("Action input must be 'add', 'remove', or 'remove from all groups'")
            log.result_message(ResultLevel.FAILED, "Action input must be 'add', 'remove', or 'remove from all groups'")
            return

        if not user_identifier:
            log.error("User identifier is empty or invalid")
            log.result_message(ResultLevel.FAILED, "User identifier is empty or invalid")
            return
        if action_input in ("add", "remove") and not user_groups:
            log.error("User groups value is empty or invalid")
            log.result_message(ResultLevel.FAILED, "User groups value is empty or invalid")
            return

        if provided_token:
            access_token = provided_token
            log.info("Using provided access token")
            if not isinstance(access_token, str) or "." not in access_token:
                log.result_message(ResultLevel.FAILED, "Provided access token is malformed (missing dots)")
                return
        elif ticket_number:
            company_identifier = get_company_identifier_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
            if not company_identifier:
                log.result_message(ResultLevel.FAILED, f"Failed to retrieve company identifier from ticket [{ticket_number}]")
                return

            if company_identifier == "MIT":
                if not validate_mit_authentication(log, http_client, vault_name, auth_code):
                    return

            client_id = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientID")
            client_secret = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientSecret")
            azure_domain = get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")

            if not all([client_id, client_secret, azure_domain]):
                log.result_message(ResultLevel.FAILED, "Failed to retrieve required secrets")
                return

            tenant_id = get_tenant_id_from_domain(log, http_client, azure_domain)
            if not tenant_id:
                log.result_message(ResultLevel.FAILED, "Failed to resolve tenant ID")
                return

            access_token = get_access_token(log, http_client, tenant_id, client_id, client_secret)
            if not isinstance(access_token, str) or "." not in access_token:
                log.result_message(ResultLevel.FAILED, "Access token is malformed (missing dots)")
                return
        else:
            log.result_message(ResultLevel.FAILED, "Either Access Token or Ticket Number is required")
            return

        exo_client_id = get_secret_value(log, http_client, vault_name, f"{company_identifier}-ExchangeApp-ClientID")
        exo_client_secret = get_secret_value(log, http_client, vault_name, f"{company_identifier}-ExchangeApp-ClientSecret")

        if not all([exo_client_id, exo_client_secret, azure_domain]):
            log.result_message(ResultLevel.FAILED, "Failed to retrieve required internal secrets")
            return

        exo_access_token = get_access_token(log, http_client, tenant_id, exo_client_id, exo_client_secret, scope="https://outlook.office365.com/.default")
        if not isinstance(exo_access_token, str) or "." not in exo_access_token:
            log.result_message(ResultLevel.FAILED, "Exchange Online access token is malformed (missing dots)")
            return

        user_result = get_user_data(log, http_client, msgraph_base_url, user_identifier, access_token)

        if isinstance(user_result, list):
            details = "\n".join([f"- {u.get('displayName')} | {u.get('userPrincipalName')} | {u.get('id')}" for u in user_result])
            log.result_message(ResultLevel.FAILED, f"Multiple users found for [{user_identifier}]\n{details}")
            return

        user_id, user_email, _ = user_result
        if not user_id:
            log.result_message(ResultLevel.FAILED, f"Failed to resolve user ID for [{user_identifier}]")
            return
        if not user_email:
            log.result_message(ResultLevel.FAILED, f"Unable to resolve user email for [{user_identifier}]")
            return

        if action_input == "remove_all":
            success, output = remove_user_from_all_groups(
                log, http_client, msgraph_base_url, azure_domain, user_id, user_email, exo_access_token, access_token
            )
            if success:
                lines = output.splitlines()
                removed_groups = [line for line in lines if line.startswith("Success: Removed")]
                failed_groups = [line for line in lines if "Error: Failed to remove" in line]

                if removed_groups:
                    log.info("Removed from groups:\n" + "\n".join(removed_groups))
                    log.result_message(ResultLevel.SUCCESS, f"EXO user [{user_email}] removed from {len(removed_groups)} groups successfully")
                elif failed_groups:
                    log.error("Some groups could not be removed:\n" + "\n".join(failed_groups))
                    log.result_message(ResultLevel.FAILED, f"Failed to remove EXO user [{user_email}] from one or more groups")
                else:
                    log.info(f"EXO user [{user_email}] had no removable group memberships")
                    log.result_message(ResultLevel.SUCCESS, f"EXO user [{user_email}] had no removable group memberships")
            else:
                log.error(f"PowerShell failed for remove_all operation for [{user_email}]")
                log.result_message(ResultLevel.FAILED, f"PowerShell failed for remove_all operation for [{user_email}]")

        elif action_input in ("add", "remove"):
            group_identifiers = [g.strip() for g in user_groups.split(",") if g.strip()]
            group_tuples = []
            for group_identifier in group_identifiers:
                result = get_group(log, http_client, msgraph_base_url, group_identifier, access_token)
                if isinstance(result, list):
                    details = "\n".join([f"- {g.get('displayName')} | {g.get('mail')} | {g.get('id')}" for g in result])
                    log.result_message(ResultLevel.FAILED, f"Multiple groups found for [{group_identifier}]\n{details}")
                    return
                group_email, group_id = result
                if not group_email:
                    log.result_message(ResultLevel.FAILED, f"Unable to resolve group email for [{group_identifier}]")
                    return
                group_tuples.append((group_email, group_id))

            if action_input == "add":
                success, output = add_user_to_groups(log, azure_domain, user_email, group_tuples, exo_access_token)
            else:
                success, output = remove_user_from_groups(log, azure_domain, user_email, group_tuples, exo_access_token)

            if success:
                lines = output.splitlines()
                success_lines = [line for line in lines if line.startswith("Success:")]
                fail_lines = [line for line in lines if line.startswith("Error:")]

                if success_lines:
                    log.info("Operation results:\n" + "\n".join(success_lines))
                    log.result_message(ResultLevel.SUCCESS, f"EXO user [{user_email}] {action_input}ed in {len(success_lines)} groups successfully")
                elif fail_lines:
                    log.error("Operation failed for some groups:\n" + "\n".join(fail_lines))
                    log.result_message(ResultLevel.FAILED, f"Failed to {action_input} EXO user [{user_email}] in one or more groups")
                else:
                    log.info(f"EXO user [{user_email}] was not {action_input}ed in any groups")
                    log.result_message(ResultLevel.SUCCESS, f"EXO user [{user_email}] was not {action_input}ed in any groups")
            else:
                log.error(f"PowerShell failed to {action_input} EXO user [{user_email}] to groups")
                log.result_message(ResultLevel.FAILED, f"PowerShell failed to {action_input} EXO user [{user_email}] to groups")

    except Exception:
        log.exception("An error occurred while processing")
        log.result_message(ResultLevel.FAILED, "Process failed")

if __name__ == "__main__":
    main()
