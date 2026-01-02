import sys
import random
import re
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

cwpsa_base_url = "https://aus.myconnectwise.net"
cwpsa_base_url_path = "/v4_6_release/apis/3.0"
msgraph_base_url_base = "https://graph.microsoft.com"
msgraph_base_url_path = "/v1.0"
msgraph_base_url_beta_base = "https://graph.microsoft.com"
msgraph_base_url_beta_path = "/beta"
vault_name = "PLACEHOLDER-akv1"
sender_email = "support@PLACEHOLDER.com.au"

data_to_log = {}
bot_name = "AAD - Group(s) Management"
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

def get_secret_value(log, http_client, vault_name, secret_name):
    log.info(f"Fetching secret [{secret_name}] from Key Vault [{vault_name}]")
    secret_url = (f"https://{vault_name}.vault.azure.net/secrets/{secret_name}?api-version=7.3")
    response = execute_api_call(
        log,
        http_client,
        "get",
        secret_url,
        integration_name="custom_wf_oauth2_client_creds",
    )
    if response:
        secret_value = response.json().get("value", "")
        if secret_value:
            log.info(f"Successfully retrieved secret [{secret_name}]")
            return secret_value
    return ""

def get_tenant_id_from_domain(log, http_client, azure_domain):
    try:
        config_url = (f"https://login.windows.net/{azure_domain}/.well-known/openid-configuration")
        log.info(f"Fetching OpenID configuration from [{config_url}]")
        response = execute_api_call(log, http_client, "get", config_url)
        if response:
            token_endpoint = response.json().get("token_endpoint", "")
            tenant_id = token_endpoint.split("/")[3] if token_endpoint else ""
            if tenant_id:
                log.info(f"Successfully extracted tenant ID [{tenant_id}]")
                return tenant_id
        return ""
    except Exception as e:
        log.exception(e, "Exception while extracting tenant ID from domain")
        return ""

def get_access_token(log, http_client, tenant_id, client_id, client_secret, scope="https://graph.microsoft.com/.default", log_prefix="Token"):
    log.info(f"[{log_prefix}] Requesting access token for scope [{scope}]")
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    payload = urllib.parse.urlencode(
        {
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": scope,
        }
    )
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

def get_company_data_from_ticket(log, http_client, cwpsa_base_url, cwpsa_base_url_path, ticket_number):
    log.info(f"Retrieving company details for ticket [{ticket_number}]")
    ticket_endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets/{ticket_number}"
    ticket_response = execute_api_call(log, http_client, "get", ticket_endpoint, integration_name="cw_psa")

    if ticket_response and ticket_response.status_code == 200:
        ticket_data = ticket_response.json()
        company = ticket_data.get("company", {})

        company_id = company["id"]
        company_identifier = company["identifier"]
        company_name = company["name"]

        log.info(f"Company ID: [{company_id}], Identifier: [{company_identifier}], Name: [{company_name}]")

        company_endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/company/companies/{company_id}"
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

def get_aad_user_data(log, http_client, msgraph_base_url_base, msgraph_base_url_path, user_identifier, token):
    log.info(f"Resolving user ID and email for [{user_identifier}]")
    headers = {"Authorization": f"Bearer {token}"}

    if re.fullmatch(
        r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}",
        user_identifier,
    ):
        endpoint = f"{msgraph_base_url_base}{msgraph_base_url_path}/users/{user_identifier}"
        response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
        if response:
            user = response.json()
            user_id = user.get("id", "")
            user_upn = user.get("userPrincipalName", "")
            user_sam = user.get("onPremisesSamAccountName", "")
            user_sync = user.get("onPremisesSyncEnabled", False)
            log.info(f"User found by GUID: ID=[{user_id}], UPN=[{user_upn}], SAM=[{user_sam}], OnPremSync=[{user_sync}]")
            return user_id, user_upn, user_sam, user_sync
        log.warning(f"No user found with GUID [{user_identifier}]")
        return "", "", "", False

    filters = [
        f"startswith(displayName,'{user_identifier}')",
        f"startswith(userPrincipalName,'{user_identifier}')",
        f"startswith(mail,'{user_identifier}')",
    ]
    filter_query = " or ".join(filters)
    endpoint = f"{msgraph_base_url_base}{msgraph_base_url_path}/users?$filter={urllib.parse.quote(filter_query)}"
    response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
    if response:
        users = response.json().get("value", [])
        log.info(f"User search returned {len(users)} result(s) for [{user_identifier}]")
        if len(users) > 1:
            log.error(f"Multiple users found for [{user_identifier}]")
            return users
        if users:
            user = users[0]
            user_id = user.get("id", "")
            user_upn = user.get("userPrincipalName", "")
            user_sam = user.get("onPremisesSamAccountName", "")
            user_sync = user.get("onPremisesSyncEnabled", False)
            log.info(f"User found: ID=[{user_id}], UPN=[{user_upn}], SAM=[{user_sam}], OnPremSync=[{user_sync}]")
            return user_id, user_upn, user_sam, user_sync
        else:
            log.warning(f"No users found for search term [{user_identifier}]")
    else:
        log.error(f"Failed to execute user search for [{user_identifier}]")
    return "", "", "", False

def get_aad_groups(log, http_client, msgraph_base_url, group_identifier, token):
    log.info(f"Resolving group name for [{group_identifier}]")
    headers = {"Authorization": f"Bearer {token}"}

    filters = [
        f"displayName eq '{group_identifier}'",
        f"mail eq '{group_identifier}'"
    ]
    filter_query = " or ".join(filters)
    endpoint = f"{msgraph_base_url_base}{msgraph_base_url_path}/groups?$filter={urllib.parse.quote(filter_query)}"

    response = execute_api_call(log, http_client, "get", endpoint, headers=headers)

    if response:
        groups = response.json().get("value", [])
        if len(groups) > 1:
            log.warning(f"Multiple groups found for [{group_identifier}]")
            return groups
        if groups:
            group = groups[0]
            group_name = group.get("displayName", group.get("mail", ""))
            group_id = group.get("id", "")
            group_email = group.get("mail", "")
            log.info(f"Group found for [{group_identifier}] - Name: {group_name}, ID: {group_id}")
            return group_name, group_id, group_email

    return "", "", ""

def get_user_aad_groups(log, http_client, msgraph_base_url, user_id, token):
    log.info(f"Fetching groups for user [{user_id}]")

    endpoint = f"{msgraph_base_url_base}{msgraph_base_url_path}/users/{user_id}/transitiveMemberOf"
    headers = {"Authorization": f"Bearer {token}"}
    all_groups = []
    dynamic_groups = []
    graph_groups = []
    skipped_mail_groups = []

    while endpoint:
        response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
        if not response:
            return [], [], []

        data = response.json()
        all_groups.extend(data.get('value', []))
        endpoint = data.get('@odata.nextLink')

    log.info(f"Total groups found for user [{user_id}]: {len(all_groups)}")

    for group in all_groups:
        group_id = group.get("id")
        group_name = group.get("displayName", "Unknown Group")
        group_types = group.get("groupTypes", [])
        mail_enabled = group.get("mailEnabled", False)

        if "DynamicMembership" in group_types:
            dynamic_groups.append((group_name, group_id))
        elif "Unified" in group_types:
            graph_groups.append((group_name, group_id))
        elif mail_enabled:
            skipped_mail_groups.append((group_name, group_id))
        else:
            graph_groups.append((group_name, group_id))

    if skipped_mail_groups:
        log.info(f"Mail-enabled (non-M365) groups found for user [{user_id}]: {len(skipped_mail_groups)}")

    return dynamic_groups, graph_groups, skipped_mail_groups

def add_user_to_aad_groups(log, http_client, msgraph_base_url, group_id, user_id, access, token):
    log.info(f"Adding user [{user_id}] to group [{group_id}] as [{access}]")
    if access not in ["Owner", "Member"]:
        log.warning(f"Invalid access role [{access}] for group [{group_id}]")
        return False

    ref_type = "owners" if access == "Owner" else "members"
    endpoint = f"{msgraph_base_url_base}{msgraph_base_url_path}/groups/{group_id}/{ref_type}/$ref"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    data = {
        "@odata.id": f"https://graph.microsoft.com/v1.0/directoryObjects/{user_id}"
    }

    response = execute_api_call(log, http_client, "post", endpoint, data=data, headers=headers)
    if response:
        log.info(f"Successfully added user [{user_id}] to group [{group_id}] as [{access}]")
        return True
    return False

def remove_user_from_aad_groups(log, http_client, msgraph_base_url, group_id, user_id, access, token):
    log.info(f"Removing user [{user_id}] from group [{group_id}] as [{access}]")
    if access not in ["Owner", "Member"]:
        log.warning(f"Invalid access role [{access}] for group [{group_id}]")
        return False

    ref_type = "owners" if access == "Owner" else "members"
    endpoint = f"{msgraph_base_url_base}{msgraph_base_url_path}/groups/{group_id}/{ref_type}/{user_id}/$ref"
    headers = {
        "Authorization": f"Bearer {token}"
    }

    response = execute_api_call(log, http_client, "delete", endpoint, headers=headers)
    if response:
        if response.status_code == 404:
            log.info(f"User [{user_id}] not found in group [{group_id}] as [{access}] - skipping (already removed)")
        else:
            log.info(f"Successfully removed user [{user_id}] from group [{group_id}] as [{access}]")
        return True
    return False

def remove_user_from_all_aad_groups(log, http_client, msgraph_base_url, user_id, user_identifier, token):
    dynamic_groups, graph_groups, skipped_mail_groups = get_user_aad_groups(log, http_client, msgraph_base_url, user_id, token)
    removed_groups = []
    skipped_groups = []
    failed_groups = []

    for group_name, group_id in dynamic_groups:
        log.warning(f"Skipping dynamic group [{group_name}] - [{group_id}] (cannot remove manually)")

    for group_name, group_id in skipped_mail_groups:
        log.warning(f"Skipping mail-enabled/distribution group [{group_name}] - [{group_id}] (not removable via Graph API)")

    for group_name, group_id in graph_groups:
        for access in ["Member", "Owner"]:
            ref_type = "members" if access == "Member" else "owners"
            log.info(f"Attempting to remove user [{user_identifier}] from group [{group_name}] - [{group_id}] as {access} using Graph API")
            endpoint = f"{msgraph_base_url_base}{msgraph_base_url_path}/groups/{group_id}/{ref_type}/{user_id}/$ref"
            headers = {"Authorization": f"Bearer {token}"}
            response = execute_api_call(log, http_client, "delete", endpoint, headers=headers)

            if response:
                log.info(f"User [{user_identifier}] removed from [{group_name}] - [{group_id}] as {access} successfully via Graph API")
                removed_groups.append(f"{group_name}:{access}")
            elif response and response.status_code == 403 and "Authorization_RequestDenied" in response.text:
                log.warning(f"Permission denied removing [{user_identifier}] as {access} from group [{group_name}] - [{group_id}]")
                failed_groups.append((group_name, group_id, access))
            elif not response or response.status_code == 404:
                log.warning(f"Skipping non-existent group [{group_name}] - [{group_id}] for {access}")
                skipped_groups.append((group_name, group_id, access))
            else:
                
                failed_groups.append((group_name, group_id, access))

    return removed_groups, skipped_groups, failed_groups

def check_group_on_premises_sync(log, http_client, msgraph_base_url, group_id, token):
    log.info(f"Checking group on-premises sync for group [{group_id}]")
    endpoint = f"{msgraph_base_url_base}{msgraph_base_url_path}/groups/{group_id}"
    headers = {"Authorization": f"Bearer {token}"}
    response = execute_api_call(log, http_client, "get", endpoint, headers=headers)

    if response:
        on_premises_sync_enabled = response.json().get("onPremisesSyncEnabled", False)
        if on_premises_sync_enabled == "true":
            on_premises_sync_enabled = True
        elif on_premises_sync_enabled == "false":
            on_premises_sync_enabled = False
        log.info(f"Group on-premises sync for group [{group_id}]: {on_premises_sync_enabled}")
        return on_premises_sync_enabled
    else:
        log.warning(f"Failed to retrieve group details for [{group_id}]")
        return False

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1755223092157")
            operation = input.get_value("Operation_1755223095419")
            user_identifier = input.get_value("User_1755223102548")
            groups_raw = input.get_value("Groups_1755223098735")
            access = input.get_value("Access_1755223106795")
            auth_code = input.get_value("AuthenticationCode_1755223110321")
            graph_token = input.get_value("GraphToken_1755223137658")

        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        operation = operation.strip() if operation else ""
        user_identifier = user_identifier.strip() if user_identifier else ""
        groups_raw = groups_raw.strip() if groups_raw else ""
        access = access.strip() if access else ""
        auth_code = auth_code.strip() if auth_code else ""
        graph_token = graph_token.strip() if graph_token else ""

        log.info(f"Input validation - Ticket: [{ticket_number}], Operation: [{operation}], User: [{user_identifier}], Groups: [{groups_raw[:50] if groups_raw else 'None'}{'...' if groups_raw and len(groups_raw) > 50 else ''}], Access: [{access}]")

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number is required but missing")
            return
        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation value is missing or invalid")
            return
        if operation in ("Add user to AAD groups", "Remove user from AAD groups"):
            if not groups_raw:
                record_result(log, ResultLevel.WARNING, f"Groups input is required for operation [{operation}]")
                return
            if not access:
                record_result(log, ResultLevel.WARNING, f"Access level input is required for operation [{operation}]")
                return
        if operation in ("Get AAD groups", "Check Group On-Premises Sync") and not groups_raw:
            record_result(log, ResultLevel.WARNING, f"Groups input is required for operation [{operation}]")
            return
        if operation not in ["Get AAD groups", "Get user AAD groups", "Add user to AAD groups", "Remove user from AAD groups", "Remove user from all AAD groups", "Check Group On-Premises Sync"]:
            record_result(log, ResultLevel.WARNING, f"Invalid operation [{operation}] specified")
            return

        company_identifier, company_name, company_id, company_type = get_company_data_from_ticket(log, http_client, cwpsa_base_url, cwpsa_base_url_path, ticket_number)
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

        azure_domain = get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")
        if not azure_domain:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve Azure domain for [{company_identifier}]")
            return

        aad_user_result = get_aad_user_data(log, http_client, msgraph_base_url_base, msgraph_base_url_path, user_identifier, graph_access_token)
        if isinstance(aad_user_result, list):
            details = "\n".join([f"- {u.get('displayName')} | {u.get('userPrincipalName')} | {u.get('id')}" for u in aad_user_result])
            record_result(log, ResultLevel.WARNING, f"Multiple users found for [{user_identifier}]\n{details}")
            return
        user_id, user_email, user_sam, user_onpremisessyncenabled = aad_user_result

        if not user_id:
            record_result(log, ResultLevel.WARNING, f"Failed to resolve user ID for [{user_identifier}]")
            return
        if not user_email:
            record_result(log, ResultLevel.WARNING, f"Unable to resolve user principal name for [{user_identifier}]")
            return

        log.info(f"Successfully resolved user: ID=[{user_id}], Email=[{user_email}], SAM=[{user_sam}], OnPremSync=[{user_onpremisessyncenabled}]")

        log.info(f"Operation: {operation}")
        successes = []
        failures = []

        if operation == "Get AAD groups":
            all_group_names = []
            for group in [g.strip() for g in groups_raw.split(",") if g.strip()]:
                group_result = get_aad_groups(log, http_client, msgraph_base_url, group, graph_access_token)
                if type(group_result) == list:
                    record_result(log, ResultLevel.SUCCESS, f"Multiple groups found for [{group}]")
                    for g in group_result:
                        # Add mail to result if it exists, otherwise set to "Not Mail Enabled"
                        mail = "Not Mail Enabled" if g.get('mail') == None else g.get('mail')
                        details = f"{g.get('displayName')} | {mail} | {g.get('id')}"
                        record_result(log, ResultLevel.SUCCESS, details)
                resolved_name, group_id, group_email = group_result
                if group_id:
                    all_group_names.append(resolved_name)

            if all_group_names:
                data_to_log["Success"] = ", ".join(all_group_names)
                record_result(log, ResultLevel.SUCCESS, f"Successfully resolved {len(all_group_names)} group(s):")
                for group in all_group_names:
                    record_result(log, ResultLevel.SUCCESS, f"- {group}")
            else:
                record_result(log, ResultLevel.WARNING, "Failed to resolve any groups")

        elif operation == "Get user AAD groups":
            dynamic_groups, graph_groups, skipped_mail_groups = get_user_aad_groups(log, http_client, msgraph_base_url, user_id, graph_access_token)
            all_group_names = [g[0] for g in graph_groups + skipped_mail_groups + dynamic_groups]

            if all_group_names:
                data_to_log["Success"] = ", ".join(all_group_names)
                record_result(log, ResultLevel.SUCCESS, f"User [{user_email}] is a member of {len(all_group_names)} AAD groups:")
                for group in all_group_names:
                    record_result(log, ResultLevel.SUCCESS, f"- {group}")
            else:
                record_result(log, ResultLevel.WARNING, f"User [{user_email}] is not a member of any AAD groups")

        elif operation == "Add user to AAD groups":
            groups_list = [g.strip() for g in groups_raw.split(",") if g.strip()]
            log.info(f"Processing {len(groups_list)} group(s) for addition: {groups_list}")
            
            for group in groups_list:
                log.info(f"Resolving group [{group}]")
                group_result = get_aad_groups(log, http_client, msgraph_base_url, group, graph_access_token)
                if isinstance(group_result, list):
                    details = "\n".join([f"- {g.get('displayName')} | {g.get('mail')} | {g.get('id')}" for g in group_result])
                    record_result(log, ResultLevel.WARNING, f"Multiple groups found for [{group}]\n{details}")
                    failures.append(f"{group}:{access}")
                    continue
                resolved_name, group_id, group_email = group_result
                if not group_id:
                    log.warning(f"Failed to resolve group [{group}] - group not found")
                    failures.append(f"{group}:{access}")
                    continue
                
                log.info(f"Adding user [{user_email}] to group [{resolved_name}] with access [{access}]")
                success = add_user_to_aad_groups(log, http_client, msgraph_base_url, group_id, user_id, access, graph_access_token)
                if success:
                    log.info(f"Successfully added user to group [{resolved_name}]")
                    group_display = f"{resolved_name} - {group_email}" if group_email else resolved_name
                    successes.append(f"[{user_email}] added to [{group_display}] as [{access}]")
                else:
                    log.warning(f"Failed to add user to group [{resolved_name}]")
                    failures.append(f"{resolved_name}:{access}")

            log.info(f"Add operation completed - Successes: {len(successes)}, Failures: {len(failures)}")
            if successes:
                data_to_log["Success"] = ", ".join(successes)
                for entry in successes:
                    record_result(log, ResultLevel.SUCCESS, entry)
            if failures:
                data_to_log["Failed"] = ", ".join([f"Add AAD group permission - {f}" for f in failures])
                for entry in failures:
                    record_result(log, ResultLevel.WARNING, f"Add AAD group permission - {entry}")
            if not successes and not failures:
                record_result(log, ResultLevel.WARNING, "No groups were processed for the add operation")

        elif operation == "Remove user from AAD groups":
            for group in [g.strip() for g in groups_raw.split(",") if g.strip()]:
                group_result = get_aad_groups(log, http_client, msgraph_base_url, group, graph_access_token)
                if isinstance(group_result, list):
                    record_result(log, ResultLevel.WARNING, f"Multiple groups found for [{group}]")
                    for g in group_result:
                        if not group_id:
                            failures.append(f"{g.get('displayName')}:{access}")
                            continue
                        success = add_user_to_aad_groups(log, http_client, msgraph_base_url, group_id, user_id, access, graph_access_token)
                        if success:
                            successes.append(f"{g.get('displayName')}:{access}")
                        else:
                            failures.append(f"{g.get('displayName')}:{access}") 
                    
                resolved_name, group_id, group_email = group_result
                if not group_id:
                    failures.append(f"{group}:{access}")
                    continue
                
                for access_role in ["Member", "Owner"]:
                    success = remove_user_from_aad_groups(log, http_client, msgraph_base_url, group_id, user_id, access_role, graph_access_token)
                    if success:
                        group_display = f"{resolved_name} - {group_email}" if group_email else resolved_name
                        successes.append(f"[{user_email}] removed from [{group_display}] as [{access_role}]")
                    else:
                        failures.append(f"{resolved_name}:{access_role}")

            if successes:
                data_to_log["Success"] = ", ".join(successes)
                for entry in successes:
                    record_result(log, ResultLevel.SUCCESS, entry)
            if failures:
                data_to_log["Failed"] = ", ".join([f"Remove AAD group permission - {f}" for f in failures])
                for entry in failures:
                    record_result(log, ResultLevel.WARNING, f"Remove AAD group permission - {entry}")

        elif operation == "Remove user from all AAD groups":
            removed, skipped, failed = remove_user_from_all_aad_groups(
                log, http_client, msgraph_base_url, user_id, user_identifier, graph_access_token
            )

            cleaned_removed = [g for g in removed if g.strip()]
            if cleaned_removed:
                data_to_log["Success"] = ", ".join(cleaned_removed)
                record_result(log, ResultLevel.SUCCESS, f"Removed [{user_email}] from AAD group(s):")
                for entry in cleaned_removed:
                    group, access = entry.rsplit(":", 1)
                    record_result(log, ResultLevel.SUCCESS, f"- {group} ({access})")

            for group_name, group_id, access in skipped:
                record_result(
                    log,
                    ResultLevel.SUCCESS,
                    f"Skipped removing [{user_email}] from [{group_name} - {group_id}] as {access} (already removed or not found)"
                )

            if failed:
                failed_entries = [f"{group_name}:{access}" for group_name, group_id, access in failed]
                data_to_log["Failed"] = ", ".join(
                    [f"Remove AAD group permission - {entry}" for entry in failed_entries]
                )
                for group_name, group_id, access in failed:
                    record_result(
                        log,
                        ResultLevel.WARNING,
                        f"Failed to remove [{user_email}] as {access} from [{group_name} - {group_id}]"
                    )

            if not cleaned_removed and not failed:
                record_result(log, ResultLevel.SUCCESS, f"No AAD groups required removal for [{user_email}]")

        elif operation == "Check Group On-Premises Sync":
            for group in [g.strip() for g in groups_raw.split(",") if g.strip()]:
                group_result = get_aad_groups(log, http_client, msgraph_base_url, group, graph_access_token)
                if isinstance(group_result, list):
                    details = "\n".join([f"- {g.get('displayName')} | {g.get('mail')} | {g.get('id')}" for g in group_result])
                    record_result(log, ResultLevel.WARNING, f"Multiple groups found for [{group}]\n{details}")
                    continue
                resolved_name, group_id, group_email = group_result
                if not group_id:
                    failures.append(f"{group}:{access}")
                    continue
                on_premises_sync_enabled = check_group_on_premises_sync(log, http_client, msgraph_base_url, group_id, graph_access_token)
                if on_premises_sync_enabled == True:
                    record_result(log, ResultLevel.SUCCESS, f"Group [{group}] is on-premises synced")
                    data_to_log["OnPremisesSyncEnabled"] = "Yes"
                else:
                    record_result(log, ResultLevel.WARNING, f"Group [{group}] is not on-premises synced")
                    data_to_log["OnPremisesSyncEnabled"] = "No"

        else:
            record_result(log, ResultLevel.WARNING, f"Unknown operation [{operation}]")
            return

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
