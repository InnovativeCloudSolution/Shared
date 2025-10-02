import sys
import random
import re
import os
import time
import urllib.parse
import requests
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(
    os.path.join(os.path.dirname(__file__), '..')))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

cwpsa_base_url = "https://au.myconnectwise.net/v4_6_release/apis/3.0"
msgraph_base_url = "https://graph.microsoft.com/v1.0"
msgraph_base_url_beta = "https://graph.microsoft.com/beta"
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
                response = getattr(http_client.third_party_integration(integration_name), method)(
                    url=endpoint, json=data) if data else getattr(http_client.third_party_integration(integration_name), method)(url=endpoint)
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
                wait_time = int(retry_after) if retry_after else base_delay * \
                    (2 ** attempt) + random.uniform(0, 3)
                log.warning(
                    f"Rate limit exceeded. Retrying in {wait_time:.2f} seconds")
                time.sleep(wait_time)
            elif response.status_code == 404:
                log.warning(f"Skipping non-existent resource [{endpoint}]")
                return None
            else:
                log.error(
                    f"API request failed Status: {response.status_code}, Response: {response.text}")
                return response
        except Exception as e:
            log.exception(e, f"Exception during API call to {endpoint}")
            return None
    return None


def get_secret_value(log, http_client, vault_name, secret_name):
    log.info(f"Fetching secret [{secret_name}] from Key Vault [{vault_name}]")

    secret_url = f"https://{vault_name}.vault.azure.net/secrets/{secret_name}?api-version=7.3"
    response = execute_api_call(
        log, http_client, "get", secret_url, integration_name="custom_wf_oauth2_client_creds")

    if response and response.status_code == 200:
        secret_value = response.json().get("value", "")
        if secret_value:
            log.info(f"Successfully retrieved secret [{secret_name}]")
            return secret_value

    log.error(
        f"Failed to retrieve secret [{secret_name}] Status code: {response.status_code if response else 'N/A'}")
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

    response = execute_api_call(
        log, http_client, "post", token_url, data=payload, retries=3, headers=headers)

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

    log.error(
        f"Failed to retrieve access token Status code: {response.status_code if response else 'N/A'}")
    return ""


def get_company_identifier_from_ticket(log, http_client, cwpsa_base_url, ticket_number):
    log.info(f"Retrieving company identifier for ticket [{ticket_number}]")
    endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}"

    response = execute_api_call(
        log, http_client, "get", endpoint, integration_name="cw_psa")

    if response:
        if response.status_code == 200:
            data = response.json()
            company_identifier = data.get("company", {}).get("identifier", "")
            if company_identifier:
                log.info(
                    f"Company identifier for ticket [{ticket_number}] is [{company_identifier}]")
                return company_identifier
            else:
                log.error(
                    f"Company identifier not found in response for ticket [{ticket_number}]")
        else:
            log.error(
                f"Failed to retrieve company identifier for ticket [{ticket_number}] Status: {response.status_code}, Body: {response.text}")
    else:
        log.error(
            f"Failed to retrieve company identifier for ticket [{ticket_number}]: No response received")

    return ""


def validate_mit_authentication(log, http_client, vault_name, auth_code):
    if not auth_code:
        log.result_message(ResultLevel.FAILED,
                           "Authentication code input is required for MIT")
        return False

    expected_code = get_secret_value(
        log, http_client, vault_name, "MIT-AuthenticationCode")
    if not expected_code:
        log.result_message(
            ResultLevel.FAILED, "Failed to retrieve expected authentication code for MIT")
        return False

    if auth_code.strip() != expected_code.strip():
        log.result_message(ResultLevel.FAILED,
                           "Provided authentication code is incorrect")
        return False

    return True


def get_aad_user(log, http_client, msgraph_base_url, user_identifier, token):
    log.info(f"Resolving user ID and email for [{user_identifier}]")
    headers = {"Authorization": f"Bearer {token}"}

    if re.fullmatch(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}", user_identifier):
        endpoint = f"{msgraph_base_url}/users/{user_identifier}"
        response = execute_api_call(
            log, http_client, "get", endpoint, headers=headers)
        if response and response.status_code == 200:
            user = response.json()
            user_id = user.get("id")
            user_email = user.get("userPrincipalName", "")
            user_sam = user.get("onPremisesSamAccountName", "")
            log.info(
                f"User found by object ID [{user_id}] with email [{user_email}] and SAM [{user_sam}]")
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
    response = execute_api_call(
        log, http_client, "get", endpoint, headers=headers)

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
            log.info(
                f"User found for [{user_identifier}] - ID: {user_id}, Email: {user_email}, SAM: {user_sam}")
            return user_id, user_email, user_sam

    log.error(f"Failed to resolve user ID and email for [{user_identifier}]")
    return "", "", ""


def get_aad_groups(log, http_client, msgraph_base_url, group_identifiers, token):
    resolved_groups = []
    headers = {"Authorization": f"Bearer {token}"}
    identifiers = [g.strip()
                   for g in group_identifiers.split(",") if g.strip()]

    for group_identifier in identifiers:
        log.info(f"Resolving group name for [{group_identifier}]")

        filters = [
            f"startswith(displayName,'{group_identifier}')",
            f"startswith(mail,'{group_identifier}')"
        ]
        filter_query = " or ".join(filters)

        encoded_filter_query = urllib.parse.quote(filter_query, safe="")

        endpoint = f"{msgraph_base_url}/groups?$filter={encoded_filter_query}"

        response = execute_api_call(
            log, http_client, "get", endpoint, headers=headers)

        if response and response.status_code == 200:
            groups = response.json().get("value", [])
            if len(groups) > 1:
                log.warning(f"Multiple groups found for [{group_identifier}]")
                resolved_groups.append(groups)
                continue
            if groups:
                group = groups[0]
                group_name = group.get("displayName", group.get("mail", ""))
                group_id = group.get("id", "")
                log.info(
                    f"Group found for [{group_identifier}] - Name: {group_name}, ID: {group_id}")
                resolved_groups.append((group_name, group_id))
                continue

        log.warning(f"Failed to resolve group for [{group_identifier}]")
        resolved_groups.append(("", ""))

    return resolved_groups


def get_user_aad_groups(log, http_client, msgraph_base_url, user_id, target_groups_csv, token):
    log.info(f"Fetching groups for user [{user_id}]")

    endpoint = f"{msgraph_base_url}/users/{user_id}/transitiveMemberOf"
    headers = {"Authorization": f"Bearer {token}"}
    all_groups = []
    dynamic_groups = []
    graph_groups = []
    skipped_mail_groups = []

    target_groups = [g.strip().lower()
                     for g in target_groups_csv.split(",") if g.strip()]

    while endpoint:
        response = execute_api_call(
            log, http_client, "get", endpoint, headers=headers)
        if not response or response.status_code != 200:
            log.error(
                f"Failed to retrieve groups for [{user_id}] Status code: {response.status_code if response else 'N/A'}")
            return [], [], []

        data = response.json()
        all_groups.extend(data.get("value", []))
        endpoint = data.get("@odata.nextLink")

    log.info(f"Total groups found for user [{user_id}]: {len(all_groups)}")

    for group in all_groups:
        group_id = group.get("id")
        group_name = group.get("displayName", "Unknown Group")
        group_types = group.get("groupTypes", [])
        mail_enabled = group.get("mailEnabled", False)

        if group_name.lower() not in target_groups:
            continue

        if "DynamicMembership" in group_types:
            dynamic_groups.append((group_name, group_id))
        elif "Unified" in group_types:
            graph_groups.append((group_name, group_id))
        elif mail_enabled:
            skipped_mail_groups.append((group_name, group_id))
        else:
            graph_groups.append((group_name, group_id))

    return dynamic_groups, graph_groups, skipped_mail_groups


def add_user_to_aad_groups(log, http_client, msgraph_base_url, group_ids_csv, user_id, access, token):
    if access not in ["Owner", "Member"]:
        log.warning(f"Invalid access role [{access}] for user [{user_id}]")
        return []

    group_ids = [g.strip() for g in group_ids_csv.split(",") if g.strip()]
    ref_type = "owners" if access == "Owner" else "members"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    data = {
        "@odata.id": f"https://graph.microsoft.com/v1.0/directoryObjects/{user_id}"
    }

    successes = []
    failures = []

    for group_id in group_ids:
        log.info(
            f"Adding user [{user_id}] to group [{group_id}] as [{access}]")
        endpoint = f"{msgraph_base_url}/groups/{group_id}/{ref_type}/$ref"

        response = execute_api_call(
            log, http_client, "post", endpoint, data=data, headers=headers)
        if response and response.status_code in [200, 204]:
            log.info(
                f"Successfully added user [{user_id}] to group [{group_id}] as [{access}]")
            successes.append(group_id)
        else:
            log.error(
                f"Failed to add user [{user_id}] to group [{group_id}] as [{access}] Status: {response.status_code if response else 'N/A'}")
            failures.append(group_id)

    return successes, failures


def remove_user_from_groups(log, http_client, msgraph_base_url, group_ids_csv, user_id, access, token):
    if access not in ["Owner", "Member"]:
        log.warning(f"Invalid access role [{access}] for user [{user_id}]")
        return []

    group_ids = [g.strip() for g in group_ids_csv.split(",") if g.strip()]
    ref_type = "owners" if access == "Owner" else "members"
    headers = {"Authorization": f"Bearer {token}"}

    successes = []
    failures = []

    for group_id in group_ids:
        log.info(
            f"Removing user [{user_id}] from group [{group_id}] as [{access}]")
        endpoint = f"{msgraph_base_url}/groups/{group_id}/{ref_type}/{user_id}/$ref"

        response = execute_api_call(
            log, http_client, "delete", endpoint, headers=headers)
        if response and response.status_code in [200, 204]:
            log.info(
                f"Successfully removed user [{user_id}] from group [{group_id}] as [{access}]")
            successes.append(group_id)
        else:
            log.error(
                f"Failed to remove user [{user_id}] from group [{group_id}] as [{access}] Status: {response.status_code if response else 'N/A'}")
            failures.append(group_id)

    return successes, failures


def remove_user_from_all_groups(log, http_client, msgraph_base_url, user_id, user_identifier, token):
    dynamic_groups, graph_groups, skipped_mail_groups = get_user_aad_groups(
        log, http_client, msgraph_base_url, user_id, token
    )
    removed_groups = []
    skipped_groups = []

    for group_name, group_id in dynamic_groups:
        log.warning(
            f"Skipping dynamic group [{group_name}] - [{group_id}] (cannot remove manually)")

    for group_name, group_id in skipped_mail_groups:
        log.warning(
            f"Skipping mail-enabled/distribution group [{group_name}] - [{group_id}] (not removable via Graph API)")

    for group_name, group_id in graph_groups:
        for role_type in ["members", "owners"]:
            log.info(
                f"Attempting to remove user [{user_identifier}] from [{group_name}] - [{group_id}] as [{role_type[:-1].capitalize()}]")
            endpoint = f"{msgraph_base_url}/groups/{group_id}/{role_type}/{user_id}/$ref"
            headers = {"Authorization": f"Bearer {token}"}

            response = execute_api_call(
                log, http_client, "delete", endpoint, headers=headers)
            if response and response.status_code == 204:
                log.info(
                    f"User [{user_identifier}] removed from [{group_name}] - [{group_id}] as [{role_type[:-1].capitalize()}] successfully")
                removed_groups.append(
                    f"{group_name}:{role_type[:-1].capitalize()}")
            elif response and response.status_code == 403 and "Authorization_RequestDenied" in response.text:
                log.warning(
                    f"Permission denied for group [{group_name}] - [{group_id}] as [{role_type[:-1].capitalize()}]")
                skipped_groups.append(
                    (group_name, group_id, role_type[:-1].capitalize()))
            elif response and response.status_code == 404:
                log.warning(
                    f"Skipping non-existent group [{group_name}] - [{group_id}] as [{role_type[:-1].capitalize()}]")
            else:
                log.error(
                    f"Graph removal failed for group [{group_name}] - [{group_id}] as [{role_type[:-1].capitalize()}] Status: {response.status_code if response else 'N/A'}")
                skipped_groups.append(
                    (group_name, group_id, role_type[:-1].capitalize()))

    return removed_groups, skipped_groups


def clean_append(results, target, failures, group_type, user_identifier):
    for g in results:
        if not isinstance(g, (tuple, list)) or len(g) < 2:
            failures.append(
                f"Malformed {group_type} group result for [{user_identifier}]: {g}")
            continue
        name, gid = g[0], g[1]
        if name and gid:
            target.append({"Name": name, "ID": gid})
        else:
            failures.append(
                f"Missing name or ID in {group_type} group membership for [{user_identifier}]: {g}")


def main():
    try:
        log.info("Starting main execution")

        try:
            operation = input.get_value("Operation_1749769645433")
            user_identifier = input.get_value("User_1749769646703")
            group_names = input.get_value("Groups_1749839230016")
            access = input.get_value("Access_1749838680700")
            ticket_number = input.get_value("TicketNumber_1749769640738")
            auth_code = input.get_value("AuthCode_1749769790141")
            provided_token = input.get_value("AccessToken_1749769800392")
        except Exception:
            record_result(log, ResultLevel.WARNING,
                          "Failed to fetch input values")
            return

        log.info("Successfully fetched all input values")

        operation = operation.strip() if operation else ""
        user_identifier = user_identifier.strip() if user_identifier else ""
        group_names = group_names.strip() if group_names else ""
        access = access.strip() if access else ""
        ticket_number = ticket_number.strip() if ticket_number else ""
        auth_code = auth_code.strip() if auth_code else ""
        provided_token = provided_token.strip() if provided_token else ""

        log.info(
            f"Inputs sanitized: operation=[{operation}], user_identifier=[{user_identifier}], groups=[{group_names}], access=[{access}], ticket=[{ticket_number}]")

        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation is required")
            return

        if provided_token and "." in provided_token:
            access_token = provided_token
            log.info("Using provided access token")
        elif ticket_number:
            log.info(f"Authenticating using ticket [{ticket_number}]")
            company_identifier = get_company_identifier_from_ticket(
                log, http_client, cwpsa_base_url, ticket_number)
            if not company_identifier:
                record_result(log, ResultLevel.WARNING,
                              f"Failed to retrieve company identifier from ticket [{ticket_number}]")
                return

            log.info(f"Company identifier retrieved: [{company_identifier}]")

            if company_identifier == "MIT":
                log.info("Validating MIT authentication")
                if not validate_mit_authentication(log, http_client, vault_name, auth_code):
                    return

            log.info("Fetching secrets for authentication")
            client_id = get_secret_value(
                log, http_client, vault_name, "MIT-PartnerApp-ClientID")
            client_secret = get_secret_value(
                log, http_client, vault_name, "MIT-PartnerApp-ClientSecret")
            azure_domain = get_secret_value(
                log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")

            if not all([client_id, client_secret, azure_domain]):
                record_result(log, ResultLevel.WARNING,
                              "Failed to retrieve required secrets for Azure authentication")
                return

            log.info("Retrieved client_id, secret, and domain successfully")

            tenant_id = get_tenant_id_from_domain(
                log, http_client, azure_domain)
            if not tenant_id:
                record_result(log, ResultLevel.WARNING,
                              "Failed to resolve tenant ID from Azure domain")
                return

            log.info(f"Tenant ID resolved: [{tenant_id}]")

            access_token = get_access_token(
                log, http_client, tenant_id, client_id, client_secret)
            if not access_token or "." not in access_token:
                record_result(log, ResultLevel.WARNING,
                              "Access token is malformed (missing dots)")
                return
        else:
            record_result(log, ResultLevel.WARNING,
                          "Either Access Token or Ticket Number is required")
            return

        log.info("Proceeding to resolve groups")
        resolved_groups = get_aad_groups(
            log, http_client, msgraph_base_url, group_names, access_token)

        log.info("Proceeding to resolve user")
        user_id, _, _ = get_aad_user(
            log, http_client, msgraph_base_url, user_identifier, access_token)
        if not user_id and operation != "Get AAD groups":
            record_result(log, ResultLevel.WARNING,
                          f"User not found for [{user_identifier}]")
            return

        log.info(f"User ID resolved: [{user_id}]")
        log.info("Fetching existing group memberships")
        dyn, graph, mail = get_user_aad_groups(
            log, http_client, msgraph_base_url, user_id, group_names, access_token) if user_id else ([], [], [])
        existing_group_ids = {g[1] for g in dyn + graph + mail}

        log.info("Starting operation block")

        if operation == "Get AAD groups":
            log.info("Operation: Get AAD groups")
            dynamic, graph_out, mail_out = [], [], []
            input_group_list = [g.strip().split(":")[0] for g in group_names.split(",") if g.strip()]
            successes = []
            failures = []

            for i, group_identifier in enumerate(input_group_list):
                log.info(f"Evaluating input group [{group_identifier}] (index {i})")
                if i >= len(resolved_groups):
                    failures.append(f"No result returned for group '{group_identifier}' (index {i})")
                    continue

                entry = resolved_groups[i]
                if entry is None:
                    failures.append(f"Failed to resolve AAD group - [{group_identifier}] (Null entry)")
                    continue

                if isinstance(entry, tuple):
                    group_name, group_id = entry
                    if not group_name or not group_id:
                        failures.append(f"Failed to resolve AAD group - [{group_identifier}] (Missing name or ID)")
                        continue
                    graph_out.append({"Name": group_name, "ID": group_id})
                    successes.append(f"Found AAD group - [{group_name}] - [{group_id}]")
                elif isinstance(entry, list):
                    if not entry:
                        failures.append(f"Failed to resolve AAD group - [{group_identifier}] (Empty list)")
                        continue
                    for g in entry:
                        if not g:
                            failures.append(f"Null entry in results for [{group_identifier}]")
                            continue
                        group_name = g.get("displayName", g.get("mail", ""))
                        group_id = g.get("id", "")
                        group_types = g.get("groupTypes", [])
                        is_dynamic = "DynamicMembership" in group_types
                        is_unified = "Unified" in group_types
                        is_mail = g.get("mailEnabled", False)
                        if not group_name or not group_id:
                            failures.append(f"Failed to resolve AAD group - [{group_identifier}] (Missing name or ID in multi-match)")
                            continue
                        if is_dynamic:
                            dynamic.append({"Name": group_name, "ID": group_id})
                        elif is_mail and not is_unified:
                            mail_out.append({"Name": group_name, "ID": group_id})
                        else:
                            graph_out.append({"Name": group_name, "ID": group_id})
                        successes.append(f"Found AAD group - [{group_name}] - [{group_id}]")
                else:
                    failures.append(f"Failed to resolve AAD group - [{group_identifier}] (Unknown entry type)")

            data_to_log["Dynamic"] = dynamic
            data_to_log["Graph"] = graph_out
            data_to_log["Mail"] = mail_out

            log.info(f"Group resolution summary: Successes={len(successes)}, Failures={len(failures)}")

            for msg in successes:
                record_result(log, ResultLevel.SUCCESS, msg)
            for msg in failures:
                record_result(log, ResultLevel.WARNING, msg)

            record_result(log, ResultLevel.WARNING if failures else ResultLevel.SUCCESS, "Completed AAD group resolution")

        elif operation == "Get user AAD groups":
            log.info("Operation: Get user AAD groups")
            dyn_clean = []
            graph_clean = []
            mail_clean = []
            failures = []
            successes = []

            clean_append(dyn, dyn_clean, failures, "Dynamic", user_identifier)
            clean_append(graph, graph_clean, failures, "Graph", user_identifier)
            clean_append(mail, mail_clean, failures, "Mail", user_identifier)

            data_to_log["Dynamic"] = dyn_clean
            data_to_log["Graph"] = graph_clean
            data_to_log["Mail"] = mail_clean

            log.info(f"Cleaned group results - Dynamic={len(dyn_clean)}, Graph={len(graph_clean)}, Mail={len(mail_clean)}")

            for g in dyn_clean + graph_clean + mail_clean:
                successes.append(f"Found AAD group membership for [{user_identifier}] - [{g['Name']}] - [{g['ID']}]")

            total_requested = [grp.strip().split(":")[0] for grp in group_names.split(",") if grp.strip()]
            total_resolved = {g["Name"] for g in dyn_clean + graph_clean + mail_clean}
            unresolved = [grp for grp in total_requested if grp not in total_resolved]

            log.info(f"Group membership summary for [{user_identifier}]: Requested={len(total_requested)}, Resolved={len(total_resolved)}, Unresolved={len(unresolved)}")

            for msg in successes:
                record_result(log, ResultLevel.SUCCESS, msg)
            for msg in failures:
                record_result(log, ResultLevel.WARNING, msg)
            for name in unresolved:
                record_result(log, ResultLevel.WARNING, f"Failed to resolve group membership for [{user_identifier}] - [{name}]")

            record_result(log, ResultLevel.WARNING if (failures or unresolved) else ResultLevel.SUCCESS, f"Completed membership lookup for [{user_identifier}]")

        elif operation == "Add user to AAD groups":
            log.info("Operation: Add user to AAD groups")
            group_ids = []
            group_id_map = {}
            group_input_list = [g.strip() for g in group_names.split(",") if g.strip()]
            fail_resolved = []

            log.info(f"Attempting to resolve [{len(group_input_list)}] input groups")

            for i, original_name in enumerate(group_input_list):
                log.info(f"Checking resolved group index [{i}] for [{original_name}]")

                if i >= len(resolved_groups):
                    fail_resolved.append(original_name)
                    record_result(log, ResultLevel.WARNING, f"No result returned for group '{original_name}' (index {i})")
                    continue

                entry = resolved_groups[i]
                group_id = ""

                if isinstance(entry, tuple):
                    group_id = entry[1]
                elif isinstance(entry, list) and len(entry) == 1:
                    group_id = entry[0].get("id", "")

                if not group_id:
                    fail_resolved.append(original_name)
                    record_result(log, ResultLevel.WARNING, f"Failed to resolve group [{original_name}]")
                    continue

                if group_id in existing_group_ids:
                    log.info(f"User [{user_identifier}] already in group [{original_name}] - skipping add")
                    record_result(log, ResultLevel.SUCCESS, f"User already a member of group [{original_name}]")
                    continue

                log.info(f"Resolved group [{original_name}] to ID [{group_id}] for add")
                group_ids.append(group_id)
                group_id_map[group_id] = original_name

            if not group_ids:
                log.info("No valid groups to add user to")
                if fail_resolved:
                    data_to_log["Failed"] = ", ".join([f"Add AAD group permission - {g} as {access}" for g in fail_resolved])
                record_result(log, ResultLevel.SUCCESS if not fail_resolved else ResultLevel.WARNING, f"No valid groups to process for [{user_identifier}]")
                return

            group_ids_str = ",".join(group_ids)
            log.info(f"Submitting add request for [{len(group_ids)}] groups: {group_ids_str}")
            success, fail = add_user_to_aad_groups(log, http_client, msgraph_base_url, group_ids_str, user_id, access, access_token)

            if success:
                log.info(f"Add succeeded for [{len(success)}] groups")
                data_to_log["Success"] = ", ".join([f"{group_id_map.get(g, g)}:{access}" for g in success])
            if fail:
                log.warning(f"Add failed for [{len(fail)}] groups")
                failed_names = [group_id_map.get(g, g) for g in fail]
                existing_failed = data_to_log.get("Failed", "")
                additional = ", ".join([f"Add AAD group permission - {g} as {access}" for g in failed_names])
                data_to_log["Failed"] = f"{existing_failed}, {additional}".strip(", ")

            for g in success:
                record_result(log, ResultLevel.SUCCESS, f"Added user to AAD group - {group_id_map.get(g, g)} as {access}")
            for g in fail:
                record_result(log, ResultLevel.WARNING, f"Failed to add user to AAD group - {group_id_map.get(g, g)} as {access}")

            record_result(
                log,
                ResultLevel.WARNING if fail or fail_resolved else ResultLevel.SUCCESS,
                f"Completed group assignment for [{user_identifier}]"
            )

        elif operation == "Remove user from AAD groups":
            log.info("Operation: Remove user from AAD groups")
            group_ids = [g.strip() for g in group_names.split(",") if g.strip()]
            group_id_map = {gid: gid for gid in group_ids}

            log.info(f"Requested removal from [{len(group_ids)}] groups")

            remove_ids = []

            for group_id in group_ids:
                if group_id not in existing_group_ids:
                    log.info(f"User [{user_identifier}] not in group [{group_id}] - skipping remove")
                    record_result(log, ResultLevel.SUCCESS, f"User not in group [{group_id}], nothing to remove")
                    continue
                remove_ids.append(group_id)

            log.info(f"User [{user_identifier}] will be removed from [{len(remove_ids)}] groups")

            if not remove_ids:
                log.warning(f"No valid groups found for removal")
                record_result(log, ResultLevel.WARNING, f"No groups to remove for [{user_identifier}]")
                return

            group_ids_str = ",".join(remove_ids)
            success, fail = remove_user_from_groups(log, http_client, msgraph_base_url, group_ids_str, user_id, access, access_token)

            if success:
                log.info(f"Successfully removed from [{len(success)}] groups")
                data_to_log["Success"] = ", ".join([f"{g}:{access}" for g in success])
            if fail:
                log.warning(f"Failed to remove from [{len(fail)}] groups")
                data_to_log["Failed"] = ", ".join([f"Remove AAD group permission - {g} as {access}" for g in fail])

            for g in success:
                record_result(log, ResultLevel.SUCCESS, f"Removed user from AAD group - {g} as {access}")
            for g in fail:
                record_result(log, ResultLevel.WARNING, f"Failed to remove user from AAD group - {g} as {access}")

            record_result(
                log,
                ResultLevel.WARNING if fail else ResultLevel.SUCCESS,
                f"Completed group removal for [{user_identifier}]"
            )

        elif operation == "Remove user from all AAD groups":
            log.info("Operation: Remove user from all AAD groups")

            removed, skipped = remove_user_from_all_groups(
                log, http_client, msgraph_base_url, user_id, user_identifier, access_token
            )

            log.info(f"Removal summary - Removed: {len(removed)}, Skipped: {len(skipped)}")

            if removed:
                data_to_log["Success"] = ", ".join(removed)
            if skipped:
                data_to_log["Failed"] = ", ".join(
                    [f"Remove AAD group permission - {g[0]} as {g[2]}" for g in skipped]
                )

            for g in removed:
                record_result(log, ResultLevel.SUCCESS, f"Removed user from AAD group - {g}")
            for g in skipped:
                record_result(
                    log,
                    ResultLevel.WARNING,
                    f"Failed to remove user from AAD group - {g[0]} as {g[2]}"
                )

            record_result(
                log,
                ResultLevel.WARNING if skipped else ResultLevel.SUCCESS,
                f"Completed removal from all AAD groups for [{user_identifier}]"
            )

        else:
            record_result(log, ResultLevel.WARNING,
                          f"Unsupported operation: {operation}")
            return

    except Exception:
        log.result_message(ResultLevel.WARNING, "Process failed")

    finally:
        log.info("Main execution complete - entering final result logging")
        log.result_data(data_to_log)


if __name__ == "__main__":
    main()