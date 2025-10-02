import sys
import random
import re
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

def get_user_groups(log, http_client, msgraph_base_url, user_id, token):
    log.info(f"Fetching groups for user [{user_id}]")

    endpoint = f"{msgraph_base_url}/users/{user_id}/transitiveMemberOf"
    headers = {"Authorization": f"Bearer {token}"}
    all_groups = []
    dynamic_groups = []
    graph_groups = []
    skipped_mail_groups = []

    while endpoint:
        response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
        if not response or response.status_code != 200:
            log.error(f"Failed to retrieve groups for [{user_id}] Status code: {response.status_code if response else 'N/A'}")
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

def remove_user_from_groups(log, http_client, msgraph_base_url, user_id, user_identifier, token, graph_groups):
    removed_groups = []
    skipped_groups = []

    for group_name, group_id in graph_groups:
        log.info(f"Attempting to remove user [{user_identifier}] from group [{group_name}] - [{group_id}] using Graph API")
        endpoint = f"{msgraph_base_url}/groups/{group_id}/members/{user_id}/$ref"
        headers = {"Authorization": f"Bearer {token}"}
        response = execute_api_call(log, http_client, "delete", endpoint, headers=headers)

        if response and response.status_code == 204:
            log.info(f"User [{user_identifier}] removed from [{group_name}] - [{group_id}] successfully via Graph API")
            removed_groups.append(f"{group_name} - {group_id}")
        elif response and response.status_code == 403 and "Authorization_RequestDenied" in response.text:
            log.warning(f"Permission denied for group [{group_name}] - [{group_id}]")
            skipped_groups.append((group_name, group_id))
        elif response and response.status_code == 404:
            log.warning(f"Skipping non-existent group [{group_name}] - [{group_id}]")
        else:
            log.error(f"Graph removal failed for group [{group_name}] - [{group_id}] Status: {response.status_code if response else 'N/A'}")
            skipped_groups.append((group_name, group_id))

    return removed_groups, skipped_groups

def main():
    try:
        try:
            user_identifier = input.get_value("User_1738815367863")
            ticket_number = input.get_value("TicketNumber_1742953896873")
            auth_code = input.get_value("AuthCode_1743025274034")
            provided_token = input.get_value("AccessToken_1742807317478")
        except Exception as e:
            log.exception(e, "Failed to fetch input values")
            log.result_message(ResultLevel.FAILED, "Failed to fetch input values")
            return
        
        user_identifier = user_identifier.strip() if user_identifier else ""
        ticket_number = ticket_number.strip() if ticket_number else ""
        auth_code = auth_code.strip() if auth_code else ""
        provided_token = provided_token.strip() if provided_token else ""

        log.info(f"Received input user = [{user_identifier}], Ticket = [{ticket_number}]")

        if not user_identifier or not user_identifier.strip():
            log.error("User identifier is empty or invalid")
            log.result_message(ResultLevel.FAILED, "User identifier is empty or invalid")
            return

        if provided_token and provided_token.strip():
            access_token = provided_token.strip()
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

        dynamic_groups, graph_groups, skipped_mail_groups = get_user_groups(log, http_client, msgraph_base_url, user_id, access_token)

        removed = []
        skipped = []

        removed_graph, skipped_graph = remove_user_from_groups(
            log, http_client, msgraph_base_url, user_id, user_identifier, access_token, graph_groups
        )
        removed.extend(removed_graph)
        skipped.extend(skipped_graph)

        if removed:
            for group in removed:
                log.result_message(ResultLevel.SUCCESS, f"Removed [{user_email}] from [{group}]")

        if dynamic_groups:
            for name, group_id in dynamic_groups:
                log.result_message(ResultLevel.WARNING, f"Skipped dynamic group (cannot remove) [{name}] - [{group_id}]")

        if skipped_mail_groups:
            for name, group_id in skipped_mail_groups:
                log.result_message(ResultLevel.WARNING, f"Skipped mail-enabled group (not targeted) [{name}] - [{group_id}]")

        if skipped:
            for name, group_id in skipped:
                log.result_message(ResultLevel.WARNING, f"Skipped group (removal failed) [{name}] - [{group_id}]")

        if not removed and not dynamic_groups and not skipped and not skipped_mail_groups:
            log.result_message(ResultLevel.SUCCESS, f"Nothing to remove for user [{user_email}]")

    except Exception:
        log.exception("An error occurred while processing")
        log.result_message(ResultLevel.FAILED, "Process failed")

if __name__ == "__main__":
    main()