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

data_to_log = {}
bot_name = "M365 - MFA Management"
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
    log.error(f"Failed to retrieve secret [{secret_name}] Status code: {response.status_code if response else 'N/A'}")
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
        log.error(f"Failed to extract tenant ID from domain [{azure_domain}]")
        return ""
    except Exception as e:
        log.exception(e, "Exception while extracting tenant ID from domain")
        return ""

def get_access_token(log, http_client, tenant_id, client_id, client_secret, scope="https://graph.microsoft.com/.default"):
    log.info(f"Requesting access token for scope [{scope}]")
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
        log.info(f"Access token length: {len(access_token)}")
        log.info(f"Access token preview: {access_token[:30]}...")
        if not isinstance(access_token, str) or "." not in access_token:
            log.error("Access token is invalid or malformed")
            return ""
        log.info("Successfully retrieved access token")
        return access_token
    log.error(f"Failed to retrieve access token Status code: {response.status_code if response else 'N/A'}")
    return ""

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
    return "", "", 0, ""

def get_ticket_data(log, http_client, cwpsa_base_url, cwpsa_base_url_path, ticket_number):
    try:
        log.info(f"Retrieving full ticket details for ticket number [{ticket_number}]")
        endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets/{ticket_number}"
        response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
        if not response:
            log.error(
                f"Failed to retrieve ticket [{ticket_number}] - Status: {response.status_code if response else 'N/A'}"
            )
            return "", "", "", ""

        ticket = response.json()
        ticket_summary = ticket.get("summary", "")
        ticket_type = ticket.get("type", {}).get("name", "")
        priority_name = ticket.get("priority", {}).get("name", "")
        due_date = ticket.get("requiredDate", "")

        log.info(
            f"Ticket [{ticket_number}] Summary = [{ticket_summary}], Type = [{ticket_type}], Priority = [{priority_name}], Due = [{due_date}]"
        )
        return ticket_summary, ticket_type, priority_name, due_date

    except Exception as e:
        log.exception(e, f"Exception occurred while retrieving ticket details for [{ticket_number}]")
        return "", "", "", ""

def validate_mit_authentication(log, http_client, vault_name, auth_code):
    if not auth_code:
        log.result_message(ResultLevel.FAILED, "Authentication code input is required for MIT")
        return False
    expected_code = get_secret_value(log, http_client, vault_name, "MIT-AuthenticationCode")
    if not expected_code:
        log.result_message(
            ResultLevel.FAILED,
            "Failed to retrieve expected authentication code for MIT",
        )
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
    endpoint = f"{msgraph_base_url_base}{msgraph_base_url_path}/users?$filter={urllib.parse.quote(filter_query)}"
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
            )
    log.error(f"Failed to resolve user ID and email for [{user_identifier}]")
    return "", "", ""

def enable_mfa_graph_per_user(log, http_client, user_id, token, mfa_state):
    log.info(f"Setting legacy MFA state to [{mfa_state}] using Graph API (beta) for user ID [{user_id}]")
    endpoint = f"{msgraph_base_url_beta_base}{msgraph_base_url_beta_path}/users/{user_id}/authentication/requirements"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    payload = {
        "perUserMfaState": mfa_state.lower()
    }
    response = execute_api_call(log, http_client, "patch", endpoint, data=payload, headers=headers)
    if response:
        log.info(f"Successfully updated MFA state to [{mfa_state}] for [{user_id}] via Graph API")
        return True
    return False

def main():
    try:
        try:
            user_identifier = input.get_value("User_1750797544070")
            operation = input.get_value("Operation_1750797542101")
            ticket_number = input.get_value("TicketNumber_1750797535962")
            auth_code = input.get_value("AuthCode_1750797547868")
            graph_token = input.get_value("GraphToken_xxxxxxxxxxxxx")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        user_identifier = user_identifier.strip() if user_identifier else ""
        operation = operation.strip() if operation else ""
        ticket_number = ticket_number.strip() if ticket_number else ""
        auth_code = auth_code.strip() if auth_code else ""
        graph_token = graph_token.strip() if graph_token else ""

        log.info(f"Received input user = [{user_identifier}]")
        log.info(f"Requested operation = [{operation}]")

        if not user_identifier:
            record_result(log, ResultLevel.WARNING, "User identifier is empty or invalid")
            return
        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation value is missing or invalid")
            return

        access_token = ""
        company_identifier = ""

        if graph_token:
            access_token = graph_token
            log.info("Using provided MS Graph token")
            if not isinstance(access_token, str) or "." not in access_token:
                record_result(log, ResultLevel.WARNING, "Provided MS Graph token is malformed (missing dots)")
                return
        elif ticket_number:
            log.info(f"Retrieving company data for ticket [{ticket_number}]")
            company_identifier, _, _, _ = get_company_data_from_ticket(log, http_client, cwpsa_base_url, cwpsa_base_url_path, ticket_number)
            if not company_identifier:
                record_result(log, ResultLevel.WARNING, f"Failed to retrieve company identifier from ticket [{ticket_number}]")
                return

            if company_identifier == "MIT":
                if not validate_mit_authentication(log, http_client, vault_name, auth_code):
                    return

            client_id = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientID")
            client_secret = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientSecret")
            azure_domain = get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")

            if not all([client_id, client_secret, azure_domain]):
                record_result(log, ResultLevel.WARNING, "Failed to retrieve required secrets")
                return

            tenant_id = get_tenant_id_from_domain(log, http_client, azure_domain)
            if not tenant_id:
                record_result(log, ResultLevel.WARNING, "Failed to resolve tenant ID")
                return

            access_token = get_access_token(log, http_client, tenant_id, client_id, client_secret)
            if not isinstance(access_token, str) or "." not in access_token:
                record_result(log, ResultLevel.WARNING, "Access token is malformed (missing dots)")
                return

        aad_user_result = get_aad_user_data(log, http_client, msgraph_base_url, user_identifier, access_token)

        if isinstance(aad_user_result, list):
            details = "\n".join(
                [f"- {u.get('displayName')} | {u.get('userPrincipalName')} | {u.get('id')}" for u in aad_user_result]
            )
            record_result(log, ResultLevel.WARNING, f"Multiple users found for [{user_identifier}]\n{details}")
            return

        user_id, user_email, _ = aad_user_result

        if not user_id:
            record_result(log, ResultLevel.WARNING, f"Failed to resolve user ID for [{user_identifier}]")
            return
        if not user_email:
            record_result(log, ResultLevel.WARNING, f"Unable to resolve user principal name for [{user_identifier}]")
            return

        log.info(f"User [{user_identifier}] matched with UPN [{user_email}]")

        valid_mfa_states = ["enabled", "enforced", "disabled"]
        mfa_state = operation.strip().lower()

        if mfa_state in valid_mfa_states:
            log.info(f"Executing operation: Set MFA state to [{mfa_state}] using Graph API")
            success = enable_mfa_graph_per_user(log, http_client, user_id, access_token, mfa_state)
            if success:
                record_result(log, ResultLevel.SUCCESS, f"Legacy MFA state [{mfa_state}] set for {user_email}")
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to set legacy MFA state [{mfa_state}] for {user_email}")
            return

        record_result(log, ResultLevel.WARNING, f"Unknown operation [{operation}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
