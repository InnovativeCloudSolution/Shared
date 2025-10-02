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
msgraph_base_url_beta = "https://graph.microsoft.com/beta"
vault_name = "mit-azu1-prod1-akv1"
immy_base_url = "https://mit.immy.bot"
immy_tenant_url = "manganoit.com.au"

data_to_log = {}
bot_name = "MIT-IMMY - Install Software"
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

def get_immy_endpoint(log, http_client, base_url, user_email, bearer_token):
    selected_computers = []
    endpoint = f"{base_url}/api/v1/computers/dx"
    headers = {
        "Authorization": bearer_token,
        "Content-Type": "application/json"
    }

    all_computers = []
    skip = 0
    page_size = 100
    more_records = True

    while more_records:
        params = {
            "skip": skip,
            "take": page_size
        }

        response = execute_api_call(log, http_client, "get", endpoint, headers=headers, params=params)

        if response:
            try:
                data = response.json()
                batch = data.get("data", [])
                all_computers.extend(batch)

                if len(batch) < page_size:
                    more_records = False
                else:
                    skip += page_size
            except Exception as e:
                log.exception(e, "Failed to parse ImmyBot computers response")
                more_records = False
        else:
            
            more_records = False

    selected_computers = [
        computer for computer in all_computers
        if (computer.get("primaryUserEmail") or "").lower() == user_email.lower()
    ]

    for comp in selected_computers:
        log.info(f"Computer matched: {comp.get('computerName')} - {comp.get('primaryUserEmail')}")

    return selected_computers

def get_immy_software(log, http_client, base_url, software_name, bearer_token):
    headers = {
        "Authorization": bearer_token,
        "Content-Type": "application/json"
    }

    local_endpoint = f"{base_url}/api/v1/software/local"
    response = execute_api_call(log, http_client, "get", local_endpoint, headers=headers)

    selected_software = None
    if response:
        try:
            softwares = response.json()
            selected_software = next(
                (s for s in softwares if (s.get("name") or "").lower() == software_name.lower()), 
                None
            )
        except Exception as e:
            log.exception(e, "Failed to parse local softwares list")

    if not selected_software:
        global_endpoint = f"{base_url}/api/v1/software/global"
        response = execute_api_call(log, http_client, "get", global_endpoint, headers=headers)
        if response:
            try:
                softwares = response.json()
                selected_software = next(
                    (s for s in softwares if (s.get("name") or "").lower() == software_name.lower()), 
                    None
                )
            except Exception as e:
                log.exception(e, "Failed to parse global softwares list")

    if selected_software:
        log.info(f"Software matched: {selected_software.get('name')}")

    return selected_software

def push_immy_software(log, http_client, base_url, software_id, selected_computers, bearer_token):
    maintenance_payload = {
        "fullMaintenance": False,
        "resolutionOnly": False,
        "detectionOnly": False,
        "inventoryOnly": False,
        "runInventoryInDetection": False,
        "cacheOnly": False,
        "useWinningDeployment": False,
        "deploymentId": None,
        "deploymentType": None,
        "maintenanceParams": {
            "maintenanceIdentifier": software_id,
            "maintenanceType": 0,
            "repair": False,
            "desiredSoftwareState": 5,
            "maintenanceTaskMode": 0
        },
        "skipBackgroundJob": True,
        "rebootPreference": 1,
        "scheduleExecutionAfterActiveHours": False,
        "useComputersTimezoneForExecution": False,
        "offlineBehavior": 2,
        "suppressRebootsDuringBusinessHours": False,
        "sendDetectionEmail": False,
        "sendDetectionEmailWhenAllActionsAreCompliant": False,
        "sendFollowUpEmail": False,
        "sendFollowUpOnlyIfActionNeeded": False,
        "showRunNowButton": False,
        "showPostponeButton": False,
        "showMaintenanceActions": False,
        "computers": [{"computerId": comp.get("id")} for comp in selected_computers],
        "tenants": []
    }

    headers = {
        "Authorization": bearer_token,
        "Content-Type": "application/json"
    }

    endpoint = f"{base_url}/api/v1/run-immy-service"
    response = execute_api_call(log, http_client, "post", endpoint, data=maintenance_payload, headers=headers)

    if response:
        return True
    else:
        
        return False

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1744072639252")
            operation = input.get_value("Operation_1744067751139")
            user_identifier = input.get_value("User_1744067723414")
            software_name = input.get_value("Software_1744067751139")
            auth_code = input.get_value("AuthCode_1744067804529")
            graph_token = input.get_value("GraphToken_xxxxxxxxxxxxx")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        operation = operation.strip() if operation else ""
        user_identifier = user_identifier.strip() if user_identifier else ""
        software_name = software_name.strip() if software_name else ""
        auth_code = auth_code.strip() if auth_code else ""
        graph_token = graph_token.strip() if graph_token else ""

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
        if not software_name:
            record_result(log, ResultLevel.WARNING, "Software name is empty or invalid")
            return

        log.info(f"Retrieving company data for ticket [{ticket_number}]")
        company_identifier, company_name, company_id, company_type = get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
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
            record_result(log, ResultLevel.WARNING, f"Unable to resolve user principal name for [{user_identifier}]")
            return

        if operation == "Install Software" or operation == "Immy Management":
            log.info(f"Executing operation: {operation}")
            
            immy_client_id = get_secret_value(log, http_client, vault_name, "MIT-AutomationApp-ClientID")
            immy_client_secret = get_secret_value(log, http_client, vault_name, "MIT-AutomationApp-ClientSecret")

            if not all([immy_client_id, immy_client_secret]):
                record_result(log, ResultLevel.WARNING, "Failed to retrieve required secrets for ImmyBot authentication")
                return

            immy_access_token = get_access_token(log, http_client, immy_tenant_url, immy_client_id, immy_client_secret, scope=f"{immy_base_url}/.default")
            if not immy_access_token:
                record_result(log, ResultLevel.WARNING, "Failed to obtain ImmyBot API token")
                return

            bearer_token = f"Bearer {immy_access_token}"

            selected_computers = get_immy_endpoint(log, http_client, immy_base_url, user_email, bearer_token)
            if not selected_computers:
                record_result(log, ResultLevel.WARNING, f"No computers found for user [{user_email}]")
                return

            selected_software = get_immy_software(log, http_client, immy_base_url, software_name, bearer_token)
            if not selected_software:
                record_result(log, ResultLevel.WARNING, f"Software [{software_name}] not found")
                return

            result = push_immy_software(log, http_client, immy_base_url, selected_software.get("id"), selected_computers, bearer_token)

            if result:
                pushed_computers = ", ".join([comp.get("computerName") for comp in selected_computers])
                record_result(log, ResultLevel.SUCCESS, f"Pushed software [{selected_software.get('name')}] to endpoints [{pushed_computers}] for user [{user_email}]")
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to trigger ImmyBot software deployment for user [{user_email}]")

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
