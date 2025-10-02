import sys
import json
import random
import os
import time
import requests
import urllib.parse
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
azure_cosmodb_helper_url = "https://prodautomationfunction.azurewebsites.net/api/cosmodb-helper"

data_to_log = {}
bot_name = "MIT-UTIL - Webhook Action Trigger"
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

def get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number):
    log.info(f"Retrieving company data for ticket [{ticket_number}]")
    
    ticket_url = f"{cwpsa_base_url}/service/tickets/{ticket_number}"
    ticket_response = execute_api_call(log, http_client, "get", ticket_url, integration_name="cwpsa")
    
    if not ticket_response or ticket_response.status_code != 200:
        log.error(f"Failed to retrieve ticket [{ticket_number}] details")
        return "", "", "", ""
    
    ticket_data = ticket_response.json()
    company_id = ticket_data.get("company", {}).get("id", "")
    
    if not company_id:
        log.error(f"No company ID found in ticket [{ticket_number}]")
        return "", "", "", ""
    
    company_url = f"{cwpsa_base_url}/company/companies/{company_id}"
    company_response = execute_api_call(log, http_client, "get", company_url, integration_name="cwpsa")
    
    if not company_response or company_response.status_code != 200:
        log.error(f"Failed to retrieve company [{company_id}] details")
        return "", "", "", ""
    
    company_data = company_response.json()
    company_identifier = company_data.get("identifier", "")
    company_name = company_data.get("name", "")
    company_types = company_data.get("types", [])
    
    log.info(f"Successfully retrieved company data: identifier=[{company_identifier}], name=[{company_name}], id=[{company_id}]")
    return company_identifier, company_name, company_id, company_types

def main():
    try:
        try:
            operation = input.get_value("Operation_1751427590883")
            ticket_number = input.get_value("TicketNumber_1751427569586")
            graph_token = input.get_value("GraphToken_xxxxxxxxxxxxx")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        operation = operation.strip() if operation else ""
        graph_token = graph_token.strip() if graph_token else ""

        log.info(f"Received input ticket number = [{ticket_number}]")
        log.info(f"Requested operation = [{operation}]")

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number is empty or invalid")
            return
        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation value is missing or invalid")
            return

        log.info(f"Retrieving company data for ticket [{ticket_number}]")
        company_identifier, company_name, company_id, company_types = get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
        if not company_identifier:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve company identifier from ticket [{ticket_number}]")
            return
        data_to_log["Company"] = company_identifier
        log.info(f"Resolved company from ticket [{ticket_number}]: identifier=[{company_identifier}], name=[{company_name}]")

        if graph_token:
            log.info("Using provided MS Graph token")
            graph_access_token = graph_token
            graph_tenant_id = ""
        else:
            graph_tenant_id, graph_access_token = get_graph_token(log, http_client, vault_name, company_identifier)
            if not graph_access_token:
                record_result(log, ResultLevel.WARNING, "Failed to obtain MS Graph access token")
                return

        if operation == "Schedule Offboarding":
            log.info("Executing operation: Schedule Offboarding")
            payload = {
                "cwpsa_ticket": int(ticket_number),
                "action": "get",
                "request_type": "user_offboarding"
            }
            log.info(f"Offboarding payload: {json.dumps(payload)}")
            success = execute_api_call(log, http_client, "post", azure_cosmodb_helper_url, data=payload)
            if success and success.status_code == 200:
                log.info(f"Offboarding action for ticket [{ticket_number}] sent successfully.")
                record_result(log, ResultLevel.SUCCESS, f"Offboarding action for ticket [{ticket_number}] completed successfully")
            else:
                log.error(f"Failed to send offboarding action for ticket [{ticket_number}]. Response: {success.text if success else 'No response'}")
                record_result(log, ResultLevel.WARNING, f"Failed to offboard user for ticket [{ticket_number}]")

        elif operation == "Schedule Onboarding":
            log.info("Executing operation: Schedule Onboarding")
            payload = {
                "cwpsa_ticket": int(ticket_number),
                "action": "get",
                "request_type": "user_onboarding"
            }
            log.info(f"Onboarding payload: {json.dumps(payload)}")
            success = execute_api_call(log, http_client, "post", azure_cosmodb_helper_url, data=payload)
            if success and success.status_code == 200:
                log.info(f"Onboarding action for ticket [{ticket_number}] sent successfully.")
                record_result(log, ResultLevel.SUCCESS, f"Onboarding action for ticket [{ticket_number}] completed successfully")
            else:
                log.error(f"Failed to send onboarding action for ticket [{ticket_number}]. Response: {success.text if success else 'No response'}")
                record_result(log, ResultLevel.WARNING, f"Failed to onboard user for ticket [{ticket_number}]")

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
