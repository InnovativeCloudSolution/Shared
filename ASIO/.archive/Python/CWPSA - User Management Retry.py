import sys
import random
import os
import time
import urllib.parse
import requests
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

            if response.status_code in [200, 202, 204]:
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


def get_ticket_details(log, http_client, cwpsa_base_url, ticket_number):
    try:
        log.info(f"Retrieving full ticket details for ticket number [{ticket_number}]")
        endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}"

        response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
        if not response or response.status_code != 200:
            log.error(f"Failed to retrieve ticket [{ticket_number}] - Status: {response.status_code if response else 'N/A'}")
            return "", "", ""

        ticket = response.json()
        ticket_summary = ticket.get("summary", "")
        ticket_type = ticket.get("type", {}).get("name", "")
        priority_name = ticket.get("priority", {}).get("name", "")

        log.info(f"Ticket [{ticket_number}] Summary = [{ticket_summary}], Type = [{ticket_type}], Priority = [{priority_name}]")
        return ticket_summary, ticket_type, priority_name

    except Exception as e:
        log.exception(e, f"Exception occurred while retrieving ticket details for [{ticket_number}]")
        return "", "", ""

def parse_onboarding_offboarding_data(log, ticket_summary):
    try:

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1748410807092")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        log.info(f"Raw inputs received: ticket_number=[{ticket_number}]")
        ticket_number = ticket_number.strip() if ticket_number else ""
        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number is required but not provided")
            return
        log.info(f"Processed inputs: ticket_number=[{ticket_number}]")

        ticket_summary, ticket_type, priority_name = get_ticket_details(log, http_client, cwpsa_base_url, ticket_number)
        onboarding_offobarding_type = parse_onboarding_offboarding_data(log, ticket_summary)
        

    except Exception:
        record_result(log, ResultLevel.WARNING, "Process failed")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()