import sys
import traceback
import json
import random
import re
import subprocess
import os
import io
import base64
import hmac
import hashlib
import time
import urllib.parse
import string
import requests
import pandas as pd
from collections import defaultdict
from datetime import datetime, timedelta, timezone
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
    secret_url = (f"https://{vault_name}.vault.azure.net/secrets/{secret_name}?api-version=7.3")
    response = execute_api_call(
        log,
        http_client,
        "get",
        secret_url,
        integration_name="custom_wf_oauth2_client_creds",
    )
    if response and response.status_code == 200:
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
    else:
        log.error(f"No response received when retrieving ticket [{ticket_number}]")

    return "", "", 0, ""

def get_ticket_data(log, http_client, cwpsa_base_url, ticket_number):
    try:
        log.info(f"Retrieving full ticket details for ticket number [{ticket_number}]")
        endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}"
        response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
        if not response or response.status_code != 200:
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

def get_m365_contact(log, exo_access_token, azure_domain, contact_email):
    ps_command = f"""
    $Token = '{exo_access_token}'
    Connect-ExchangeOnline -AccessToken $Token -Organization '{azure_domain}' -ShowBanner:$false
    Get-MailContact -Identity "{contact_email}" | Select DisplayName,ExternalEmailAddress,FirstName,LastName,PhoneNumber | ConvertTo-Json -Compress
    """
    return execute_powershell(log, ps_command, shell="pwsh")

def add_m365_contact(log, exo_access_token, azure_domain, contact_email, display_name):
    ps_command = f"""
    $Token = '{exo_access_token}'
    Connect-ExchangeOnline -AccessToken $Token -Organization '{azure_domain}' -ShowBanner:$false
    New-MailContact -Name "{display_name}" -ExternalEmailAddress "{contact_email}" -DisplayName "{display_name}"
    """
    return execute_powershell(log, ps_command, shell="pwsh")


def update_m365_contact(log, exo_access_token, azure_domain, contact_email, display_name=""):
    if not display_name:
        log.warning("No display name provided for update")
        return False, "No update fields"

    ps_command = f"""
    $Token = '{exo_access_token}'
    Connect-ExchangeOnline -AccessToken $Token -Organization '{azure_domain}' -ShowBanner:$false
    Set-MailContact -Identity "{contact_email}" -DisplayName "{display_name}"
    """
    return execute_powershell(log, ps_command, shell="pwsh")

def remove_m365_contact(log, exo_access_token, azure_domain, contact_email):
    ps_command = f"""
    $Token = '{exo_access_token}'
    Connect-ExchangeOnline -AccessToken $Token -Organization '{azure_domain}' -ShowBanner:$false
    Remove-MailContact -Identity "{contact_email}" -Confirm:$false
    """
    return execute_powershell(log, ps_command, shell="pwsh")

def main():
    try:
        try:
            operation = input.get_value("Operation_1751842650075")
            ticket_number = input.get_value("TicketNumber_1751842647575")
            auth_code = input.get_value("AuthCode_1751842657377")
            provided_token = input.get_value("AccessToken_1751842659105")
            contact_email = input.get_value("Email_1751842655111")
            display_name = input.get_value("DisplayName_1751842653133")
        except Exception:
            log.warning("Failed to fetch input values")
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        operation = operation.strip() if operation else ""
        ticket_number = ticket_number.strip() if ticket_number else ""
        auth_code = auth_code.strip() if auth_code else ""
        provided_token = provided_token.strip() if provided_token else ""
        contact_email = contact_email.strip() if contact_email else ""
        display_name = display_name.strip() if display_name else ""

        log.info(f"Requested operation = [{operation}]")
        log.info(f"Target contact = [{contact_email}]")

        if operation not in ["Get Contact", "Add Contact", "Update Contact", "Remove Contact"]:
            log.warning(f"Unsupported operation [{operation}]")
            record_result(log, ResultLevel.WARNING, f"Unsupported operation [{operation}]")
            return
        if not contact_email and operation != "Add Contact":
            log.warning("Contact email is missing")
            record_result(log, ResultLevel.WARNING, "Contact email is missing")
            return
        if operation == "Add Contact" and (not display_name or not contact_email):
            log.warning("Display name and email are required to add a contact")
            record_result(log, ResultLevel.WARNING, "Display name and email are required to add a contact")
            return
        if operation == "Update Contact" and (not display_name or not contact_email):
            log.warning("Display name and email are required to update a contact")
            record_result(log, ResultLevel.WARNING, "Display name and email are required to update a contact")
            return

        access_token = ""
        company_identifier = ""
        azure_domain = ""
        tenant_id = ""

        if provided_token:
            access_token = provided_token
            log.info("Using provided access token")
            if not isinstance(access_token, str) or "." not in access_token:
                log.warning("Provided access token is malformed (missing dots)")
                record_result(log, ResultLevel.WARNING, "Provided access token is malformed (missing dots)")
                return
        elif ticket_number:
            company_identifier, _, _, _ = get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
            if not company_identifier:
                log.warning(f"Failed to retrieve company identifier from ticket [{ticket_number}]")
                record_result(log, ResultLevel.WARNING, f"Failed to retrieve company identifier from ticket [{ticket_number}]")
                return

            if company_identifier == "MIT":
                if not validate_mit_authentication(log, http_client, vault_name, auth_code):
                    return

            client_id = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientID")
            client_secret = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientSecret")
            azure_domain = get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")

            if not all([client_id, client_secret, azure_domain]):
                log.error("Missing secrets for token generation")
                record_result(log, ResultLevel.WARNING, "Missing secrets for token generation")
                return

            tenant_id = get_tenant_id_from_domain(log, http_client, azure_domain)
            if not tenant_id:
                log.error("Failed to resolve tenant ID")
                record_result(log, ResultLevel.WARNING, "Failed to resolve tenant ID")
                return

            access_token = get_access_token(log, http_client, tenant_id, client_id, client_secret)
            if not access_token or "." not in access_token:
                log.warning("Failed to retrieve valid access token")
                record_result(log, ResultLevel.WARNING, "Failed to retrieve valid access token")
                return
        else:
            log.warning("Either Access Token or Ticket Number is required")
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
        
        if operation == "Get Contact":
            success, output = get_m365_contact(log, exo_access_token, azure_domain, contact_email)
            if success and output:
                try:
                    parsed = json.loads(output)
                    data_to_log["ContactDetails"] = parsed
                except:
                    log.warning("Failed to parse contact output as JSON")
                log.info(f"Retrieved contact [{contact_email}]")
                record_result(log, ResultLevel.SUCCESS, f"Retrieved contact [{contact_email}]")
            else:
                log.warning(f"Failed to retrieve contact [{contact_email}]")
                record_result(log, ResultLevel.WARNING, f"Failed to retrieve contact [{contact_email}]")
        
        elif operation == "Add Contact":
            success, output = add_m365_contact(log, exo_access_token, azure_domain, contact_email, display_name)
            if success:
                log.info(f"Successfully added contact [{contact_email}]")
                record_result(log, ResultLevel.SUCCESS, f"Successfully added contact [{contact_email}]")
            else:
                log.warning(f"Failed to add contact [{contact_email}]")
                record_result(log, ResultLevel.WARNING, f"Failed to add contact [{contact_email}]")

        elif operation == "Update Contact":
            check_success, _ = get_m365_contact(log, exo_access_token, azure_domain, contact_email)
            if not check_success:
                log.warning(f"Contact [{contact_email}] does not exist for update")
                record_result(log, ResultLevel.WARNING, f"Contact [{contact_email}] does not exist for update")
                return

            success, output = update_m365_contact(log, exo_access_token, azure_domain, contact_email, display_name)
            if success:
                log.info(f"Successfully updated contact [{contact_email}]")
                record_result(log, ResultLevel.SUCCESS, f"Successfully updated contact [{contact_email}]")
            else:
                log.warning(f"Failed to update contact [{contact_email}]")
                record_result(log, ResultLevel.WARNING, f"Failed to update contact [{contact_email}]")

        elif operation == "Remove Contact":
            success, output = remove_m365_contact(log, exo_access_token, azure_domain, contact_email)
            if success:
                log.info(f"Successfully removed contact [{contact_email}]")
                record_result(log, ResultLevel.SUCCESS, f"Successfully removed contact [{contact_email}]")
            else:
                log.warning(f"Failed to remove contact [{contact_email}]")
                record_result(log, ResultLevel.WARNING, f"Failed to remove contact [{contact_email}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()