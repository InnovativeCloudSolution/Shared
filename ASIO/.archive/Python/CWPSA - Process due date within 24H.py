import sys
import random
import os
import time
import urllib.parse
import requests
from datetime import datetime, timedelta, timezone
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

cwpsa_base_url = "https://au.myconnectwise.net/v4_6_release/apis/3.0"
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


def proccess_24_hour_due_date(log, ticket_due):
    try:
        log.info(f"Evaluating ticket required date: {ticket_due}")
        if not ticket_due:
            return False, "No required date set on the ticket"

        required_utc = datetime.strptime(ticket_due, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
        now_aest = datetime.now(timezone(timedelta(hours=10)))
        required_aest = required_utc.astimezone(timezone(timedelta(hours=10)))

        log.info(f"Converted required date to AEST: {required_aest.isoformat()}")
        delta = required_aest - now_aest

        if timedelta(0) <= delta <= timedelta(hours=24):
            return True, f"Ticket is due within 24 hours (AEST): {required_aest.isoformat()}"
        return False, f"Ticket is not due within 24 hours (AEST): {required_aest.isoformat()}"
    except Exception as e:
        log.exception(e, "Error parsing or converting required date")
        return False, "Error evaluating due date"


def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1742953896873")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        log.info(f"Starting ticket due date evaluation for ticket [{ticket_number}]")

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number input is required")
            return

        _, _, _, ticket_due = get_ticket_data(
            log, http_client, cwpsa_base_url, ticket_number
        )

        data_to_log.update({
            "ticket_due": ticket_due
        })

        within_24, message = proccess_24_hour_due_date(log, ticket_due)
        if within_24:
            record_result(log, ResultLevel.SUCCESS, message)
        else:
            record_result(log, ResultLevel.WARNING, message)

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)


if __name__ == "__main__":
    main()