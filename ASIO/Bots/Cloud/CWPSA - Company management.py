import sys
import random
import os
import time
import requests
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

cwpsa_base_url = "https://aus.myconnectwise.net"
cwpsa_base_url_path = "/v4_6_release/apis/3.0"
vault_name = "PLACEHOLDER-akv1"

data_to_log = {}
bot_name = "CWPSA - Company management"
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

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1757463987470")
            operation = input.get_value("Operation_1757465437201")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        operation = operation.strip() if operation else ""

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number input is required")
            return
        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation input is required")
            return

        company_identifier, company_name, company_id, company_types = get_company_data_from_ticket(
            log, http_client, cwpsa_base_url, ticket_number
        )
        if not company_id:
            record_result(log, ResultLevel.WARNING, f"Unable to retrieve company details from ticket [{ticket_number}]")
            return

        if operation == "Check ASIO Company Type":
            matched_type = next((t for t in company_types if t in ["ASIO - Hybrid", "ASIO - Cloud"]), None)
            if matched_type:
                data_to_log["company_identifier"] = company_identifier
                data_to_log["company_id"] = company_id
                data_to_log["company_name"] = company_name
                data_to_log["company_type"] = matched_type
                record_result(log, ResultLevel.SUCCESS, f"Valid ASIO company type found: {matched_type}")
            else:
                record_result(log, ResultLevel.WARNING, f"No valid ASIO company type found. Types: {', '.join(company_types)}")
        else:
            record_result(log, ResultLevel.WARNING, f"Unknown operation [{operation}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
