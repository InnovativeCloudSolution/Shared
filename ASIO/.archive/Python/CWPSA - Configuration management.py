import sys
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

def get_configuration_data(log, http_client, cwpsa_base_url, company_id, flag_caption):
    conditions = f"company/id={company_id}"
    custom_conditions = f'caption="{flag_caption}" AND value=true'
    query_string = f"conditions={urllib.parse.quote(conditions)}&customFieldConditions={urllib.parse.quote(custom_conditions)}"
    url = f"{cwpsa_base_url}/company/configurations?{query_string}"

    response = execute_api_call(log, http_client, "get", url, integration_name="cw_psa")
    if response and response.status_code == 200:
        configs = response.json()
        for config in configs:
            custom_fields = config.get("customFields", [])
            has_flag = any(
                f.get("caption") == flag_caption and f.get("value") is True
                for f in custom_fields
            )
            if has_flag:
                endpoint_id = next(
                    (f.get("value") for f in custom_fields if f.get("caption") == "Endpoint ID"),
                    ""
                )
                return {
                    "name": config.get("name", ""),
                    "id": config.get("id"),
                    "endpoint_id": endpoint_id
                }
    return None

def get_user_configuration_items(log, http_client, cwpsa_base_url, company_id, contact_id):
    log.info(f"Fetching configurations for company ID [{company_id}] and contact ID [{contact_id}]")
    conditions = f"company/id={company_id} and contact/id={contact_id}"
    query_string = f"conditions={urllib.parse.quote(conditions)}"
    url = f"{cwpsa_base_url}/company/configurations?{query_string}"

    response = execute_api_call(log, http_client, "get", url, integration_name="cw_psa")
    if response and response.status_code == 200:
        try:
            items = response.json()
        except Exception:
            log.error("Failed to decode configuration response JSON")
            return "", ""
        formatted_line = []
        formatted_note = []
        for item in items:
            fields = [
                item.get("name", ""),
                item.get("tagNumber", ""),
                item.get("serialNumber", ""),
                item.get("manufacturer", {}).get("name", ""),
                item.get("modelNumber", "")
            ]
            non_blank_fields = [f"[{f}]" for f in fields if f]
            entry = " - ".join(non_blank_fields)
            formatted_line.append(entry)
            formatted_note.append(entry)
        return ", ".join(formatted_line), "\n".join(formatted_note)
    log.warning(f"No configuration items found for contact ID [{contact_id}] in company [{company_id}]")
    return "", ""

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1751355370627")
            operation = input.get_value("Operation_1751355374117")
            contact_id = input.get_value("ContactID_1751503960999").strip()

        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        operation = operation.strip() if operation else ""
        contact_id = contact_id.strip() if contact_id else ""

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number input is required")
            return
        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation input is required")
            return
        
        company_identifier, company_name, company_id, company_type = get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
        if not company_id:
            record_result(log, ResultLevel.WARNING, f"Unable to retrieve company details from ticket [{ticket_number}]")
            return

        if operation == "Get PDC and ADS":
            pdc = get_configuration_data(log, http_client, cwpsa_base_url, company_id, "PDC")
            if pdc:
                data_to_log["pdc_name"] = pdc["name"]
                data_to_log["pdc_id"] = pdc["id"]
                data_to_log["pdc_endpoint_id"] = pdc["endpoint_id"]
                record_result(log, ResultLevel.SUCCESS, f"PDC Name: {pdc['name']}")
                record_result(log, ResultLevel.SUCCESS, f"PDC ID: {pdc['id']}")
                record_result(log, ResultLevel.SUCCESS, f"PDC Endpoint ID: {pdc['endpoint_id']}")
                record_result(log, ResultLevel.SUCCESS, f"PDC: {pdc['name']} - {pdc['id']} - {pdc['endpoint_id']}")
            else:
                record_result(log, ResultLevel.WARNING, "No configuration with PDC=True found")

            ads = get_configuration_data(log, http_client, cwpsa_base_url, company_id, "ADS")
            if ads:
                data_to_log["ads_name"] = ads["name"]
                data_to_log["ads_id"] = ads["id"]
                data_to_log["ads_endpoint_id"] = ads["endpoint_id"]
                record_result(log, ResultLevel.SUCCESS, f"ADS Name: {ads['name']}")
                record_result(log, ResultLevel.SUCCESS, f"ADS ID: {ads['id']}")
                record_result(log, ResultLevel.SUCCESS, f"ADS Endpoint ID: {ads['endpoint_id']}")
                record_result(log, ResultLevel.SUCCESS, f"ADS: {ads['name']} - {ads['id']} - {ads['endpoint_id']}")
            else:
                record_result(log, ResultLevel.WARNING, "No configuration with ADS=True found")

        elif operation == "Get User Configuration Items":
            if not contact_id or not company_id:
                record_result(log, ResultLevel.WARNING, "Missing ContactID or CompanyID for configuration lookup")
                return
            device_line, device_note = get_user_configuration_items(log, http_client, cwpsa_base_url, company_id, contact_id)
            if device_line:
                data_to_log["Device"] = device_line
                record_result(log, ResultLevel.SUCCESS, "Configuration found assigned to user:")
                for line in device_note.split("\n"):
                    record_result(log, ResultLevel.SUCCESS, line)
            else:
                record_result(log, ResultLevel.SUCCESS, f"No configuration items found for contact ID [{contact_id}]")

        else:
            record_result(log, ResultLevel.WARNING, f"Unknown operation [{operation}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()