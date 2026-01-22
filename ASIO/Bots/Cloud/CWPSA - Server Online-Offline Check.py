import sys
import random
import os
import time
import urllib.parse
import requests
import re
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

cwpsa_base_url = "https://aus.myconnectwise.net"
cwpsa_base_url_path = "/v4_6_release/apis/3.0"

data_to_log = {}
bot_name = "CWPSA - Server Online-Offline Check"
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

def get_ticket_data(log, http_client, cwpsa_base_url, cwpsa_base_url_path, ticket_number):
    try:
        log.info(f"Retrieving full ticket details for ticket number [{ticket_number}]")
        endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets/{ticket_number}"
        response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
        if response:
            ticket = response.json()
            return ticket
        return None
    except Exception as e:
        log.exception(e, f"Exception occurred while retrieving ticket details for [{ticket_number}]")
        return None

def get_ticket_notes(log, http_client, cwpsa_base_url, cwpsa_base_url_path, ticket_number):
    log.info(f"Retrieving notes for ticket [{ticket_number}]")
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets/{ticket_number}/notes"
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    if response:
        try:
            notes = response.json()
            log.info(f"Retrieved {len(notes)} notes for ticket [{ticket_number}]")
            return notes
        except Exception as e:
            log.exception(e, f"Failed to parse notes response for ticket [{ticket_number}]")
            return []
    log.warning(f"No notes found for ticket [{ticket_number}]")
    return []

def extract_server_name(log, summary):
    patterns = [
        r'\b[A-Z]{2,}-SRV-[A-Z]{2,}-[0-9]{3,}\b',
        r'\b[A-Z]{2,}-SRV-[0-9]{3,}\b',
        r'\b[A-Z]{2,4}-[A-Z0-9]{5,}\b'
    ]
    for pattern in patterns:
        matches = re.findall(pattern, summary, re.IGNORECASE)
        if matches:
            server_name = matches[0].upper()
            log.info(f"Found server name: {server_name}")
            return server_name
    log.info("No server name found in summary")
    return ""

def find_configuration_by_name(log, http_client, cwpsa_base_url, cwpsa_base_url_path, config_name, company_id, config_type="Server"):
    log.info(f"Searching for configuration with name [{config_name}] and type [{config_type}] in company [{company_id}]")
    conditions = f'name="{config_name}" AND type/name="{config_type}" AND company/id={company_id}'
    encoded_conditions = urllib.parse.quote(conditions)
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/company/configurations?conditions={encoded_conditions}"
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    if response:
        try:
            configs = response.json()
            if configs and len(configs) > 0:
                config = configs[0]
                log.info(f"Found existing configuration ID [{config.get('id')}] with name: [{config_name}]")
                return config
        except Exception as e:
            log.exception(e, f"Failed to parse configuration search response for name: [{config_name}]")
            return None
    log.info(f"No configuration found with name: [{config_name}]")
    return None

def attach_configuration_to_ticket(log, http_client, cwpsa_base_url, cwpsa_base_url_path, ticket_id, config_id):
    log.info(f"Attaching configuration [{config_id}] to ticket [{ticket_id}]")
    data = {"id": config_id}
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets/{ticket_id}/configurations"
    response = execute_api_call(log, http_client, "post", endpoint, data=data, integration_name="cw_psa")
    if response and response.status_code in [200, 201]:
        log.info(f"Successfully attached configuration [{config_id}] to ticket [{ticket_id}]")
        return True
    else:
        log.error(f"Failed to attach configuration [{config_id}] to ticket [{ticket_id}]")
        return False

def main():
    try:
        try:
            ticket_number = input.get_value("cwTicketId")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch ticket number input")
            return

        if not ticket_number or ticket_number == "None":
            record_result(log, ResultLevel.WARNING, "Ticket number input is required")
            return

        ticket_number = str(ticket_number).strip()
        log.info(f"Processing ticket [{ticket_number}] for server online/offline check")

        ticket_data = get_ticket_data(log, http_client, cwpsa_base_url, cwpsa_base_url_path, ticket_number)
        if not ticket_data:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve ticket data for [{ticket_number}]")
            return

        ticket_summary = ticket_data.get("summary", "")
        company_id = ticket_data.get("company", {}).get("id")
        log.info(f"Ticket summary: [{ticket_summary}]")
        log.info(f"Company ID: [{company_id}]")
        data_to_log["ticket_summary"] = ticket_summary

        if not company_id:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve company ID from ticket [{ticket_number}]")
            return

        if "reporting offline" not in ticket_summary.lower():
            log.info("Ticket summary does not contain 'reporting offline'")
            data_to_log["is_server_alert"] = False
            record_result(log, ResultLevel.SUCCESS, "Ticket is not a server offline alert")
            return

        server_name = extract_server_name(log, ticket_summary)
        if not server_name:
            log.info("No server name found in summary")
            data_to_log["is_server_alert"] = False
            data_to_log["server_name"] = ""
            record_result(log, ResultLevel.SUCCESS, "Ticket contains 'reporting offline' but no server name detected")
            return

        data_to_log["is_server_alert"] = True
        data_to_log["server_name"] = server_name
        log.info(f"Detected as server offline alert for: {server_name}")

        server_config = find_configuration_by_name(log, http_client, cwpsa_base_url, cwpsa_base_url_path, server_name, company_id, "Server")
        if not server_config:
            log.info(f"Server configuration not found for [{server_name}] - skipping attachment")
            data_to_log["server_config_found"] = False
            data_to_log["server_config_attached"] = False
        else:
            data_to_log["server_config_found"] = True
            data_to_log["server_config_id"] = server_config.get("id")
            log.info(f"Found existing server configuration ID [{server_config.get('id')}]")
            server_config_id = server_config.get("id")
            attach_success = attach_configuration_to_ticket(log, http_client, cwpsa_base_url, cwpsa_base_url_path, ticket_number, server_config_id)
            if attach_success:
                data_to_log["server_config_attached"] = True
                log.info(f"Successfully attached server configuration [{server_config_id}] to ticket [{ticket_number}]")
            else:
                data_to_log["server_config_attached"] = False
                log.warning(f"Failed to attach server configuration [{server_config_id}] to ticket [{ticket_number}]")

        notes = get_ticket_notes(log, http_client, cwpsa_base_url, cwpsa_base_url_path, ticket_number)
        if not notes or len(notes) == 0:
            log.info("No notes found for ticket")
            data_to_log["server_status"] = "offline"
            data_to_log["has_resolved_note"] = False
            record_result(log, ResultLevel.SUCCESS, f"Server offline alert detected for {server_name} - No notes found, status: offline")
            return

        notes_sorted = sorted(notes, key=lambda x: x.get("dateCreated", ""), reverse=True)
        latest_note = notes_sorted[0]
        latest_note_text = latest_note.get("text", "")
        log.info(f"Latest note text: {latest_note_text[:150]}...")
        data_to_log["latest_note_preview"] = latest_note_text[:200]

        resolved_message = "Our Monitoring agent has detected that this issue is now resolved. You can close this ticket."
        if resolved_message in latest_note_text:
            log.info("Resolved message found in latest note - server is online")
            data_to_log["server_status"] = "online"
            data_to_log["has_resolved_note"] = True
            record_result(log, ResultLevel.SUCCESS, f"Server {server_name} is now ONLINE - Resolved message detected")
        else:
            log.info("Resolved message not found in latest note - server is offline")
            data_to_log["server_status"] = "offline"
            data_to_log["has_resolved_note"] = False
            record_result(log, ResultLevel.SUCCESS, f"Server {server_name} is OFFLINE - No resolved message detected")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
