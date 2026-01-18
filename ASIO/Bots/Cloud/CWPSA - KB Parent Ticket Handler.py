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

cwpsa_base_url = "https://aus.myconnectwise.net/v4_6_release/apis/3.0"

data_to_log = {}
bot_name = "CWPSA - KB Parent Ticket Handler"
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

def get_ticket_data(log, http_client, cwpsa_base_url, ticket_number):
    try:
        log.info(f"Retrieving full ticket details for ticket number [{ticket_number}]")
        endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}"
        response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
        if response:
            ticket = response.json()
            return ticket
        return None
    except Exception as e:
        log.exception(e, f"Exception occurred while retrieving ticket details for [{ticket_number}]")
        return None

def get_ticket_notes(log, http_client, cwpsa_base_url, ticket_number):
    log.info(f"Retrieving notes for ticket [{ticket_number}]")
    endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}/notes"
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

def extract_kb_numbers(log, text):
    kb_pattern = r'(?i)\bkb\d+\b'
    matches = re.findall(kb_pattern, text)
    if matches:
        log.info(f"Found KB numbers: {matches}")
        return [match.upper() for match in matches]
    log.info("No KB numbers found in text")
    return []

def find_ticket_by_summary(log, http_client, cwpsa_base_url, summary, company_id):
    log.info(f"Searching for ticket with summary: [{summary}] in company ID [{company_id}]")
    conditions = f'summary="{summary}" AND company/id={company_id}'
    encoded_conditions = urllib.parse.quote(conditions)
    endpoint = f"{cwpsa_base_url}/service/tickets?conditions={encoded_conditions}"
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    if response:
        try:
            tickets = response.json()
            if tickets and len(tickets) > 0:
                ticket = tickets[0]
                log.info(f"Found existing ticket ID [{ticket.get('id')}] with summary: [{summary}]")
                return ticket
        except Exception as e:
            log.exception(e, f"Failed to parse ticket search response for summary: [{summary}]")
            return None
    log.info(f"No ticket found with summary: [{summary}]")
    return None

def create_ticket(log, http_client, cwpsa_base_url, summary, company_id, board_name="Service Desk", status_name="New"):
    log.info(f"Creating new ticket with summary: [{summary}] in board [{board_name}]")
    
    ticket_data = {
        "summary": summary,
        "company": {"id": company_id},
        "board": {"name": board_name},
        "status": {"name": status_name}
    }
    
    endpoint = f"{cwpsa_base_url}/service/tickets"
    response = execute_api_call(log, http_client, "post", endpoint, data=ticket_data, integration_name="cw_psa")
    
    if response and response.status_code == 201:
        try:
            ticket = response.json()
            ticket_id = ticket.get("id")
            log.info(f"Successfully created ticket ID [{ticket_id}] with summary: [{summary}]")
            return ticket
        except Exception as e:
            log.exception(e, f"Failed to parse create ticket response for summary: [{summary}]")
            return None
    else:
        log.error(f"Failed to create ticket with summary: [{summary}]")
        return None

def child_ticket(log, http_client, cwpsa_base_url, parent_ticket_id, child_ticket_id):
    log.info(f"Attaching child ticket [{child_ticket_id}] to parent ticket [{parent_ticket_id}]")
    
    endpoint = f"{cwpsa_base_url}/service/tickets/{parent_ticket_id}/attachChildren"
    data = {"childTicketIds": [child_ticket_id]}
    
    response = execute_api_call(log, http_client, "post", endpoint, data=data, integration_name="cw_psa")
    
    if response and response.status_code == 200:
        log.info(f"Successfully attached child ticket [{child_ticket_id}] to parent [{parent_ticket_id}]")
        return True
    else:
        log.error(f"Failed to attach child ticket [{child_ticket_id}] to parent [{parent_ticket_id}]")
        return False

def update_ticket_status(log, http_client, cwpsa_base_url, ticket_id, status_name):
    log.info(f"Updating ticket [{ticket_id}] status to [{status_name}]")
    
    patch_data = [
        {
            "op": "replace",
            "path": "/status/name",
            "value": status_name
        }
    ]
    
    endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_id}"
    response = execute_api_call(log, http_client, "patch", endpoint, data=patch_data, integration_name="cw_psa")
    
    if response and response.status_code == 200:
        log.info(f"Successfully updated ticket [{ticket_id}] status to [{status_name}]")
        return True
    else:
        log.error(f"Failed to update ticket [{ticket_id}] status to [{status_name}]")
        return False

def get_priority_level(priority_name):
    if not priority_name:
        return 999
    match = re.search(r'Priority (\d+)', priority_name, re.IGNORECASE)
    if match:
        return int(match.group(1))
    return 999

def update_ticket_priority(log, http_client, cwpsa_base_url, ticket_id, priority_name):
    log.info(f"Updating ticket [{ticket_id}] priority to [{priority_name}]")
    
    patch_data = [
        {
            "op": "replace",
            "path": "/priority/name",
            "value": priority_name
        }
    ]
    
    endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_id}"
    response = execute_api_call(log, http_client, "patch", endpoint, data=patch_data, integration_name="cw_psa")
    
    if response and response.status_code == 200:
        log.info(f"Successfully updated ticket [{ticket_id}] priority to [{priority_name}]")
        return True
    else:
        log.error(f"Failed to update ticket [{ticket_id}] priority to [{priority_name}]")
        return False

def get_ticket_configurations(log, http_client, cwpsa_base_url, ticket_id):
    log.info(f"Retrieving configurations attached to ticket [{ticket_id}]")
    endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_id}/configurations"
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    if response:
        try:
            configs = response.json()
            log.info(f"Retrieved {len(configs)} configurations for ticket [{ticket_id}]")
            return configs
        except Exception as e:
            log.exception(e, f"Failed to parse configurations response for ticket [{ticket_id}]")
            return []
    log.warning(f"No configurations found for ticket [{ticket_id}]")
    return []

def attach_configuration_to_ticket(log, http_client, cwpsa_base_url, ticket_id, config_id):
    log.info(f"Attaching configuration [{config_id}] to ticket [{ticket_id}]")
    
    data = {"id": config_id}
    endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_id}/configurations"
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
            ticket_number = input.get_value("TicketID_1767833300027")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch ticket number input")
            return

        ticket_number = str(ticket_number).strip() if ticket_number else ""

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number input is required")
            return
        
        try:
            ticket_number = int(ticket_number)
        except ValueError:
            record_result(log, ResultLevel.WARNING, f"Invalid ticket number: [{ticket_number}]")
            return
        
        log.info(f"Processing ticket [{ticket_number}] for KB parent ticket handling")
        
        child_ticket_data = get_ticket_data(log, http_client, cwpsa_base_url, ticket_number)
        if not child_ticket_data:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve ticket data for [{ticket_number}]")
            return
        
        company_id = child_ticket_data.get("company", {}).get("id")
        company_name = child_ticket_data.get("company", {}).get("name")
        child_priority = child_ticket_data.get("priority", {}).get("name", "")
        
        if not company_id:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve company ID from ticket [{ticket_number}]")
            return
        
        log.info(f"Ticket belongs to company: [{company_name}] (ID: {company_id})")
        data_to_log["company_id"] = company_id
        data_to_log["company_name"] = company_name
        data_to_log["child_ticket_id"] = ticket_number
        
        notes = get_ticket_notes(log, http_client, cwpsa_base_url, ticket_number)
        if not notes or len(notes) == 0:
            record_result(log, ResultLevel.WARNING, f"No notes found for ticket [{ticket_number}]")
            return
        
        first_note = notes[0]
        first_note_text = first_note.get("text", "")
        log.info(f"First note text: {first_note_text[:100]}...")
        
        kb_matches = extract_kb_numbers(log, first_note_text)
        
        if not kb_matches or len(kb_matches) == 0:
            record_result(log, ResultLevel.INFO, "No KB numbers found in first note")
            data_to_log["kb_found"] = False
            data_to_log["parent_ticket_id"] = None
            data_to_log["kb_number"] = None
        else:
            kb_number = kb_matches[0]
            data_to_log["kb_number"] = kb_number
            data_to_log["kb_found"] = True
            log.info(f"Processing KB number: [{kb_number}]")
            
            parent_summary = f"Vulnerability Management | {kb_number} | Install Patch"
            parent_ticket = find_ticket_by_summary(log, http_client, cwpsa_base_url, parent_summary, company_id)
            
            if not parent_ticket:
                log.info(f"Parent ticket not found, creating new ticket with summary: [{parent_summary}]")
                parent_ticket = create_ticket(log, http_client, cwpsa_base_url, parent_summary, company_id, "Service Desk", "New")
                if not parent_ticket:
                    record_result(log, ResultLevel.WARNING, f"Failed to create parent ticket for KB [{kb_number}]")
                    return
                data_to_log["parent_ticket_created"] = True
            else:
                log.info(f"Parent ticket already exists: ID [{parent_ticket.get('id')}]")
                data_to_log["parent_ticket_created"] = False
            
            parent_ticket_id = parent_ticket.get("id")
            parent_status = parent_ticket.get("status", {}).get("name", "")
            parent_priority = parent_ticket.get("priority", {}).get("name", "")
            
            data_to_log["parent_ticket_id"] = parent_ticket_id
            data_to_log["parent_status"] = parent_status
            
            child_success = child_ticket(log, http_client, cwpsa_base_url, parent_ticket_id, ticket_number)
            if child_success:
                record_result(log, ResultLevel.SUCCESS, f"Successfully linked ticket [{ticket_number}] as child of [{parent_ticket_id}]")
                data_to_log["childed"] = True
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to link ticket [{ticket_number}] as child of [{parent_ticket_id}]")
                data_to_log["childed"] = False
            
            if parent_status and "closed" in parent_status.lower():
                log.info(f"Parent ticket [{parent_ticket_id}] is closed, reopening to 'Reopened' status")
                reopen_success = update_ticket_status(log, http_client, cwpsa_base_url, parent_ticket_id, "Reopened")
                if reopen_success:
                    record_result(log, ResultLevel.SUCCESS, f"Reopened parent ticket [{parent_ticket_id}]")
                    data_to_log["parent_reopened"] = True
                else:
                    record_result(log, ResultLevel.WARNING, f"Failed to reopen parent ticket [{parent_ticket_id}]")
                    data_to_log["parent_reopened"] = False
            
            child_priority_level = get_priority_level(child_priority)
            parent_priority_level = get_priority_level(parent_priority)
            
            log.info(f"Child priority: [{child_priority}] (level {child_priority_level})")
            log.info(f"Parent priority: [{parent_priority}] (level {parent_priority_level})")
            
            if child_priority_level < parent_priority_level:
                log.info(f"Child has higher priority, updating parent to [{child_priority}]")
                priority_success = update_ticket_priority(log, http_client, cwpsa_base_url, parent_ticket_id, child_priority)
                if priority_success:
                    record_result(log, ResultLevel.SUCCESS, f"Updated parent ticket priority to [{child_priority}]")
                    data_to_log["priority_updated"] = True
                else:
                    data_to_log["priority_updated"] = False
            elif parent_priority_level < child_priority_level:
                log.info(f"Parent has higher priority, updating child to [{parent_priority}]")
                priority_success = update_ticket_priority(log, http_client, cwpsa_base_url, ticket_number, parent_priority)
                if priority_success:
                    record_result(log, ResultLevel.SUCCESS, f"Updated child ticket priority to [{parent_priority}]")
                    data_to_log["priority_updated"] = True
                else:
                    data_to_log["priority_updated"] = False
            else:
                log.info("Priorities are already equal, no update needed")
                data_to_log["priority_updated"] = False
            
            child_configs = get_ticket_configurations(log, http_client, cwpsa_base_url, ticket_number)
            parent_configs = get_ticket_configurations(log, http_client, cwpsa_base_url, parent_ticket_id)
            parent_config_ids = [c.get("id") for c in parent_configs]
            
            configs_attached_count = 0
            for config in child_configs:
                config_id = config.get("id")
                if config_id not in parent_config_ids:
                    log.info(f"Attaching configuration [{config_id}] to parent ticket [{parent_ticket_id}]")
                    attach_success = attach_configuration_to_ticket(log, http_client, cwpsa_base_url, parent_ticket_id, config_id)
                    if attach_success:
                        configs_attached_count += 1
                else:
                    log.info(f"Configuration [{config_id}] already attached to parent ticket")
            
            if configs_attached_count > 0:
                record_result(log, ResultLevel.SUCCESS, f"Attached {configs_attached_count} configurations to parent ticket")
                data_to_log["configs_attached_to_parent"] = configs_attached_count
            else:
                log.info("No new configurations to attach to parent")
                data_to_log["configs_attached_to_parent"] = 0
        
        record_result(log, ResultLevel.SUCCESS, f"Completed KB parent ticket handling for ticket [{ticket_number}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
