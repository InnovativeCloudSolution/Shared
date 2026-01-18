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
nist_nvd2_base_url = "https://api.vulncheck.com"

data_to_log = {}
bot_name = "CWPSA - CVE KB management"
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

def get_ticket_notes(log, http_client, cwpsa_base_url, ticket_number):
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

def extract_kb_numbers(log, text):
    kb_pattern = r'(?i)\bkb\d+\b'
    matches = re.findall(kb_pattern, text)
    if matches:
        log.info(f"Found KB numbers: {matches}")
        return [match.upper() for match in matches]
    log.info("No KB numbers found in text")
    return []

def extract_cve_numbers(log, text):
    cve_pattern = r'CVE-\d+-\d+'
    matches = re.findall(cve_pattern, text, re.IGNORECASE)
    if matches:
        unique_cves = list(set([match.upper() for match in matches]))
        log.info(f"Found CVE numbers: {unique_cves}")
        return unique_cves
    log.info("No CVE numbers found in text")
    return []

def find_ticket_by_summary(log, http_client, cwpsa_base_url, summary, company_id):
    """Search for a ticket with specific summary in the company."""
    log.info(f"Searching for ticket with summary: [{summary}] in company ID [{company_id}]")
    conditions = f'summary="{summary}" AND company/id={company_id}'
    encoded_conditions = urllib.parse.quote(conditions)
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets?conditions={encoded_conditions}"
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
    """Create a new ticket with specified summary, company, board, and status."""
    log.info(f"Creating new ticket with summary: [{summary}] in board [{board_name}]")
    
    ticket_data = {
        "summary": summary,
        "company": {"id": company_id},
        "board": {"name": board_name},
        "status": {"name": status_name}
    }
    
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets"
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
    """Attach a child ticket to a parent ticket."""
    log.info(f"Attaching child ticket [{child_ticket_id}] to parent ticket [{parent_ticket_id}]")
    
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets/{parent_ticket_id}/attachChildren"
    data = {"childTicketIds": [child_ticket_id]}
    
    response = execute_api_call(log, http_client, "post", endpoint, data=data, integration_name="cw_psa")
    
    if response and response.status_code == 200:
        log.info(f"Successfully attached child ticket [{child_ticket_id}] to parent [{parent_ticket_id}]")
        return True
    else:
        log.error(f"Failed to attach child ticket [{child_ticket_id}] to parent [{parent_ticket_id}]")
        return False

def update_ticket_status(log, http_client, cwpsa_base_url, ticket_id, status_name):
    """Update ticket status using PATCH."""
    log.info(f"Updating ticket [{ticket_id}] status to [{status_name}]")
    
    patch_data = [
        {
            "op": "replace",
            "path": "/status/name",
            "value": status_name
        }
    ]
    
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets/{ticket_id}"
    response = execute_api_call(log, http_client, "patch", endpoint, data=patch_data, integration_name="cw_psa")
    
    if response and response.status_code == 200:
        log.info(f"Successfully updated ticket [{ticket_id}] status to [{status_name}]")
        return True
    else:
        log.error(f"Failed to update ticket [{ticket_id}] status to [{status_name}]")
        return False

def get_priority_level(priority_name):
    """Extract priority level from priority name (e.g., 'Priority 1 - blah' -> 1)."""
    if not priority_name:
        return 999
    match = re.search(r'Priority (\d+)', priority_name, re.IGNORECASE)
    if match:
        return int(match.group(1))
    return 999

def update_ticket_priority(log, http_client, cwpsa_base_url, ticket_id, priority_id):
    """Update ticket priority using PATCH."""
    log.info(f"Updating ticket [{ticket_id}] priority to ID [{priority_id}]")
    
    patch_data = [
        {
            "op": "replace",
            "path": "/priority/id",
            "value": priority_id
        }
    ]
    
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets/{ticket_id}"
    response = execute_api_call(log, http_client, "patch", endpoint, data=patch_data, integration_name="cw_psa")
    
    if response and response.status_code == 200:
        log.info(f"Successfully updated ticket [{ticket_id}] priority to ID [{priority_id}]")
        return True
    else:
        log.error(f"Failed to update ticket [{ticket_id}] priority to ID [{priority_id}]")
        return False

def get_ticket_configurations(log, http_client, cwpsa_base_url, ticket_id):
    """Get all configurations attached to a ticket."""
    log.info(f"Retrieving configurations attached to ticket [{ticket_id}]")
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets/{ticket_id}/configurations"
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
    """Attach a configuration to a ticket."""
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

def find_configuration_by_name(log, http_client, cwpsa_base_url, config_name, company_id, config_type="CVE Vulnerability"):
    """Search for a configuration by name and type in a specific company."""
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

def create_configuration(log, http_client, cwpsa_base_url, config_name, company_id, config_type="CVE Vulnerability", status="Active", questions=None):
    """Create a new configuration item."""
    log.info(f"Creating new configuration with name: [{config_name}], type: [{config_type}]")
    
    config_data = {
        "name": config_name,
        "type": {"name": config_type},
        "company": {"id": company_id},
        "status": {"name": status}
    }
    
    if questions:
        config_data["questions"] = questions
        log.info(f"Including {len(questions)} questions in initial POST")
    
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/company/configurations"
    response = execute_api_call(log, http_client, "post", endpoint, data=config_data, integration_name="cw_psa")
    
    if response and response.status_code == 201:
        try:
            config = response.json()
            config_id = config.get("id")
            log.info(f"Successfully created configuration ID [{config_id}] with name: [{config_name}]")
            return config
        except Exception as e:
            log.exception(e, f"Failed to parse create configuration response for name: [{config_name}]")
            return None
    else:
        log.error(f"Failed to create configuration with name: [{config_name}]")
        return None

def get_cve_data_from_nist(log, cve_number):
    """Retrieve CVE data from NIST NVD2 API via VulnCheck."""
    log.info(f"Retrieving CVE data for [{cve_number}] from NIST NVD2")
    
    endpoint = f"{nist_nvd2_base_url}/v3/index/nist-nvd2?cve={cve_number}"
    
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="custom_wf_apikey")
    
    if response:
        try:
            cve_data = response.json()
            log.info(f"Successfully retrieved CVE data for [{cve_number}]")
            return cve_data
        except Exception as e:
            log.exception(e, f"Failed to parse CVE data for [{cve_number}]")
    else:
        log.warning(f"Failed to retrieve CVE data for [{cve_number}]")
    
    return None

def update_configuration_questions(log, http_client, cwpsa_base_url, config_id, questions_data):
    """
    Update configuration questions with CVE data.
    questions_data should be a list of dicts with 'questionId' and 'answer'.
    
    USER NOTE: You need to provide the question mappings here.
    For example:
    questions_data = [
        {"questionId": 123, "answer": "Some value from CVE data"},
        {"questionId": 456, "answer": "Another value from CVE data"}
    ]
    """
    log.info(f"Updating configuration [{config_id}] questions with CVE data")
    
    patch_data = [
        {
            "op": "replace",
            "path": "/questions",
            "value": questions_data
        }
    ]
    
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/company/configurations/{config_id}"
    response = execute_api_call(log, http_client, "patch", endpoint, data=patch_data, integration_name="cw_psa")
    
    if response and response.status_code == 200:
        log.info(f"Successfully updated configuration [{config_id}] questions")
        return True
    else:
        log.error(f"Failed to update configuration [{config_id}] questions")
        return False

def main():
    try:
        # Get input
        try:
            ticket_number = input.get_value("cwTicketId")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch ticket number input")
            return

        if not ticket_number or ticket_number == "None":
            record_result(log, ResultLevel.WARNING, "Ticket number input is required")
            return
        
        log.info(f"Processing ticket [{ticket_number}] for CVE/KB handling")
        
        # Step 1: Get ticket data
        child_ticket_data = get_ticket_data(log, http_client, cwpsa_base_url, cwpsa_base_url_path, ticket_number)
        if not child_ticket_data:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve ticket data for [{ticket_number}]")
            return
        
        # Guard: Check if ticket has already been processed (is a child ticket)
        parent_ticket_id_check = child_ticket_data.get("parentTicketId")
        if parent_ticket_id_check:
            log.info(f"Ticket [{ticket_number}] is already a child ticket (parent: {parent_ticket_id_check}), skipping processing to prevent loop")
            record_result(log, ResultLevel.INFO, "Ticket already processed - is a child ticket")
            data_to_log["already_processed"] = True
            data_to_log["reason"] = "Already a child ticket"
            return
        
        company_id = child_ticket_data.get("company", {}).get("id")
        company_name = child_ticket_data.get("company", {}).get("name")
        child_priority_name = child_ticket_data.get("priority", {}).get("name", "")
        child_priority_id = child_ticket_data.get("priority", {}).get("id")
        
        if not company_id:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve company ID from ticket [{ticket_number}]")
            return
        
        log.info(f"Ticket belongs to company: [{company_name}] (ID: {company_id})")
        data_to_log["company_name"] = company_name
        data_to_log["child_ticket_id"] = ticket_number
        
        # Step 2: Get ticket notes and find first note
        notes = get_ticket_notes(log, http_client, cwpsa_base_url, ticket_number)
        if not notes or len(notes) == 0:
            record_result(log, ResultLevel.WARNING, f"No notes found for ticket [{ticket_number}]")
            return
        
        first_note = notes[0]
        first_note_text = first_note.get("text", "")
        log.info(f"First note text: {first_note_text[:100]}...")
        
        # Step 3: Match KB regex in first note
        kb_matches = extract_kb_numbers(log, first_note_text)
        parent_ticket_id = None
        
        if not kb_matches or len(kb_matches) == 0:
            record_result(log, ResultLevel.INFO, "No KB numbers found in first note")
            data_to_log["kb_found"] = False
        else:
            kb_number = kb_matches[0]
            data_to_log["kb_number"] = kb_number
            data_to_log["kb_found"] = True
            log.info(f"Processing KB number: [{kb_number}]")
            
            # Step 4: Look for or create parent ticket
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
            parent_priority_name = parent_ticket.get("priority", {}).get("name", "")
            parent_priority_id = parent_ticket.get("priority", {}).get("id")
            
            data_to_log["parent_ticket_id"] = parent_ticket_id
            data_to_log["parent_status"] = parent_status
            
            # Step 5: Child the input ticket to parent ticket
            child_success = child_ticket(log, http_client, cwpsa_base_url, parent_ticket_id, ticket_number)
            if child_success:
                record_result(log, ResultLevel.SUCCESS, f"Successfully linked ticket [{ticket_number}] as child of [{parent_ticket_id}]")
                data_to_log["childed"] = True
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to link ticket [{ticket_number}] as child of [{parent_ticket_id}]")
                data_to_log["childed"] = False
            
            # Step 6: If parent is closed, reopen it
            if parent_status and "closed" in parent_status.lower():
                log.info(f"Parent ticket [{parent_ticket_id}] is closed, reopening to 'Reopened' status")
                reopen_success = update_ticket_status(log, http_client, cwpsa_base_url, parent_ticket_id, "Reopened")
                if reopen_success:
                    record_result(log, ResultLevel.SUCCESS, f"Reopened parent ticket [{parent_ticket_id}]")
                    data_to_log["parent_reopened"] = True
                else:
                    record_result(log, ResultLevel.WARNING, f"Failed to reopen parent ticket [{parent_ticket_id}]")
                    data_to_log["parent_reopened"] = False
            
            # Step 7: Compare priorities and set to higher priority
            child_priority_level = get_priority_level(child_priority_name)
            parent_priority_level = get_priority_level(parent_priority_name)
            
            log.info(f"Child priority: [{child_priority_name}] (level {child_priority_level})")
            log.info(f"Parent priority: [{parent_priority_name}] (level {parent_priority_level})")
            
            if child_priority_level < parent_priority_level:
                log.info(f"Child has higher priority, updating parent to [{child_priority_name}] (ID: {child_priority_id})")
                priority_success = update_ticket_priority(log, http_client, cwpsa_base_url, parent_ticket_id, child_priority_id)
                if priority_success:
                    record_result(log, ResultLevel.SUCCESS, f"Updated parent ticket priority to [{child_priority_name}]")
                    data_to_log["priority_updated"] = True
                else:
                    data_to_log["priority_updated"] = False
            elif parent_priority_level < child_priority_level:
                log.info(f"Parent has higher priority, updating child to [{parent_priority_name}] (ID: {parent_priority_id})")
                priority_success = update_ticket_priority(log, http_client, cwpsa_base_url, ticket_number, parent_priority_id)
                if priority_success:
                    record_result(log, ResultLevel.SUCCESS, f"Updated child ticket priority to [{parent_priority_name}]")
                    data_to_log["priority_updated"] = True
                else:
                    data_to_log["priority_updated"] = False
            else:
                log.info("Priorities are already equal, no update needed")
                data_to_log["priority_updated"] = False
            
            # Step 8: Get configurations from child ticket and attach to parent
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
        
        # Step 9: Get all notes from child ticket and extract CVEs
        all_cves = []
        for note in notes:
            note_text = note.get("text", "")
            cves_in_note = extract_cve_numbers(log, note_text)
            all_cves.extend(cves_in_note)
        
        # Remove duplicates
        all_cves = list(set(all_cves))
        data_to_log["cves_found"] = all_cves
        data_to_log["cve_count"] = len(all_cves)
        
        if len(all_cves) == 0:
            record_result(log, ResultLevel.INFO, "No CVE numbers found in ticket notes")
        else:
            record_result(log, ResultLevel.INFO, f"Found {len(all_cves)} unique CVE numbers: {', '.join(all_cves)}")
            
            # Step 10: Process each CVE
            cve_configs_created = 0
            cve_configs_found = 0
            
            for cve_number in all_cves:
                log.info(f"Processing CVE: [{cve_number}]")
                
                # Check if CVE config exists
                cve_config = find_configuration_by_name(log, http_client, cwpsa_base_url, cve_number, company_id, "CVE Vulnerability")
                
                if not cve_config:
                    # Get CVE data from NIST first
                    log.info(f"Retrieving CVE data for [{cve_number}] before creating configuration")
                    cve_data = get_cve_data_from_nist(log, cve_number)
                    
                    questions_data = None
                    if cve_data:
                        log.info(f"CVE data structure keys: {list(cve_data.keys())}")
                        
                        q1_cve_id = cve_number
                        q2_publish_date = ""
                        q3_last_modified = ""
                        q4_cvss_version = ""
                        q5_cvss_score = ""
                        q6_advisory_url = ""
                        q7_advisory_sources = ""
                        q8_advisory_tags = ""
                        q9_description = "No description available"
                        q10_nvd_status = ""
                        
                        if "data" in cve_data and isinstance(cve_data["data"], list) and len(cve_data["data"]) > 0:
                            cve_details = cve_data["data"][0]
                            log.info(f"CVE details keys: {list(cve_details.keys())}")
                            
                            q1_cve_id = cve_details.get("id", cve_number)
                            
                            published_raw = cve_details.get("published", "")
                            q2_publish_date = published_raw.split("T")[0] if published_raw else ""
                            
                            modified_raw = cve_details.get("lastModified", "")
                            q3_last_modified = modified_raw.split("T")[0] if modified_raw else ""
                            
                            q10_nvd_status = cve_details.get("vulnStatus", "")
                            
                            if "descriptions" in cve_details and isinstance(cve_details["descriptions"], list) and len(cve_details["descriptions"]) > 0:
                                q9_description = cve_details["descriptions"][0].get("value", "No description available")
                                log.info(f"Extracted description: [{q9_description[:100]}...]")
                            
                            if "metrics" in cve_details and "cvssMetricV31" in cve_details["metrics"]:
                                metrics = cve_details["metrics"]["cvssMetricV31"]
                                if isinstance(metrics, list) and len(metrics) > 0:
                                    cvss_data = metrics[0].get("cvssData", {})
                                    base_score = cvss_data.get("baseScore", "")
                                    cvss_version = cvss_data.get("version", "")
                                    if base_score:
                                        q5_cvss_score = str(base_score)
                                    if cvss_version:
                                        q4_cvss_version = f"CVSS Version {cvss_version}"
                                    log.info(f"Extracted CVSS: Version [{q4_cvss_version}], Score [{q5_cvss_score}]")
                            
                            if "references" in cve_details and isinstance(cve_details["references"], list) and len(cve_details["references"]) > 0:
                                first_ref = cve_details["references"][0]
                                q6_advisory_url = first_ref.get("url", "")
                                q7_advisory_sources = first_ref.get("source", "")
                                tags = first_ref.get("tags", [])
                                if tags and isinstance(tags, list):
                                    q8_advisory_tags = ", ".join(tags)
                                log.info(f"Extracted advisory: URL [{q6_advisory_url}], Source [{q7_advisory_sources}], Tags [{q8_advisory_tags}]")
                        
                        questions_data = [
                            {"questionId": 484, "answer": q1_cve_id},
                            {"questionId": 485, "answer": q2_publish_date},
                            {"questionId": 486, "answer": q3_last_modified},
                            {"questionId": 487, "answer": q4_cvss_version},
                            {"questionId": 488, "answer": q5_cvss_score},
                            {"questionId": 489, "answer": q6_advisory_url},
                            {"questionId": 490, "answer": q7_advisory_sources},
                            {"questionId": 491, "answer": q8_advisory_tags},
                            {"questionId": 492, "answer": q9_description},
                            {"questionId": 493, "answer": q10_nvd_status}
                        ]
                        log.info(f"Mapped all CVE fields to 10 questions")
                    else:
                        log.warning(f"No CVE data retrieved for [{cve_number}] - using default values")
                        questions_data = [
                            {"questionId": 484, "answer": cve_number},
                            {"questionId": 485, "answer": ""},
                            {"questionId": 486, "answer": ""},
                            {"questionId": 487, "answer": ""},
                            {"questionId": 488, "answer": ""},
                            {"questionId": 489, "answer": ""},
                            {"questionId": 490, "answer": ""},
                            {"questionId": 491, "answer": ""},
                            {"questionId": 492, "answer": "No description available"},
                            {"questionId": 493, "answer": ""}
                        ]
                    
                    # Create CVE config with all 10 questions
                    log.info(f"Creating configuration item for CVE: [{cve_number}]")
                    cve_config = create_configuration(log, http_client, cwpsa_base_url, cve_number, company_id, "CVE Vulnerability", "Active", questions_data)
                    if cve_config:
                        cve_configs_created += 1
                        record_result(log, ResultLevel.SUCCESS, f"Created configuration for CVE: [{cve_number}]")
                    else:
                        record_result(log, ResultLevel.WARNING, f"Failed to create configuration for CVE: [{cve_number}]")
                        continue
                else:
                    cve_configs_found += 1
                    log.info(f"Configuration already exists for CVE: [{cve_number}] (ID: {cve_config.get('id')})")
                
                # Attach CVE config to BOTH tickets (child and parent if exists)
                cve_config_id = cve_config.get("id")
                
                # Attach to child ticket
                child_configs = get_ticket_configurations(log, http_client, cwpsa_base_url, ticket_number)
                child_config_ids = [c.get("id") for c in child_configs]
                
                if cve_config_id not in child_config_ids:
                    attach_success = attach_configuration_to_ticket(log, http_client, cwpsa_base_url, ticket_number, cve_config_id)
                    if attach_success:
                        log.info(f"Attached CVE config [{cve_config_id}] to child ticket [{ticket_number}]")
                else:
                    log.info(f"CVE config [{cve_config_id}] already attached to child ticket")
                
                # Attach to parent ticket (if KB was found and parent exists)
                if kb_matches and len(kb_matches) > 0 and parent_ticket_id:
                    parent_configs = get_ticket_configurations(log, http_client, cwpsa_base_url, parent_ticket_id)
                    parent_config_ids = [c.get("id") for c in parent_configs]
                    
                    if cve_config_id not in parent_config_ids:
                        attach_success = attach_configuration_to_ticket(log, http_client, cwpsa_base_url, parent_ticket_id, cve_config_id)
                        if attach_success:
                            log.info(f"Attached CVE config [{cve_config_id}] to parent ticket [{parent_ticket_id}]")
                    else:
                        log.info(f"CVE config [{cve_config_id}] already attached to parent ticket")
            
            data_to_log["cve_configs_created"] = cve_configs_created
            data_to_log["cve_configs_found"] = cve_configs_found
            record_result(log, ResultLevel.SUCCESS, f"Processed {len(all_cves)} CVEs: {cve_configs_created} created, {cve_configs_found} already existed")
        
        record_result(log, ResultLevel.SUCCESS, f"Completed CVE/KB processing for ticket [{ticket_number}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
