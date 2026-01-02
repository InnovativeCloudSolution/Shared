import sys
import random
import os
import time
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

cwpsa_base_url = "https://aus.myconnectwise.net/v4_6_release/apis/3.0"

data_to_log = {}
bot_name = "CWPSA - Config Attacher"
log.info("Static variables set")

def record_result(log, level, message):
    log.result_message(level, f"[{bot_name}]: {message}")
    if level == ResultLevel.WARNING:
        data_to_log["status_result"] = "Fail"
    elif level == ResultLevel.SUCCESS:
        if "status_result" not in data_to_log or data_to_log["status_result"] != "Fail":
            data_to_log["status_result"] = "Success"

def execute_api_call(log, http_client, method, endpoint, data=None, retries=5, integration_name=None):
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
                log.error("Integration name is required for API calls")
                return None

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

def get_ticket_configurations(log, http_client, cwpsa_base_url, ticket_id):
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
            child_ticket_id = input.get_value("ChildTicketID_1765965840181")
            parent_ticket_id = input.get_value("ParentTicketID_1765965841865")
            config_id = input.get_value("ConfigID_1765965843374")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch required input values")
            return

        child_ticket_id = child_ticket_id.strip() if child_ticket_id else ""
        parent_ticket_id = parent_ticket_id.strip() if parent_ticket_id else ""
        config_id = config_id.strip() if config_id else ""
        
        if not child_ticket_id:
            record_result(log, ResultLevel.WARNING, "Child Ticket ID input is required")
            return
        
        if not config_id:
            record_result(log, ResultLevel.WARNING, "Configuration ID input is required")
            return
        
        try:
            config_id = int(config_id)
        except:
            record_result(log, ResultLevel.WARNING, f"Invalid configuration ID: [{config_id}]")
            return
        
        ticket_ids = [child_ticket_id]
        if parent_ticket_id:
            ticket_ids.append(parent_ticket_id)
            log.info(f"Processing child ticket [{child_ticket_id}] and parent ticket [{parent_ticket_id}]")
        else:
            log.info(f"Processing child ticket [{child_ticket_id}] only (no parent ticket provided)")
        
        if not ticket_ids:
            record_result(log, ResultLevel.WARNING, "No valid ticket IDs provided")
            return
        
        log.info(f"Attaching configuration [{config_id}] to {len(ticket_ids)} ticket(s)")
        data_to_log["config_id"] = config_id
        data_to_log["child_ticket_id"] = child_ticket_id
        data_to_log["parent_ticket_id"] = parent_ticket_id if parent_ticket_id else None
        data_to_log["ticket_ids"] = ticket_ids
        
        attached_count = 0
        already_attached_count = 0
        failed_count = 0
        
        for ticket_id in ticket_ids:
            try:
                ticket_id_int = int(ticket_id)
            except:
                log.warning(f"Skipping invalid ticket ID: [{ticket_id}]")
                failed_count += 1
                continue
            
            existing_configs = get_ticket_configurations(log, http_client, cwpsa_base_url, ticket_id_int)
            existing_config_ids = [c.get("id") for c in existing_configs]
            
            if config_id in existing_config_ids:
                log.info(f"Configuration [{config_id}] already attached to ticket [{ticket_id}]")
                already_attached_count += 1
            else:
                attach_success = attach_configuration_to_ticket(log, http_client, cwpsa_base_url, ticket_id_int, config_id)
                if attach_success:
                    attached_count += 1
                else:
                    failed_count += 1
        
        data_to_log["attached_count"] = attached_count
        data_to_log["already_attached_count"] = already_attached_count
        data_to_log["failed_count"] = failed_count
        
        if attached_count > 0:
            record_result(log, ResultLevel.SUCCESS, f"Attached configuration to {attached_count} ticket(s)")
        if already_attached_count > 0:
            record_result(log, ResultLevel.INFO, f"Configuration already attached to {already_attached_count} ticket(s)")
        if failed_count > 0:
            record_result(log, ResultLevel.WARNING, f"Failed to attach configuration to {failed_count} ticket(s)")
        
        log.info(f"Completed configuration attachment")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
