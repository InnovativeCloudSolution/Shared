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
bot_name = "CWPSA - Ticket deduplication management"
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

def clean_ticket_summary(log, summary):
    log.info(f"Cleaning ticket summary: [{summary}]")
    original_summary = summary
    summary = summary.strip()
    prefixes_to_remove = [
        r'^RE:\s*',
        r'^Re:\s*',
        r'^re:\s*',
        r'^FW:\s*',
        r'^Fw:\s*',
        r'^fw:\s*',
        r'^FWD:\s*',
        r'^Fwd:\s*',
        r'^fwd:\s*',
        r'^AW:\s*',
        r'^Aw:\s*',
        r'^aw:\s*',
        r'^R:\s*',
        r'^RIF:\s*',
        r'^TR:\s*',
        r'^ENC:\s*',
        r'^ODP:\s*',
        r'^PD:\s*',
        r'^SV:\s*',
        r'^VS:\s*',
        r'^VB:\s*',
        r'^RES:\s*',
        r'^Automatic reply:\s*',
        r'^Out of Office:\s*',
        r'^OOO:\s*',
        r'^\[EXTERNAL\]\s*',
        r'^\[External\]\s*',
        r'^\[external\]\s*',
        r'^\[SPAM\]\s*',
        r'^\[Spam\]\s*',
        r'^\[spam\]\s*'
    ]
    cleaned = summary
    changed = True
    max_iterations = 10
    iteration = 0
    while changed and iteration < max_iterations:
        changed = False
        for prefix_pattern in prefixes_to_remove:
            new_cleaned = re.sub(prefix_pattern, '', cleaned)
            if new_cleaned != cleaned:
                cleaned = new_cleaned.strip()
                changed = True
                break
        iteration += 1
    if cleaned != original_summary:
        log.info(f"Cleaned summary from [{original_summary}] to [{cleaned}]")
    else:
        log.info("Summary already clean, no prefixes removed")
    return cleaned

def get_tickets_by_summary(log, http_client, cwpsa_base_url, cwpsa_base_url_path, summary):
    log.info(f"Retrieving tickets with summary [{summary}]")
    escaped_summary = summary.replace('"', '\\"')
    conditions = f'summary="{escaped_summary}"'
    encoded_conditions = urllib.parse.quote(conditions)
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets?conditions={encoded_conditions}&orderBy=id asc"
    log.info(f"Search conditions: {conditions}")
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    if response:
        try:
            tickets = response.json()
            log.info(f"Found {len(tickets)} tickets with matching summary")
            return tickets
        except Exception as e:
            log.exception(e, f"Failed to parse tickets response")
            return []
    log.info("No matching tickets found")
    return []

def get_board_statuses(log, http_client, cwpsa_base_url, cwpsa_base_url_path, board_id):
    log.info(f"Retrieving statuses for board ID [{board_id}]")
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/boards/{board_id}/statuses"
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    if response:
        try:
            statuses = response.json()
            log.info(f"Retrieved {len(statuses)} statuses for board [{board_id}]")
            return statuses
        except Exception as e:
            log.exception(e, f"Failed to parse board statuses response")
            return []
    log.warning(f"Failed to retrieve statuses for board [{board_id}]")
    return []

def get_status_id_by_name(log, statuses, status_name):
    for status in statuses:
        if status.get("name", "").lower() == status_name.lower():
            status_id = status.get("id")
            log.info(f"Found status ID [{status_id}] for status name [{status_name}]")
            return status_id
    log.warning(f"Status [{status_name}] not found in board statuses")
    return None

def update_ticket_board(log, http_client, cwpsa_base_url, cwpsa_base_url_path, ticket_id, board_id):
    log.info(f"Moving ticket [{ticket_id}] to board ID [{board_id}]")
    patch_data = [
        {
            "op": "replace",
            "path": "/board/id",
            "value": board_id
        }
    ]
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets/{ticket_id}"
    response = execute_api_call(log, http_client, "patch", endpoint, data=patch_data, integration_name="cw_psa")
    if response and response.status_code == 200:
        log.info(f"Successfully moved ticket [{ticket_id}] to board [{board_id}]")
        return True
    else:
        log.error(f"Failed to move ticket [{ticket_id}] to board")
        return False

def update_ticket_status(log, http_client, cwpsa_base_url, cwpsa_base_url_path, ticket_id, status_id):
    log.info(f"Updating ticket [{ticket_id}] to status ID [{status_id}]")
    patch_data = [
        {
            "op": "replace",
            "path": "/status/id",
            "value": status_id
        }
    ]
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets/{ticket_id}"
    response = execute_api_call(log, http_client, "patch", endpoint, data=patch_data, integration_name="cw_psa")
    if response and response.status_code == 200:
        log.info(f"Successfully updated ticket [{ticket_id}] status")
        return True
    else:
        log.error(f"Failed to update ticket [{ticket_id}] status")
        return False

def child_ticket(log, http_client, cwpsa_base_url, cwpsa_base_url_path, parent_ticket_id, child_ticket_id):
    log.info(f"Attaching child ticket [{child_ticket_id}] to parent ticket [{parent_ticket_id}]")
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets/{parent_ticket_id}/attachChildren"
    data = {"childTicketIds": [int(child_ticket_id)]}
    response = execute_api_call(log, http_client, "post", endpoint, data=data, integration_name="cw_psa")
    if response and response.status_code == 200:
        log.info(f"Successfully attached child ticket [{child_ticket_id}] to parent [{parent_ticket_id}]")
        return True
    else:
        log.error(f"Failed to attach child ticket [{child_ticket_id}] to parent [{parent_ticket_id}]")
        return False

def merge_ticket(log, http_client, cwpsa_base_url, cwpsa_base_url_path, parent_ticket_id, merge_ticket_id, status_id):
    log.info(f"Merging ticket [{merge_ticket_id}] into parent ticket [{parent_ticket_id}]")
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets/{parent_ticket_id}/merge"
    data = {"mergeTicketIds": [int(merge_ticket_id)], "status": {"id": status_id}}
    response = execute_api_call(log, http_client, "post", endpoint, data=data, integration_name="cw_psa")
    if response and response.status_code == 200:
        log.info(f"Successfully merged ticket [{merge_ticket_id}] into parent [{parent_ticket_id}]")
        return True
    else:
        log.error(f"Failed to merge ticket [{merge_ticket_id}] into parent [{parent_ticket_id}]")
        return False

def main():
    try:
        try:
            ticket_number = input.get_value("cwTicketId")
            operation = input.get_value("Operation_1772303669651")
            ticket_status = input.get_value("TicketStatus_1772303762954")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        operation = operation.strip() if operation else ""
        ticket_status = ticket_status.strip() if ticket_status else ""

        if not ticket_number or ticket_number == "None":
            record_result(log, ResultLevel.WARNING, "Ticket number is required")
            return
        
        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation is required (Merge or Bundle)")
            return
        
        if operation not in ["Merge", "Bundle"]:
            record_result(log, ResultLevel.WARNING, f"Invalid operation [{operation}]. Must be 'Merge' or 'Bundle'")
            return
        
        if not ticket_status:
            record_result(log, ResultLevel.WARNING, "Ticket status is required")
            return

        log.info(f"Processing ticket [{ticket_number}] for deduplication with operation [{operation}]")
        data_to_log["operation"] = operation
        data_to_log["ticket_status"] = ticket_status
        data_to_log["ticket_summary_original"] = ""
        data_to_log["board_name"] = ""
        data_to_log["ticket_summary_cleaned"] = ""
        data_to_log["duplicates_found"] = 0
        data_to_log["parent_ticket_id"] = 0
        data_to_log["ticket_merge_success"] = False
        data_to_log["ticket_bundle_success"] = False
        data_to_log["ticket_status_updated"] = False

        current_ticket = get_ticket_data(log, http_client, cwpsa_base_url, cwpsa_base_url_path, ticket_number)
        if not current_ticket:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve ticket data for ticket number [{ticket_number}]")
            return

        current_ticket_id = current_ticket.get("id")
        current_summary = current_ticket.get("summary", "")
        current_board_id = current_ticket.get("board", {}).get("id")
        current_board_name = current_ticket.get("board", {}).get("name", "")
        current_parent_id = current_ticket.get("parentTicketId")
        current_merged_parent_id = current_ticket.get("mergedParentTicketId")
        merged_flag = current_ticket.get("mergedFlag", False)

        data_to_log["ticket_id"] = current_ticket_id
        data_to_log["ticket_summary_original"] = current_summary
        data_to_log["board_name"] = current_board_name

        log.info(f"Ticket flags - parentTicketId: [{current_parent_id}], mergedParentTicketId: [{current_merged_parent_id}], mergedFlag: [{merged_flag}]")

        if operation == "Bundle" and current_parent_id:
            log.info(f"Ticket [{ticket_number}] is already a child of parent [{current_parent_id}], skipping Bundle operation")
            record_result(log, ResultLevel.INFO, "Ticket already bundled as child")
            data_to_log["parent_ticket_id"] = current_parent_id
            return

        if operation == "Merge" and (current_merged_parent_id or merged_flag):
            log.info(f"Ticket [{ticket_number}] is already merged (mergedParentTicketId: [{current_merged_parent_id}], mergedFlag: [{merged_flag}]), skipping Merge operation")
            if current_merged_parent_id:
                data_to_log["parent_ticket_id"] = current_merged_parent_id
            record_result(log, ResultLevel.INFO, "Ticket already merged")
            return

        if not current_board_id:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve board ID from ticket [{ticket_number}]")
            return

        log.info(f"Current ticket ID: [{current_ticket_id}], Board: [{current_board_name}] (ID: {current_board_id})")
        log.info(f"Original summary: [{current_summary}]")

        cleaned_summary = clean_ticket_summary(log, current_summary)
        data_to_log["ticket_summary_cleaned"] = cleaned_summary

        matching_tickets = get_tickets_by_summary(log, http_client, cwpsa_base_url, cwpsa_base_url_path, cleaned_summary)
        
        if not matching_tickets or len(matching_tickets) == 0:
            record_result(log, ResultLevel.INFO, f"No duplicate tickets found with summary [{cleaned_summary}]")
            return

        log.info(f"Found {len(matching_tickets)} tickets with matching summary")
        data_to_log["duplicates_found"] = len(matching_tickets)

        parent_ticket = None
        for ticket in matching_tickets:
            ticket_id = ticket.get("id")
            ticket_parent_id = ticket.get("parentTicketId")
            if ticket_parent_id:
                log.info(f"Ticket [{ticket_id}] is already a child of parent [{ticket_parent_id}]")
                parent_ticket_data = get_ticket_data(log, http_client, cwpsa_base_url, cwpsa_base_url_path, ticket_parent_id)
                if parent_ticket_data:
                    parent_ticket = parent_ticket_data
                    log.info(f"Using existing parent ticket [{ticket_parent_id}]")
                    break

        if not parent_ticket:
            parent_ticket = matching_tickets[0]
            log.info(f"No existing parent found, using oldest ticket [{parent_ticket.get('id')}] as parent")

        parent_ticket_id = parent_ticket.get("id")
        parent_board_id = parent_ticket.get("board", {}).get("id")
        parent_board_name = parent_ticket.get("board", {}).get("name", "")
        data_to_log["parent_ticket_id"] = parent_ticket_id

        if current_ticket_id == parent_ticket_id:
            log.info(f"Current ticket [{current_ticket_id}] is already the parent ticket, no action needed")
            record_result(log, ResultLevel.INFO, "Current ticket is the parent ticket")
            return

        log.info(f"Parent ticket board: [{parent_board_name}] (ID: {parent_board_id})")

        if operation == "Merge":
            board_statuses = get_board_statuses(log, http_client, cwpsa_base_url, cwpsa_base_url_path, parent_board_id)
            if not board_statuses:
                record_result(log, ResultLevel.WARNING, f"Failed to retrieve statuses for parent board [{parent_board_name}]")
                return

            target_status_id = get_status_id_by_name(log, board_statuses, ticket_status)
            if not target_status_id:
                record_result(log, ResultLevel.WARNING, f"Status [{ticket_status}] not found in parent board [{parent_board_name}]")
                return

            merge_success = merge_ticket(log, http_client, cwpsa_base_url, cwpsa_base_url_path, parent_ticket_id, current_ticket_id, target_status_id)
            data_to_log["ticket_merge_success"] = merge_success
            data_to_log["ticket_status_updated"] = merge_success
            if merge_success:
                record_result(log, ResultLevel.SUCCESS, f"Successfully merged ticket [{current_ticket_id}] into parent [{parent_ticket_id}]")
                log.info(f"Merge operation set ticket status to [{ticket_status}]")
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to merge ticket [{current_ticket_id}] into parent [{parent_ticket_id}]")
                return

        elif operation == "Bundle":
            child_success = child_ticket(log, http_client, cwpsa_base_url, cwpsa_base_url_path, parent_ticket_id, current_ticket_id)
            data_to_log["ticket_bundle_success"] = child_success
            if child_success:
                record_result(log, ResultLevel.SUCCESS, f"Successfully bundled ticket [{current_ticket_id}] as child of [{parent_ticket_id}]")
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to bundle ticket [{current_ticket_id}] as child of [{parent_ticket_id}]")
                return

            child_ticket_after_bundle = get_ticket_data(log, http_client, cwpsa_base_url, cwpsa_base_url_path, current_ticket_id)
            if not child_ticket_after_bundle:
                record_result(log, ResultLevel.WARNING, "Failed to retrieve child ticket after bundling")
                return

            child_board_id = child_ticket_after_bundle.get("board", {}).get("id")
            child_board_name = child_ticket_after_bundle.get("board", {}).get("name", "")
            log.info(f"Child ticket is now in board: [{child_board_name}] (ID: {child_board_id})")

            if child_board_id != parent_board_id:
                log.info(f"Child ticket board [{child_board_id}] differs from parent board [{parent_board_id}], moving child to parent board")
                move_success = update_ticket_board(log, http_client, cwpsa_base_url, cwpsa_base_url_path, current_ticket_id, parent_board_id)
                if not move_success:
                    log.warning("Failed to move child ticket to parent board")
                else:
                    child_board_id = parent_board_id
                    child_board_name = parent_board_name

            board_statuses = get_board_statuses(log, http_client, cwpsa_base_url, cwpsa_base_url_path, child_board_id)
            if not board_statuses:
                record_result(log, ResultLevel.WARNING, f"Failed to retrieve statuses for child board [{child_board_name}]")
                return

            target_status_id = get_status_id_by_name(log, board_statuses, ticket_status)
            if not target_status_id:
                record_result(log, ResultLevel.WARNING, f"Status [{ticket_status}] not found in child board [{child_board_name}]")
                return

            status_update_success = update_ticket_status(log, http_client, cwpsa_base_url, cwpsa_base_url_path, current_ticket_id, target_status_id)
            data_to_log["ticket_status_updated"] = status_update_success
            if status_update_success:
                record_result(log, ResultLevel.SUCCESS, f"Updated ticket [{current_ticket_id}] status to [{ticket_status}]")
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to update ticket [{current_ticket_id}] status")

        record_result(log, ResultLevel.SUCCESS, f"Completed deduplication for ticket [{ticket_number}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
