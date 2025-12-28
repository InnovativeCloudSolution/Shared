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

cwpsa_base_url = "https://aus.myconnectwise.net/v4_6_release/apis/3.0"

data_to_log = {}
bot_name = "CWPSA - Ticket custom field updater"
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

def get_ticket_custom_fields(log, http_client, cwpsa_base_url, ticket_number):
    log.info(f"Retrieving custom fields for ticket [{ticket_number}]")
    endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}"
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    if response:
        ticket_data = response.json()
        custom_fields = ticket_data.get("customFields", [])
        log.info(f"Found [{len(custom_fields)}] custom fields on ticket [{ticket_number}]")
        return custom_fields
    return []

def find_custom_field_by_name(log, custom_fields, field_name):
    log.info(f"Searching for custom field with caption [{field_name}]")
    for field in custom_fields:
        if field.get("caption", "").lower() == field_name.lower():
            field_id = field.get("id")
            current_value = field.get("value")
            log.info(f"Found custom field [{field_name}] with ID [{field_id}], current value: [{current_value}]")
            return field_id, current_value
    log.warning(f"Custom field [{field_name}] not found")
    return None, None

def update_ticket_custom_field(log, http_client, cwpsa_base_url, ticket_number, field_id, new_value):
    log.info(f"Updating custom field ID [{field_id}] on ticket [{ticket_number}] to [{new_value}]")
    endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}"
    patch_data = [
        {
            "op": "replace",
            "path": "customFields",
            "value": [
                {
                    "id": field_id,
                    "value": new_value
                }
            ]
        }
    ]
    response = execute_api_call(log, http_client, "patch", endpoint, data=patch_data, integration_name="cw_psa")
    if response and 200 <= response.status_code < 300:
        log.info(f"Successfully updated custom field [{field_id}] on ticket [{ticket_number}]")
        return True
    else:
        log.error(f"Failed to update custom field [{field_id}] on ticket [{ticket_number}]TicketNumber_xxxxxxxxxxxxx
def main():
    try:
        try:
            tiCustomFieldName_xxxxxxxxxxxxx("TicketNumber_1766001954665")
            cNewValue_xxxxxxxxxxxxxinput.get_value("CustomFieldName_1766001955978")
            new_value = input.get_value("Value_1766001957111")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        custom_field_name = custom_field_name.strip() if custom_field_name else ""
        new_value = new_value.strip() if new_value else ""

        if not new_value:
            new_value = None
            log.info("New value is blank - will set custom field to null")

        log.info(f"Ticket Number = [{ticket_number}]")
        log.info(f"Custom Field Name = [{custom_field_name}]")
        log.info(f"New Value = [{new_value}]")

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number is required but missing")
            return
        if not custom_field_name:
            record_result(log, ResultLevel.WARNING, "Custom field name is required but missing")
            return

        custom_fields = get_ticket_custom_fields(log, http_client, cwpsa_base_url, ticket_number)
        if not custom_fields:
            record_result(log, ResultLevel.WARNING, f"No custom fields found on ticket [{ticket_number}]")
            return

        field_id, current_value = find_custom_field_by_name(log, custom_fields, custom_field_name)
        if not field_id:
            record_result(log, ResultLevel.WARNING, f"Custom field [{custom_field_name}] not found on ticket [{ticket_number}]")
            return

        data_to_log["CustomFieldID"] = field_id
        data_to_log["PreviousValue"] = current_value
        data_to_log["NewValue"] = new_value if new_value is not None else "null"

        if update_ticket_custom_field(log, http_client, cwpsa_base_url, ticket_number, field_id, new_value):
            action = "cleared" if new_value is None else f"updated to [{new_value}]"
            record_result(log, ResultLevel.SUCCESS, f"Successfully {action} custom field [{custom_field_name}] on ticket [{ticket_number}] (previous value: [{current_value}])")
        else:
            record_result(log, ResultLevel.WARNING, f"Failed to update custom field [{custom_field_name}] on ticket [{ticket_number}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()