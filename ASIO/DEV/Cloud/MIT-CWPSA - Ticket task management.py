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

cwpsa_base_url = "https://au.myconnectwise.net/v4_6_release/apis/3.0"
msgraph_base_url = "https://graph.microsoft.com/v1.0"
msgraph_base_url_beta = "https://graph.microsoft.com/beta"
vault_name = "mit-azu1-prod1-akv1"

data_to_log = {}
bot_name = "MIT-CWPSA - Ticket task management"
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

def add_task_to_ticket(log, http_client, cwpsa_base_url, ticket_number, notes, resolution=None, closed_flag=False):
    endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}/tasks"

    body = {
        "notes": notes,
        "closedFlag": closed_flag
    }
    if resolution:
        body["resolution"] = resolution

    response = execute_api_call(log, http_client, "post", endpoint, data=body, integration_name="cw_psa")
    return response and response.status_code in (200, 201)

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1756861315488")
            tasks = input.get_value("Tasks_1756861353228")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        tasks = tasks.strip() if tasks else ""
        resolution = "Completed manual task"
        closed_flag = False

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number is required")
            return

        tasks_list = [task.strip() for task in tasks.split(",") if task.strip()]
        if not tasks_list:
            record_result(log, ResultLevel.WARNING, "No valid task notes provided")
            return

        failures = []

        for task in tasks_list:
            success = add_task_to_ticket(log, http_client, cwpsa_base_url, ticket_number, task, resolution, closed_flag)
            if success:
                record_result(log, ResultLevel.SUCCESS, f"Task [{task}] added to ticket [{ticket_number}]")
            if not success:
                failures.append(task)

        if failures:
            for failure in failures:
                record_result(log, ResultLevel.WARNING, f"Task [{failure}] failed to add to ticket [{ticket_number}]")

    except Exception:
        record_result(log, ResultLevel.WARNING, "Process failed")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()