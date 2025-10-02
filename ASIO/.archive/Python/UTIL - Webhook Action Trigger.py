import sys
import json
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

data_to_log = {}
bot_name = "[UTIL - Webhook Action Trigger]"
azure_cosmodb_helper_url = "https://1ce978e3-0165-4625-bc15-bdcdfaeb9c7e.webhook.ae.azure-automation.net/webhooks?token=xUFmMKhtDzQVjGtnqDE6Xd8zMkBi2Zabt9BnzGkCFro%3d"
log.info("Static variables set")

def record_result(log, level, message):
    log.result_message(level, f"{bot_name}: {message}")

    if level == ResultLevel.WARNING:
        data_to_log["Result"] = "Fail"
    elif level == ResultLevel.SUCCESS:
        if "Result" not in data_to_log or data_to_log["Result"] != "Fail":
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

def main():
    try:
        try:
            operation = input.get_value("Operation_1751427590883")
            ticket_number = input.get_value("TicketNumber_1751427569586")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        operation = operation.strip() if operation else ""

        log.info(f"Received input ticket number = [{ticket_number}]")
        log.info(f"Requested operation = [{operation}]")

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number is empty or invalid")
            return
        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation value is missing or invalid")
            return

        if operation == "Schedule Offboarding":
            log.info(f"Offboarding user for ticket [{ticket_number}]")
            payload = {
                "cwpsa_ticket": int(ticket_number),
                "action": "get",
                "request_type": "user_offboarding"
            }
            log.info(f"Offboarding payload: {json.dumps(payload)}")
            success = execute_api_call(log, http_client, "post", azure_cosmodb_helper_url, data=payload)
            if success and success.status_code == 200:
                log.info(f"Offboarding action for ticket [{ticket_number}] sent successfully.")
                record_result(log, ResultLevel.SUCCESS, f"Offboarding action for ticket [{ticket_number}] completed successfully")
            else:
                log.error(f"Failed to send offboarding action for ticket [{ticket_number}]. Response: {success.text if success else 'No response'}")
                record_result(log, ResultLevel.WARNING, f"Failed to offboard user for ticket [{ticket_number}]")

        elif operation == "Schedule Onboarding":
            log.info(f"Onboarding user for ticket [{ticket_number}]")
            payload = {
                "cwpsa_ticket": int(ticket_number),
                "action": "get",
                "request_type": "user_onboarding"
            }
            log.info(f"Onboarding payload: {json.dumps(payload)}")
            success = execute_api_call(log, http_client, "post", azure_cosmodb_helper_url, data=payload)
            if success and success.status_code == 200:
                log.info(f"Onboarding action for ticket [{ticket_number}] sent successfully.")
                record_result(log, ResultLevel.SUCCESS, f"Onboarding action for ticket [{ticket_number}] completed successfully")
            else:
                log.error(f"Failed to send onboarding action for ticket [{ticket_number}]. Response: {success.text if success else 'No response'}")
                record_result(log, ResultLevel.WARNING, f"Failed to onboard user for ticket [{ticket_number}]")

        else:
            record_result(log, ResultLevel.WARNING, f"Unknown operation [{operation}]")
            return

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()