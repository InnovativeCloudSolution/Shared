import sys
import traceback
import json
import random
import re
import subprocess
import os
import io
import base64
import hashlib
import time
import urllib.parse
import string
import requests
import pandas as pd
from datetime import datetime, timedelta, timezone
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

cwpsa_base_url = "https://au.myconnectwise.net/v4_6_release/apis/3.0"
data_to_log = {}
onboarding_webhook_url = "https://webhook-test.com/e3a4085a08f22b5d05e39e85019bae8a"  # Replace when in production with "https://au.webhook.myconnectwise.net/HbSHrnJJbexELm2IDb0OW85moFmd1VwQzpVBVzIqtNKFGdjrlzxOqNThFH-kbyccLctvZA=="
offboarding_webhook_url = "https://webhook-test.com/d73a09e47509d8b62cd8e5b360a6b02c" # Replace when in production with "https://au.webhook.myconnectwise.net/HeGF_HNLPL1EfjHZUr0OCp9ioFmY0VMQms5NU2R_stbTGN7gtLytQ0ABUa59JwbSg1pMMQ=="
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
                response = getattr(http_client.third_party_integration(integration_name), method)(url=endpoint, json=data) if data else getattr(http_client.third_party_integration(integration_name), method)(url=endpoint)
            else:
                request_args = {"url": endpoint}
                if params:
                    request_args["params"] = params
                if headers:
                    request_args["headers"] = headers
                if data:
                    if headers and headers.get("Content-Type") == "application/x-www-form-urlencoded":
                        request_args["data"] = data
                    else:
                        request_args["json"] = data
                response = getattr(requests, method)(**request_args)

            if response.status_code in [200, 204]:
                return response
            if response.status_code in [429, 503]:
                retry_after = response.headers.get("Retry-After")
                wait_time = int(retry_after) if retry_after else base_delay * (2 ** attempt) + random.uniform(0, 3)
                log.warning(f"Rate limit exceeded. Retrying in {wait_time:.2f} seconds")
                time.sleep(wait_time)
            elif response.status_code == 404:
                log.warning(f"Skipping non-existent resource [{endpoint}]")
                return None
            else:
                log.error(f"API request failed Status: {response.status_code}, Response: {response.text}")
                return response
        except Exception as e:
            log.exception(e, f"Exception during API call to {endpoint}")
            return None
    return None

def main():
    try:
        log.info("Starting main function")
        try:
            cwpsa_ticket = input.get_value("TicketNumber_1748410807092")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            log.result_data(data_to_log)
            return

        cwpsa_ticket = cwpsa_ticket.strip() if cwpsa_ticket else ""
        log.info(f"Ticket Number: {cwpsa_ticket}")

        if cwpsa_ticket:
            # Set CWPSA base URL
            endpoint = f"{cwpsa_base_url}/service/tickets/{cwpsa_ticket}"
            # Get the ticket information
            response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
            if response.status_code == 200:
                data = response.json()
                # Pull the summary from the ticket data
                ticket_summary = data.get("summary")
                log.info(f"Summary for ticket [{cwpsa_ticket}] is [{ticket_summary}]")
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to fetch ticket summary for ticket [{cwpsa_ticket}]")
                data_to_log["TicketNumber"] = cwpsa_ticket
                log.result_data(data_to_log)
                return

            # Isolate the operation from the ticket summary
            result = re.search(r"User (O\w+):.+", ticket_summary)
            operation = result.group(1) if result else None

            # Determine if the operation is Onboarding or Offboarding and set the appropriate webhook URL
            if operation == "Onboarding":
                log.info("Onboarding operation detected")
                response = execute_api_call(log, http_client, "post", onboarding_webhook_url, data={"cwpsa_ticket": cwpsa_ticket})
                log.info(f"Response from onboarding webhook: {response.text}")
                if response and response.status_code == 200:
                    record_result(log, ResultLevel.SUCCESS, "Onboarding retry executed successfully")
                    data_to_log["TicketNumber"] = cwpsa_ticket
                    data_to_log["Operation"] = "Onboarding"
                    log.result_data(data_to_log)

            elif operation == "Offboarding":
                log.info("Offboarding operation detected")
                response = execute_api_call(log, http_client, "post", offboarding_webhook_url, data={"cwpsa_ticket": cwpsa_ticket})
                log.info(f"Response from onboarding webhook: {response.text}")
                if response and response.status_code == 200:
                    record_result(log, ResultLevel.SUCCESS, "Offboarding retry executed successfully")
                    data_to_log["TicketNumber"] = cwpsa_ticket
                    data_to_log["Operation"] = "Onboarding"
                    log.result_data(data_to_log)

        else:
            record_result(log, ResultLevel.WARNING, "Either Access Token or Ticket Number is required")
            data_to_log["TicketNumber"] = cwpsa_ticket
            log.result_data(data_to_log)
            return

    except Exception:
        record_result(log, ResultLevel.WARNING, "Unexpected error occurred")
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()