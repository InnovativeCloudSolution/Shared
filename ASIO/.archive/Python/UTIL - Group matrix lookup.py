import sys
import random
import os
import io
import time
import requests
import pandas as pd

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from cw_rpa import Logger, Input, HttpClient, ResultLevel

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

cwpsa_base_url = "https://au.myconnectwise.net/v4_6_release/apis/3.0"
msgraph_base_url = "https://graph.microsoft.com/v1.0"
vault_name = "mit-azu1-prod1-akv1"
log.info("Static variables set")

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

def get_company_identifier_from_ticket(log, http_client, cwpsa_base_url, ticket_number):
    log.info(f"Retrieving company identifier for ticket [{ticket_number}]")
    endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}"

    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")

    if response:
        if response.status_code == 200:
            data = response.json()
            company_identifier = data.get("company", {}).get("identifier", "")
            if company_identifier:
                log.info(f"Company identifier for ticket [{ticket_number}] is [{company_identifier}]")
                return company_identifier
            else:
                log.error(f"Company identifier not found in response for ticket [{ticket_number}]")
        else:
            log.error(f"Failed to retrieve company identifier for ticket [{ticket_number}] Status: {response.status_code}, Body: {response.text}")
    else:
        log.error(f"Failed to retrieve company identifier for ticket [{ticket_number}]: No response received")

    return ""

def get_matrix_object_id(log, http_client, company_identifier, lookup_column, lookup_value, return_columns):
    try:
        file_url = f"https://mitazu1pubfilestore.blob.core.windows.net/automation/{company_identifier}-Group-Matrix.csv"
        log.info(f"Downloading matrix file from [{file_url}]")

        response = execute_api_call(log, http_client, "get", file_url)
        if response and response.status_code == 200:
            csv_content = response.content.decode('utf-8')
            df = pd.read_csv(io.StringIO(csv_content))

            required_columns = ["FriendlyName", "Name", "ID", "Type", "Source"]
            if not all(col in df.columns for col in required_columns):
                log.error(f"Matrix file missing required columns {required_columns}")
                return {}

            matching_rows = df[df[lookup_column].str.strip().str.lower() == lookup_value.strip().lower()]
            if not matching_rows.empty:
                matched_data = {}
                for col in return_columns:
                    matched_data[col.lower()] = matching_rows.iloc[0][col]
                log.info(f"Matched row data: {matched_data}")
                return matched_data
            else:
                log.error(f"No match for [{lookup_value}] under column [{lookup_column}]")
                return {}
        else:
            log.error(f"Failed to download matrix file [{file_url}] Status code: {response.status_code if response else 'N/A'}")
            return {}
    except Exception as e:
        log.exception(e, "Exception while parsing group matrix file")
        return {}

def main():
    try:
        try:
            group_name = input.get_value("GroupName_1744233674989")
            ticket_number = input.get_value("TicketNumber_1744234484500")
        except Exception as e:
            log.exception(e, "Failed to fetch input values")
            log.result_message(ResultLevel.FAILED, "Failed to fetch input values")
            return
        
        group_name = group_name.strip() if group_name else ""
        ticket_number = ticket_number.strip() if ticket_number else ""

        log.info(f"Received inputs GroupName = [{group_name}], Ticket = [{ticket_number}]")

        if not group_name:
            log.error("Group name is empty or invalid")
            log.result_message(ResultLevel.FAILED, "Group name is empty or invalid")
            return

        if not ticket_number:
            log.error("Ticket number is empty or invalid")
            log.result_message(ResultLevel.FAILED, "Ticket number is empty or invalid")
            return

        company_identifier = get_company_identifier_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
        if not company_identifier:
            log.result_message(ResultLevel.FAILED, f"Failed to retrieve company identifier from ticket [{ticket_number}]")
            return

        matched_data = get_matrix_object_id(log, http_client, company_identifier, "FriendlyName", group_name, ["FriendlyName", "Name", "ID", "Type", "Source"])

        if not matched_data:
            log.result_message(ResultLevel.FAILED, f"No match found for [{group_name}]")
            return

        data_to_log = {
            "friendly_name": matched_data.get("friendlyname", ""),
            "name": matched_data.get("name", ""),
            "id": matched_data.get("id", ""),
            "type": matched_data.get("type", ""),
            "source": matched_data.get("source", "")
        }

        log.result_data(data_to_log)
        log.result_message(ResultLevel.SUCCESS, f"Successfully matched group [{data_to_log.get('friendly_name')}]")

    except Exception:
        log.exception("An error occurred while processing")
        log.result_message(ResultLevel.FAILED, "Process failed")

if __name__ == "__main__":
    main()