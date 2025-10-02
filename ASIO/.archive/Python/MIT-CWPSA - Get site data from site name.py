import sys
import random
import os
import time
import urllib.parse
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
bot_name = "MIT-CWPSA - Get site data from site name"
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

def get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number):
    log.info(f"Retrieving company details for ticket [{ticket_number}]")
    ticket_endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}"
    ticket_response = execute_api_call(log, http_client, "get", ticket_endpoint, integration_name="cw_psa")

    if ticket_response and ticket_response.status_code == 200:
        ticket_data = ticket_response.json()
        company = ticket_data.get("company", {})

        company_id = company["id"]
        company_identifier = company["identifier"]
        company_name = company["name"]

        log.info(f"Company ID: [{company_id}], Identifier: [{company_identifier}], Name: [{company_name}]")

        company_endpoint = f"{cwpsa_base_url}/company/companies/{company_id}"
        company_response = execute_api_call(log, http_client, "get", company_endpoint, integration_name="cw_psa")

        company_type = ""
        if company_response and company_response.status_code == 200:
            company_data = company_response.json()
            company_type = company_data.get("type", {}).get("name", "")
            log.info(f"Company type for ID [{company_id}]: [{company_type}]")
        else:
            log.warning(f"Unable to retrieve company type for ID [{company_id}]")

        return company_identifier, company_name, company_id, company_type

    elif ticket_response:
        log.error(
            f"Failed to retrieve ticket [{ticket_number}] "
            f"Status: {ticket_response.status_code}, Body: {ticket_response.text}"
        )
    return "", "", 0, ""

def get_company_site_data(log, http_client, cwpsa_base_url, company_id, site_name):
    log.info(f"Searching for site [{site_name}] under company ID [{company_id}]")

    encoded_name = urllib.parse.quote(site_name)
    endpoint = f"{cwpsa_base_url}/company/companies/{company_id}/sites?conditions=name='{encoded_name}'"
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")

    if response:
        sites = response.json()
        if not sites:
            log.warning(f"No site found with name [{site_name}] for company ID [{company_id}]")
            return {}

        site = sites[0]
        site_id = site.get("id")
        log.info(f"Found site ID [{site_id}] for name [{site_name}]")


        detail_endpoint = f"{cwpsa_base_url}/company/companies/{company_id}/sites/{site_id}"
        detail_response = execute_api_call(log, http_client, "get", detail_endpoint, integration_name="cw_psa")

        if detail_response and detail_response.status_code == 200:
            site_data = detail_response.json()
            custom_fields = site_data.get("customFields", [])
            try:
                friendly_name_data = [f for f in custom_fields if f.get("caption") == "Friendly Site Name"]
                friendly_name = friendly_name_data[0]["value"] if friendly_name_data else ""
            except Exception:
                log.warning(f"Failed to decode site data custom fields JSON, or no friendly site name found")
                friendly_name = ""
            address = {
                "site_id": site_id,
                "site_friendly": friendly_name,
                "site_name": friendly_name if friendly_name else site.get("name", ""),
                "site_streetaddress": f"{site_data.get('addressLine1', '')}, {site_data.get('addressLine2', '')}",
                "site_city": site_data.get("city", ""),
                "site_state": site_data.get("stateReference", {}).get("identifier", ""),
                "site_postal_code": site_data.get("zip", ""),
                "site_country": site_data.get("country", {}).get("name", "")
            }

            log.info(f"Resolved site data: {address}")
            return address
        else:
            log.error(f"Failed to fetch full site data for site ID [{site_id}]")
            return {}
    else:
        log.error(f"Failed to query site list for company ID [{company_id}] using name [{site_name}]")
        return {}

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1757476657657")
            site_name = input.get_value("SiteName_1757476659264")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        site_name = site_name.strip() if site_name else ""

        log.info(f"Received input ticket = [{ticket_number}]")
        log.info(f"Received input site = [{site_name}]")

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number is required")
            return

        company_identifier, _, company_id, _ = get_company_data_from_ticket(
            log, http_client, cwpsa_base_url, ticket_number
        )
        if not company_identifier:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve company identifier from ticket [{ticket_number}]")
            return

        site_data = get_company_site_data(log, http_client, cwpsa_base_url, company_id, site_name)
        if site_data:
            data_to_log.update(site_data)
            record_result(log, ResultLevel.SUCCESS, "Successfully retrieved site details")
        else:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve site data for name [{site_name}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
