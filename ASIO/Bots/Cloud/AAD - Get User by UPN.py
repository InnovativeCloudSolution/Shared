import random
import time
import urllib.parse
import requests
from cw_rpa import Logger, Input, HttpClient, ResultLevel

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

msgraph_base_url_base = "https://graph.microsoft.com"
msgraph_base_url_path = "/v1.0"

data_to_log = {}
bot_name = "AAD - Get User by UPN"
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

def main():
    try:
        try:
            user_upn = input.get_value("UserUPN_1772806337064")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        user_upn = user_upn.strip() if user_upn else ""

        log.info(f"User UPN = [{user_upn}]")

        if not user_upn:
            record_result(log, ResultLevel.WARNING, "UPN is required but missing")
            return

        encoded_upn = urllib.parse.quote(user_upn, safe="@.")
        endpoint = f"{msgraph_base_url_base}{msgraph_base_url_path}/users/{encoded_upn}?$select=id,displayName,userPrincipalName,assignedLicenses"
        response = execute_api_call(log, http_client, "get", endpoint, integration_name="custom_wf_oauth2_client_creds")

        if not response:
            record_result(log, ResultLevel.WARNING, f"User not found for UPN [{user_upn}]")
            return

        user_data = response.json()
        if "error" in user_data:
            record_result(log, ResultLevel.WARNING, f"Graph API error: {user_data['error'].get('message', 'Unknown error')}")
            return

        log.info(f"User ID: [{user_data.get('id', '')}]")
        log.info(f"Display Name: [{user_data.get('displayName', '')}]")
        log.info(f"UPN: [{user_data.get('userPrincipalName', '')}]")
        log.info(f"Licensed: [{bool(user_data.get('assignedLicenses', []))}]")

        data_to_log["user_id"] = user_data.get("id", "")
        data_to_log["display_name"] = user_data.get("displayName", "")
        data_to_log["upn"] = user_data.get("userPrincipalName", "")
        data_to_log["licensed"] = bool(user_data.get("assignedLicenses", []))

        record_result(log, ResultLevel.SUCCESS, f"User [{user_data.get('displayName', '')}] retrieved successfully")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
