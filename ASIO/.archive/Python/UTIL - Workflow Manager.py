import sys
import traceback
import json
import random
import re
import subprocess
import os
import io
import base64
import hmac
import hashlib
import time
import urllib.parse
import string
import requests
import pandas as pd
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from cw_rpa import Logger, Input, HttpClient, ResultLevel

# === SHARED ===
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

# === SHARED ===
log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

# === CLOUD SCRIPT ===
cwpsa_base_url = "https://au.myconnectwise.net/v4_6_release/apis/3.0"
msgraph_base_url = "https://graph.microsoft.com/v1.0"
msgraph_base_url_beta = "https://graph.microsoft.com/beta"
vault_name = "mit-azu1-prod1-akv1"

# === SHARED ===
data_to_log = {}
log.info("Static variables set")


# === SHARED ===
def record_result(log, level, message):
    log.result_message(level, message)
    if level == ResultLevel.WARNING:
        data_to_log["Result"] = "Fail"
    elif level == ResultLevel.SUCCESS:
        if "Result" not in data_to_log:
            data_to_log["Result"] = "Success"


# === CLOUD SCRIPT ===
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

# === CLOUD SCRIPT ===
def generate_cosmosdb_auth_header(verb, resource_type, resource_link, date_utc, master_key):
    key = base64.b64decode(master_key)
    text = f"{verb.lower()}\n{resource_type.lower()}\n{resource_link}\n{date_utc.lower()}\n\n"
    signature = base64.b64encode(hmac.new(key, text.encode("utf-8"), hashlib.sha256).digest()).decode()
    return f"type=master&ver=1.0&sig={urllib.parse.quote(signature)}"

# === CLOUD SCRIPT ===
def execute_cosmosdb_call(log, http_client, method, endpoint, data=None, headers=None):
    try:
        if not all([endpoint, headers]):
            log.error("Missing required Cosmos DB call parameters")
            return None

        response = execute_api_call(log, http_client, method, endpoint, data=data, headers=headers)

        if response and response.status_code in [200, 201]:
            return response

        log.error(
            f"Cosmos DB call failed Status: {response.status_code if response else 'N/A'}, Body: {response.text if response else 'N/A'}"
        )
        return None

    except Exception as e:
        log.exception(e, "Exception occurred during Cosmos DB call")
        return None

# === CLOUD SCRIPT ===
def get_cosmosdb_document(log, http_client, cosmos_account, db_id, container_id, document_id, partition_key, vault_name, secret_name,):
    try:
        date_utc = datetime.utcnow().strftime("%a, %d %b %Y %H:%M:%S GMT")
        resource_type = "docs"
        resource_link = f"dbs/{db_id}/colls/{container_id}/docs/{document_id}"
        verb = "GET"

        cosmos_key = get_secret_value(log, http_client, vault_name, secret_name)
        if not cosmos_key:
            log.error("Missing Cosmos DB key")
            return None

        auth_token = generate_cosmosdb_auth_header(verb, resource_type, resource_link, date_utc, cosmos_key)

        headers = {
            "Authorization": auth_token,
            "x-ms-date": date_utc,
            "x-ms-version": "2018-12-31",
            "x-ms-documentdb-partitionkey": f'["{partition_key}"]',
            "Accept": "application/json",
        }

        endpoint = f"https://{cosmos_account}.documents.azure.com/{resource_link}"

        response = execute_cosmosdb_call(log, http_client, "get", endpoint, headers=headers)
        return response.json() if response else None

    except Exception as e:
        log.exception(e, "Exception occurred in get_cosmosdb_document")
        return None

# === SHARED ===
def main():
    try:
        try:
            # === SHARED ===
            user_identifier = input.get_value("User_xxxxxxxxxxxxx")
            operation = input.get_value("Operation_xxxxxxxxxxxxx")

            # === CLOUD ONLY ===
            ticket_number = input.get_value("TicketNumber_xxxxxxxxxxxxx")
            auth_code = input.get_value("AuthCode_xxxxxxxxxxxxx")
            provided_token = input.get_value("AccessToken_xxxxxxxxxxxxx")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        # === SHARED ===
        user_identifier = user_identifier.strip() if user_identifier else ""

        # === SHARED ===
        log.info(f"Received input user = [{user_identifier}]")
        log.info(f"Requested operation = [{operation}]")

        # === SHARED ===
        if not user_identifier:
            record_result(log, ResultLevel.WARNING, "User identifier is empty or invalid")
            return
        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation value is missing or invalid")
            return

        # === CLOUD ONLY ===
        access_token = ""
        exo_access_token = ""
        company_identifier = ""
        user_id = user_email = user_sam = ""

        if provided_token:
            access_token = provided_token
            log.info("Using provided access token")
            if not isinstance(access_token, str) or "." not in access_token:
                record_result(
                    log,
                    ResultLevel.WARNING,
                    "Provided access token is malformed (missing dots)",
                )
                return
        elif ticket_number:
            log.info(f"Retrieving company data for ticket [{ticket_number}]")
            company_identifier, company_name, company_id, company_type = get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
            if not company_identifier:
                record_result(
                    log,
                    ResultLevel.WARNING,
                    f"Failed to retrieve company identifier from ticket [{ticket_number}]",
                )
                return

            # === MIT Validator ===
            if company_identifier == "MIT":
                if not validate_mit_authentication(log, http_client, vault_name, auth_code):
                    return

            # === AAD Token ===
            client_id = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientID")
            client_secret = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientSecret")
            azure_domain = get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")

            if not all([client_id, client_secret, azure_domain]):
                record_result(log, ResultLevel.WARNING, "Failed to retrieve required secrets")
                return

            tenant_id = get_tenant_id_from_domain(log, http_client, azure_domain)
            if not tenant_id:
                record_result(log, ResultLevel.WARNING, "Failed to resolve tenant ID")
                return

            access_token = get_access_token(log, http_client, tenant_id, client_id, client_secret)
            if not isinstance(access_token, str) or "." not in access_token:
                record_result(log, ResultLevel.WARNING, "Access token is malformed (missing dots)")
                return

            # === EXO Token ===
            exo_client_id = get_secret_value(
                log,
                http_client,
                vault_name,
                f"{company_identifier}-ExchangeApp-ClientID",
            )
            exo_client_secret = get_secret_value(
                log,
                http_client,
                vault_name,
                f"{company_identifier}-ExchangeApp-ClientSecret",
            )

            if not all([exo_client_id, exo_client_secret, azure_domain]):
                record_result(
                    log,
                    ResultLevel.WARNING,
                    "Failed to retrieve required internal secrets",
                )
                return

            exo_access_token = get_access_token(
                log,
                http_client,
                tenant_id,
                exo_client_id,
                exo_client_secret,
                scope="https://outlook.office365.com/.default",
            )
            if not isinstance(exo_access_token, str) or "." not in exo_access_token:
                record_result(
                    log,
                    ResultLevel.WARNING,
                    "Exchange Online access token is malformed (missing dots)",
                )
                return

        # === CLOUD: AAD user resolution ===
        aad_user_result = get_aad_user_data(log, http_client, msgraph_base_url, user_identifier, access_token)

        if isinstance(aad_user_result, list):
            details = "\n".join(
                [f"- {u.get('displayName')} | {u.get('userPrincipalName')} | {u.get('id')}" for u in aad_user_result]
            )
            record_result(
                log,
                ResultLevel.WARNING,
                f"Multiple users found for [{user_identifier}]\n{details}",
            )
            return

        user_id, user_email, user_sam = aad_user_result

        # === DEVICE: AD user resolution ===
        ad_user_result = get_ad_user_data(log, user_identifier)
        user_email, user_sam = ad_user_result

        # === CLOUD ===
        if not user_id:
            record_result(
                log,
                ResultLevel.WARNING,
                f"Failed to resolve user ID for [{user_identifier}]",
            )
            return

        # === SHARED ===
        if not user_email:
            record_result(
                log,
                ResultLevel.WARNING,
                f"Unable to resolve user principal name for [{user_identifier}]",
            )
            return
        if not user_sam:
            record_result(
                log,
                ResultLevel.WARNING,
                f"No SAM account name found for [{user_identifier}]",
            )
            return

        # === SHARED ===
        log.info(f"User [{user_identifier}] matched with UPN [{user_email}] and SAM [{user_sam}]")

        # === SHARED ===
        data_to_log["user_id"] = user_id
        data_to_log["user_upn"] = user_email
        data_to_log["user_sam"] = user_sam
        record_result(log, ResultLevel.SUCCESS, "User successfully resolved")

        # === SHARED ===
        if operation == "Option A":
            log.info("Executing operation: Option A")
            # PLACE LOGIC HERE

        elif operation == "Option B":
            log.info("Executing operation: Option B")
            # PLACE LOGIC HERE

        elif operation == "Option C":
            log.info("Executing operation: Option C")
            # PLACE LOGIC HERE

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
