import re
import sys
import random
import os
import time
import urllib.parse
import requests
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

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

def get_secret_value(log, http_client, vault_name, secret_name):
    log.info(f"Fetching secret [{secret_name}] from Key Vault [{vault_name}]")

    secret_url = f"https://{vault_name}.vault.azure.net/secrets/{secret_name}?api-version=7.3"
    response = execute_api_call(log, http_client, "get", secret_url, integration_name="custom_wf_oauth2_client_creds")

    if response and response.status_code == 200:
        secret_value = response.json().get("value", "")
        if secret_value:
            log.info(f"Successfully retrieved secret [{secret_name}]")
            return secret_value

    log.error(f"Failed to retrieve secret [{secret_name}] Status code: {response.status_code if response else 'N/A'}")
    return ""

def get_tenant_id_from_domain(log, http_client, azure_domain):
    try:
        config_url = f"https://login.windows.net/{azure_domain}/.well-known/openid-configuration"
        log.info(f"Fetching OpenID configuration from [{config_url}]")

        response = execute_api_call(log, http_client, "get", config_url)

        if response and response.status_code == 200:
            token_endpoint = response.json().get("token_endpoint", "")
            tenant_id = token_endpoint.split("/")[3] if token_endpoint else ""
            if tenant_id:
                log.info(f"Successfully extracted tenant ID [{tenant_id}]")
                return tenant_id

        log.error(f"Failed to extract tenant ID from domain [{azure_domain}]")
        return ""
    except Exception as e:
        log.exception(e, "Exception while extracting tenant ID from domain")
        return ""

def get_access_token(log, http_client, tenant_id, client_id, client_secret, scope="https://graph.microsoft.com/.default"):
    log.info(f"Requesting access token for scope [{scope}]")

    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    payload = urllib.parse.urlencode({
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": scope
    })
    headers = {"Content-Type": "application/x-www-form-urlencoded"}

    response = execute_api_call(log, http_client, "post", token_url, data=payload, retries=3, headers=headers)

    if response and response.status_code == 200:
        token_data = response.json()
        access_token = str(token_data.get("access_token", "")).strip()
        log.info(f"Access token length: {len(access_token)}")
        log.info(f"Access token preview: {access_token[:30]}...")

        if not isinstance(access_token, str) or "." not in access_token:
            log.error("Access token is invalid or malformed")
            return ""

        log.info("Successfully retrieved access token")
        return access_token

    log.error(f"Failed to retrieve access token Status code: {response.status_code if response else 'N/A'}")
    return ""

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

def validate_mit_authentication(log, http_client, vault_name, auth_code):
    if not auth_code:
        log.result_message(ResultLevel.FAILED, "Authentication code input is required for MIT")
        return False

    expected_code = get_secret_value(log, http_client, vault_name, "MIT-AuthenticationCode")
    if not expected_code:
        log.result_message(ResultLevel.FAILED, "Failed to retrieve expected authentication code for MIT")
        return False

    if auth_code.strip() != expected_code.strip():
        log.result_message(ResultLevel.FAILED, "Provided authentication code is incorrect")
        return False

    return True

def get_policy_id(log, http_client, msgraph_base_url, policy_name_or_id, token):
    if re.fullmatch(r"[0-9a-fA-F\-]{36}", policy_name_or_id):
        log.info(f"Input [{policy_name_or_id}] looks like a Policy ID, fetching full policy")

        headers = {"Authorization": f"Bearer {token}"}
        endpoint = f"{msgraph_base_url}/identity/conditionalAccess/policies/{policy_name_or_id}"
        response = execute_api_call(log, http_client, "get", endpoint, headers=headers)

        if response and response.status_code == 200:
            return policy_name_or_id, response.json()

        log.error(f"Failed to retrieve CA policy by ID [{policy_name_or_id}] Status: {response.status_code if response else 'N/A'}")
        return "", {}

    log.info(f"Searching for Conditional Access policy by name [{policy_name_or_id}]")

    headers = {"Authorization": f"Bearer {token}"}
    endpoint = f"{msgraph_base_url}/identity/conditionalAccess/policies?$filter=displayName eq '{policy_name_or_id}'&$select=id,displayName"

    response = execute_api_call(log, http_client, "get", endpoint, headers=headers)

    if response and response.status_code == 200:
        policies = response.json().get("value", [])
        if not policies:
            log.error(f"No CA policy found with name [{policy_name_or_id}]")
            return "", {}
        if len(policies) > 1:
            log.error(f"Multiple CA policies found with name [{policy_name_or_id}]")
            return "", {}
        policy_id = policies[0].get("id", "")
        full_data = policies[0]
        log.info(f"Found CA Policy ID [{policy_id}] for policy name [{policy_name_or_id}]")
        return policy_id, full_data

    log.error(f"Failed to retrieve CA policy for name [{policy_name_or_id}] Status: {response.status_code if response else 'N/A'}")
    return "", {}

def get_location_id(log, http_client, msgraph_base_url, location_name_or_id, token):
    if re.fullmatch(r"[0-9a-fA-F\-]{36}", location_name_or_id):
        log.info(f"Input [{location_name_or_id}] looks like a Location ID, fetching full object")

        headers = {"Authorization": f"Bearer {token}"}
        endpoint = f"{msgraph_base_url}/identity/conditionalAccess/namedLocations/{location_name_or_id}"
        response = execute_api_call(log, http_client, "get", endpoint, headers=headers)

        if response and response.status_code == 200:
            return location_name_or_id, response.json()

        log.error(f"Failed to retrieve named location by ID [{location_name_or_id}] Status: {response.status_code if response else 'N/A'}")
        return "", {}

    log.info(f"Searching for named location by name [{location_name_or_id}]")

    headers = {"Authorization": f"Bearer {token}"}
    endpoint = f"{msgraph_base_url}/identity/conditionalAccess/namedLocations?$filter=displayName eq '{location_name_or_id}'&$select=id,displayName"

    response = execute_api_call(log, http_client, "get", endpoint, headers=headers)

    if response and response.status_code == 200:
        locations = response.json().get("value", [])
        if not locations:
            log.error(f"No named location found with name [{location_name_or_id}]")
            return "", {}
        if len(locations) > 1:
            log.error(f"Multiple named locations found with name [{location_name_or_id}]")
            return "", {}
        location_id = locations[0].get("id", "")
        full_data = locations[0]
        log.info(f"Found Location ID [{location_id}] for location name [{location_name_or_id}]")
        return location_id, full_data

    log.error(f"Failed to retrieve named location for name [{location_name_or_id}] Status: {response.status_code if response else 'N/A'}")
    return "", {}

def build_patch_data(log, update_type, update_values):
    try:
        log.info(f"Building patch data for update type [{update_type}]")

        if update_type == "excludeLocations":
            patch_data = {
                "conditions": {
                    "locations": {
                        "excludeLocations": update_values
                    }
                }
            }
            return patch_data

        elif update_type == "includeUsers":
            patch_data = {
                "conditions": {
                    "users": {
                        "includeUsers": update_values
                    }
                }
            }
            return patch_data

        elif update_type == "grantControls":
            patch_data = {
                "grantControls": update_values
            }
            return patch_data

        elif update_type == "sessionControls":
            patch_data = {
                "sessionControls": update_values
            }
            return patch_data

        else:
            log.error(f"Unsupported update type [{update_type}]")
            return {}

    except Exception as e:
        log.exception(e, "Exception occurred while building patch data")
        return {}

def update_ca_policy(log, http_client, msgraph_base_url, policy_id, patch_data, token):
    try:
        log.info(f"Updating CA Policy [{policy_id}] with dynamic patch")

        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        patch_endpoint = f"{msgraph_base_url}/identity/conditionalAccess/policies/{policy_id}"

        patch_response = execute_api_call(log, http_client, "patch", patch_endpoint, data=patch_data, headers=headers)

        if patch_response and patch_response.status_code in [200, 204]:
            log.info(f"Successfully updated CA policy [{policy_id}]")
            return True
        else:
            log.error(f"Failed to update CA policy [{policy_id}] Status: {patch_response.status_code if patch_response else 'N/A'}")
            return False

    except Exception as e:
        log.exception(e, f"Exception occurred while updating CA policy [{policy_id}]")
        return False

def main():
    try:
        try:
            policy_identifier = input.get_value("Policy_1744581140429")
            location_identifiers_input = input.get_value("Location_1744581143193") or ""
            ticket_number = input.get_value("TicketNumber_1744581146005")
            auth_code = input.get_value("AuthCode_1744581147644")
            provided_token = input.get_value("AccessToken_1744581149059")
            action = input.get_value("Action_1744581151590")
        except Exception as e:
            log.exception(e, "Failed to fetch input values")
            log.result_message(ResultLevel.FAILED, "Failed to fetch input values")
            return

        policy_identifier = policy_identifier.strip() if policy_identifier else ""
        location_identifiers = [l.strip() for l in location_identifiers_input.split(",") if l.strip()]
        ticket_number = ticket_number.strip() if ticket_number else ""
        auth_code = auth_code.strip() if auth_code else ""
        provided_token = provided_token.strip() if provided_token else ""
        action = action.strip().lower() if action else ""

        if not policy_identifier:
            log.result_message(ResultLevel.FAILED, "Policy identifier is required")
            return

        if not location_identifiers:
            log.result_message(ResultLevel.FAILED, "At least one location identifier is required")
            return

        if action not in ["add exclude location", "remove exclude location"]:
            log.result_message(ResultLevel.FAILED, f"Unsupported or missing action [{action}]")
            return

        if provided_token:
            access_token = provided_token
            log.info("Using provided access token")
            if not isinstance(access_token, str) or "." not in access_token:
                log.result_message(ResultLevel.FAILED, "Provided access token is malformed (missing dots)")
                return

        elif ticket_number:
            company_identifier = get_company_identifier_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
            if not company_identifier:
                log.result_message(ResultLevel.FAILED, f"Failed to retrieve company identifier from ticket [{ticket_number}]")
                return

            if company_identifier == "MIT":
                if not validate_mit_authentication(log, http_client, vault_name, auth_code):
                    return

            client_id = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientID")
            client_secret = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientSecret")
            azure_domain = get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")

            if not all([client_id, client_secret, azure_domain]):
                log.result_message(ResultLevel.FAILED, "Failed to retrieve required secrets")
                return

            tenant_id = get_tenant_id_from_domain(log, http_client, azure_domain)
            if not tenant_id:
                log.result_message(ResultLevel.FAILED, "Failed to resolve tenant ID")
                return

            access_token = get_access_token(log, http_client, tenant_id, client_id, client_secret)
            if not isinstance(access_token, str) or "." not in access_token:
                log.result_message(ResultLevel.FAILED, "Access token is malformed (missing dots)")
                return

        else:
            log.result_message(ResultLevel.FAILED, "Either Access Token or Ticket Number is required")
            return

        policy_id, policy_data = get_policy_id(log, http_client, msgraph_base_url, policy_identifier, access_token)
        if not policy_id:
            log.result_message(ResultLevel.FAILED, f"Failed to resolve policy ID for [{policy_identifier}]")
            return

        location_ids = []
        for identifier in location_identifiers:
            location_id, _ = get_location_id(log, http_client, msgraph_base_url, identifier, access_token)
            if location_id:
                location_ids.append(location_id)

        if not location_ids:
            log.result_message(ResultLevel.FAILED, "No valid location IDs found after resolution")
            return

        conditions = policy_data.get("conditions", {}) or {}
        locations = conditions.get("locations", {}) or {}
        existing_exclude = locations.get("excludeLocations", []) or []

        if action == "add exclude location":
            locations["excludeLocations"] = list(set(existing_exclude + location_ids))
        elif action == "remove exclude location":
            locations["excludeLocations"] = [l for l in existing_exclude if l not in location_ids]
        else:
            log.result_message(ResultLevel.FAILED, f"Unsupported action type [{action}]")
            return

        conditions["locations"] = locations
        patch_data = {"conditions": conditions}

        update_success = update_ca_policy(log, http_client, msgraph_base_url, policy_id, patch_data, access_token)

        if update_success:
            log.result_message(ResultLevel.SUCCESS, f"Successfully updated CA policy [{policy_id}]")
        else:
            log.result_message(ResultLevel.FAILED, f"Failed to update CA policy [{policy_id}]")

    except Exception:
        log.exception("An error occurred while processing")
        log.result_message(ResultLevel.FAILED, "Process failed")

if __name__ == "__main__":
    main()