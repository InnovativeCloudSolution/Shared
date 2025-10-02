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

def get_group(log, http_client, msgraph_base_url, group_identifier, token):
    log.info(f"Resolving group email for [{group_identifier}]")
    headers = {"Authorization": f"Bearer {token}"}

    filters = [
        f"startswith(displayName,'{group_identifier}')",
        f"startswith(mail,'{group_identifier}')"
    ]
    filter_query = " or ".join(filters)
    endpoint = f"{msgraph_base_url}/groups?$filter={urllib.parse.quote(filter_query)}"

    response = execute_api_call(log, http_client, "get", endpoint, headers=headers)

    if response and response.status_code == 200:
        groups = response.json().get("value", [])
        if len(groups) > 1:
            log.error(f"Multiple groups found for [{group_identifier}]")
            return groups
        if groups:
            group = groups[0]
            group_email = group.get("mail", "")
            group_id = group.get("id", "")
            log.info(f"Group found for [{group_identifier}] - Email: {group_email}, ID: {group_id}")
            return group_email, group_id

    log.error(f"Failed to resolve group for [{group_identifier}]")
    return "", ""

def get_ca_policies(log, http_client, msgraph_base_url, group_identifier, token):
    try:
        log.info(f"Searching Conditional Access policies excluding group [{group_identifier}]")
        headers = {"Authorization": f"Bearer {token}"}
        endpoint = f"{msgraph_base_url}/identity/conditionalAccess/policies?$select=id,displayName,conditions"

        response = execute_api_call(log, http_client, "get", endpoint, headers=headers)

        if not response or response.status_code != 200:
            log.error(f"Failed to retrieve Conditional Access policies Status: {response.status_code if response else 'N/A'}")
            return []

        policies = response.json().get("value", [])
        if not policies:
            log.error("No Conditional Access policies found")
            return []

        matching_policies = []
        for policy in policies:
            policy_id = policy.get("id", "")
            policy_name = policy.get("displayName", "Unnamed Policy")
            excluded_groups = policy.get("conditions", {}).get("users", {}).get("excludeGroups", [])

            if excluded_groups:
                # Match by Object ID directly
                if group_identifier in excluded_groups:
                    matching_policies.append({
                        "id": policy_id,
                        "name": policy_name
                    })
                else:
                    # Match by Name (if provided and needed)
                    if policy_name.lower() == group_identifier.lower():
                        matching_policies.append({
                            "id": policy_id,
                            "name": policy_name
                        })

        log.info(f"Found {len(matching_policies)} policy(ies) matching [{group_identifier}]")
        return matching_policies

    except Exception as e:
        log.exception(e, "Exception occurred while finding Conditional Access policies")
        return []

def main():
    try:
        try:
            group_identifier = input.get_value("Group_1744579845409")
            ticket_number = input.get_value("TicketNumber_1744579898396")
            auth_code = input.get_value("AuthCode_1744579852815")
            provided_token = input.get_value("AccessToken_1744579854813")
        except Exception as e:
            log.exception(e, "Failed to fetch input values")
            log.result_message(ResultLevel.FAILED, "Failed to fetch input values")
            return
        
        group_identifier = group_identifier.strip() if group_identifier else ""
        ticket_number = ticket_number.strip() if ticket_number else ""
        auth_code = auth_code.strip() if auth_code else ""
        provided_token = provided_token.strip() if provided_token else ""

        log.info(f"Received input Group Identifier = [{group_identifier}], Ticket = [{ticket_number}]")

        if not group_identifier:
            log.error("Group identifier is empty or invalid")
            log.result_message(ResultLevel.FAILED, "Group identifier input is required")
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

        if not re.fullmatch(r"[0-9a-fA-F-]{36}", group_identifier):
            _, resolved_group_id = get_group(log, http_client, msgraph_base_url, group_identifier, access_token)
            if not resolved_group_id:
                log.result_message(ResultLevel.FAILED, f"Unable to resolve group ID for [{group_identifier}]")
                return
            group_identifier = resolved_group_id

        ca_policies = get_ca_policies(log, http_client, msgraph_base_url, group_identifier, access_token)

        if not ca_policies:
            log.result_message(ResultLevel.INFO, f"No Conditional Access policies exclude or match group [{group_identifier}]")
            return

        if len(ca_policies) == 1:
            log.result_data({"excluded_policies": ca_policies[0]})
        else:
            log.result_data({"excluded_policies": ca_policies})

        details = "\n".join(f"- {policy['name']}" for policy in ca_policies)
        log.result_message(ResultLevel.SUCCESS, f"Group [{group_identifier}] is excluded from the following Conditional Access policies:\n{details}")

    except Exception:
        log.exception("An error occurred while processing")
        log.result_message(ResultLevel.FAILED, "Process failed")

if __name__ == "__main__":
    main()