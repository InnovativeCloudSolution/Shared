import sys
import random
import re
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
msgraph_base_url_beta = "https://graph.microsoft.com/beta"
vault_name = "PLACEHOLDER-akv1"

data_to_log = {}
bot_name = "AZ - CA Policy Management"
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

def get_secret_value(log, http_client, vault_name, secret_name):
    log.info(f"Fetching secret [{secret_name}] from Key Vault [{vault_name}]")
    secret_url = (f"https://{vault_name}.vault.azure.net/secrets/{secret_name}?api-version=7.3")
    response = execute_api_call(
        log,
        http_client,
        "get",
        secret_url,
        integration_name="custom_wf_oauth2_client_creds",
    )
    if response:
        secret_value = response.json().get("value", "")
        if secret_value:
            log.info(f"Successfully retrieved secret [{secret_name}]")
            return secret_value
    log.error(f"Failed to retrieve secret [{secret_name}] Status code: {response.status_code if response else 'N/A'}")
    return ""

def get_tenant_id_from_domain(log, http_client, azure_domain):
    try:
        config_url = (f"https://login.windows.net/{azure_domain}/.well-known/openid-configuration")
        log.info(f"Fetching OpenID configuration from [{config_url}]")
        response = execute_api_call(log, http_client, "get", config_url)
        if response:
            token_endpoint = response.json().get("token_endpoint", "")
            tenant_id = token_endpoint.split("/")[3] if token_endpoint else ""
            if tenant_id:
                log.info(f"Successfully extracted tenant ID [{tenant_id}]")
                return tenant_id
        return ""
    except Exception as e:
        log.exception(e, "Exception while extracting tenant ID from domain")
        return ""

def get_access_token(log, http_client, tenant_id, client_id, client_secret, scope="https://graph.microsoft.com/.default", log_prefix="Token"):
    log.info(f"[{log_prefix}] Requesting access token for scope [{scope}]")
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    payload = urllib.parse.urlencode(
        {
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": scope,
        }
    )
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    response = execute_api_call(log, http_client, "post", token_url, data=payload, retries=3, headers=headers)
    if response:
        token_data = response.json()
        access_token = str(token_data.get("access_token", "")).strip()
        log.info(f"[{log_prefix}] Access token length: {len(access_token)}")
        log.info(f"[{log_prefix}] Access token preview: {access_token[:30]}...")
        if not isinstance(access_token, str) or "." not in access_token:
            log.error(f"[{log_prefix}] Access token is invalid or malformed")
            return ""
        log.info(f"[{log_prefix}] Successfully retrieved access token")
        return access_token
    return ""

def get_graph_token(log, http_client, vault_name, company_identifier):
    client_id = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientID")
    client_secret = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientSecret")
    azure_domain = get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")
    if not all([client_id, client_secret, azure_domain]):
        log.error("Failed to retrieve required secrets for MS Graph")
        return "", ""
    tenant_id = get_tenant_id_from_domain(log, http_client, azure_domain)
    if not tenant_id:
        log.error("Failed to resolve tenant ID for MS Graph")
        return "", ""
    token = get_access_token(log, http_client, tenant_id, client_id, client_secret, scope="https://graph.microsoft.com/.default", log_prefix="Graph")
    if not isinstance(token, str) or "." not in token:
        log.error("MS Graph access token is malformed (missing dots)")
        return "", ""
    return tenant_id, token

def get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number):
    log.info(f"Retrieving company details for ticket [{ticket_number}]")
    ticket_endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}"
    ticket_response = execute_api_call(log, http_client, "get", ticket_endpoint, integration_name="cw_psa")
    if ticket_response:
        ticket_data = ticket_response.json()
        company = ticket_data.get("company", {})
        company_id = company["id"]
        company_identifier = company["identifier"]
        company_name = company["name"]
        log.info(f"Company ID: [{company_id}], Identifier: [{company_identifier}], Name: [{company_name}]")
        company_endpoint = f"{cwpsa_base_url}/company/companies/{company_id}"
        company_response = execute_api_call(log, http_client, "get", company_endpoint, integration_name="cw_psa")
        company_types = []
        if company_response:
            company_data = company_response.json()
            types = company_data.get("types", [])
            company_types = [t.get("name", "") for t in types if "name" in t]
            log.info(f"Company types for ID [{company_id}]: {company_types}")
        else:
            log.warning(f"Unable to retrieve company types for ID [{company_id}]")
        return company_identifier, company_name, company_id, company_types
    return "", "", 0, []

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
        if response:
            return policy_name_or_id, response.json()
        
        return "", {}

    log.info(f"Searching for Conditional Access policy by name [{policy_name_or_id}]")
    headers = {"Authorization": f"Bearer {token}"}
    endpoint = f"{msgraph_base_url}/identity/conditionalAccess/policies?$filter=displayName eq '{policy_name_or_id}'&$select=id,displayName"
    response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
    if response:
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
    
    return "", {}

def get_location_id(log, http_client, msgraph_base_url, location_name_or_id, token):
    if re.fullmatch(r"[0-9a-fA-F\-]{36}", location_name_or_id):
        log.info(f"Input [{location_name_or_id}] looks like a Location ID, fetching full object")
        headers = {"Authorization": f"Bearer {token}"}
        endpoint = f"{msgraph_base_url}/identity/conditionalAccess/namedLocations/{location_name_or_id}"
        response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
        if response:
            return location_name_or_id, response.json()
        
        return "", {}

    log.info(f"Searching for named location by name [{location_name_or_id}]")
    headers = {"Authorization": f"Bearer {token}"}
    endpoint = f"{msgraph_base_url}/identity/conditionalAccess/namedLocations?$filter=displayName eq '{location_name_or_id}'&$select=id,displayName"
    response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
    if response:
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
    
    return "", {}

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
    if response:
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
        if not response:
            
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
                if group_identifier in excluded_groups:
                    matching_policies.append({
                        "id": policy_id,
                        "name": policy_name
                    })
                else:
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

def get_named_location(log, http_client, msgraph_base_url, location_identifier, token):
    try:
        log.info(f"Searching Conditional Access Named Location [{location_identifier}]")
        headers = {"Authorization": f"Bearer {token}"}
        endpoint = f"{msgraph_base_url}/identity/conditionalAccess/namedLocations"
        response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
        if not response:
            
            return []
        locations = response.json().get("value", [])
        if not locations:
            log.error("No Conditional Access Named Locations found")
            return []
        matching_locations = []
        for location in locations:
            location_id = location.get("id", "")
            location_name = location.get("displayName", "Unnamed Location")
            if location_identifier.lower() in [location_id.lower(), location_name.lower()]:
                matching_locations.append({
                    "id": location_id,
                    "name": location_name,
                    "type": location.get("@odata.type", ""),
                    "countryLookupMethod": location.get("countryLookupMethod", ""),
                    "countriesAndRegions": location.get("countriesAndRegions", []),
                    "ipRanges": location.get("ipRanges", [])
                })
        log.info(f"Found {len(matching_locations)} matching Conditional Access Named Location(s)")
        return matching_locations
    except Exception as e:
        log.exception(e, "Exception occurred while finding Conditional Access Named Locations")
        return []

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
        return False
    except Exception as e:
        log.exception(e, f"Exception occurred while updating CA policy [{policy_id}]")
        return False

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1756860025428")
            operation = input.get_value("Operation_1756860027714")
            policy_identifier = input.get_value("CAPolicy_1756860039923")
            location_identifiers_input = input.get_value("CANamedLocations_1756860038704")
            group_identifier = input.get_value("Groups_1756860042508")
            auth_code = input.get_value("AuthenticationCode_1756860041166")
            graph_token = input.get_value("GraphToken_1756860044094")
        except Exception as e:
            log.exception(e, "Failed to fetch input values")
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        operation = operation.strip() if operation else ""
        policy_identifier = policy_identifier.strip() if policy_identifier else ""
        location_identifiers_input = location_identifiers_input.strip() if location_identifiers_input else ""
        location_identifiers = [l.strip() for l in location_identifiers_input.split(",") if l.strip()] if location_identifiers_input else []
        group_identifier = group_identifier.strip() if group_identifier else ""
        auth_code = auth_code.strip() if auth_code else ""
        graph_token = graph_token.strip() if graph_token else ""

        log.info(f"Ticket Number = [{ticket_number}]")
        log.info(f"Requested operation = [{operation}]")

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number is required but missing")
            return
        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation value is missing or invalid")
            return

        if operation not in ["Add Exclude Location", "Remove Exclude Location", "Policy Lookup", "Location Lookup"]:
            record_result(log, ResultLevel.WARNING, f"Unsupported operation [{operation}]")
            return

        company_identifier, company_name, company_id, company_type = get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
        if not company_identifier:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve company identifier from ticket [{ticket_number}]")
            return

        if company_identifier == "MIT":
            if not validate_mit_authentication(log, http_client, vault_name, auth_code):
                return

        if graph_token:
            log.info("Using provided MS Graph token")
            graph_access_token = graph_token
            graph_tenant_id = ""
        else:
            graph_tenant_id, graph_access_token = get_graph_token(log, http_client, vault_name, company_identifier)
            if not graph_access_token:
                record_result(log, ResultLevel.WARNING, "Failed to obtain MS Graph access token")
                return

        azure_domain = get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")
        if not azure_domain:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve Azure domain for [{company_identifier}]")
            return

        if operation == "Add Exclude Location":
            if not policy_identifier:
                record_result(log, ResultLevel.WARNING, "Policy identifier is required for add exclude location operation")
                return
            if not location_identifiers:
                record_result(log, ResultLevel.WARNING, "At least one location identifier is required for add exclude location operation")
                return

            policy_id, policy_data = get_policy_id(log, http_client, msgraph_base_url, policy_identifier, graph_access_token)
            if not policy_id:
                record_result(log, ResultLevel.WARNING, f"Failed to resolve policy ID for [{policy_identifier}]")
                return

            location_ids = []
            for identifier in location_identifiers:
                location_id, _ = get_location_id(log, http_client, msgraph_base_url, identifier, graph_access_token)
                if location_id:
                    location_ids.append(location_id)

            if not location_ids:
                record_result(log, ResultLevel.WARNING, "No valid location IDs found after resolution")
                return

            conditions = policy_data.get("conditions", {}) or {}
            locations = conditions.get("locations", {}) or {}
            existing_exclude = locations.get("excludeLocations", []) or []

            updated_exclude_locations = list(set(existing_exclude + location_ids))
            patch_data = build_patch_data(log, "excludeLocations", updated_exclude_locations)

            update_success = update_ca_policy(log, http_client, msgraph_base_url, policy_id, patch_data, graph_access_token)

            if update_success:
                record_result(log, ResultLevel.SUCCESS, f"Successfully added exclude locations to CA policy [{policy_id}]")
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to add exclude locations to CA policy [{policy_id}]")

        elif operation == "Remove Exclude Location":
            if not policy_identifier:
                record_result(log, ResultLevel.WARNING, "Policy identifier is required for remove exclude location operation")
                return
            if not location_identifiers:
                record_result(log, ResultLevel.WARNING, "At least one location identifier is required for remove exclude location operation")
                return

            policy_id, policy_data = get_policy_id(log, http_client, msgraph_base_url, policy_identifier, graph_access_token)
            if not policy_id:
                record_result(log, ResultLevel.WARNING, f"Failed to resolve policy ID for [{policy_identifier}]")
                return

            location_ids = []
            for identifier in location_identifiers:
                location_id, _ = get_location_id(log, http_client, msgraph_base_url, identifier, graph_access_token)
                if location_id:
                    location_ids.append(location_id)

            if not location_ids:
                record_result(log, ResultLevel.WARNING, "No valid location IDs found after resolution")
                return

            conditions = policy_data.get("conditions", {}) or {}
            locations = conditions.get("locations", {}) or {}
            existing_exclude = locations.get("excludeLocations", []) or []

            updated_exclude_locations = [l for l in existing_exclude if l not in location_ids]
            patch_data = build_patch_data(log, "excludeLocations", updated_exclude_locations)

            update_success = update_ca_policy(log, http_client, msgraph_base_url, policy_id, patch_data, graph_access_token)

            if update_success:
                record_result(log, ResultLevel.SUCCESS, f"Successfully removed exclude locations from CA policy [{policy_id}]")
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to remove exclude locations from CA policy [{policy_id}]")

        elif operation == "Policy Lookup":
            if not group_identifier:
                record_result(log, ResultLevel.WARNING, "Group identifier is required for policy lookup operation")
                return

            if not re.fullmatch(r"[0-9a-fA-F-]{36}", group_identifier):
                _, resolved_group_id = get_group(log, http_client, msgraph_base_url, group_identifier, graph_access_token)
                if not resolved_group_id:
                    record_result(log, ResultLevel.WARNING, f"Unable to resolve group ID for [{group_identifier}]")
                    return
                group_identifier = resolved_group_id

            ca_policies = get_ca_policies(log, http_client, msgraph_base_url, group_identifier, graph_access_token)

            if not ca_policies:
                record_result(log, ResultLevel.SUCCESS, f"No Conditional Access policies exclude or match group [{group_identifier}]")
                return

            excluded_policies_data = []
            for policy in ca_policies:
                excluded_policies_data.append({
                    "id": policy["id"],
                    "name": policy["name"]
                })
            data_to_log["excluded_policies"] = excluded_policies_data

            details = "\n".join(f"- {policy['name']}" for policy in ca_policies)
            record_result(log, ResultLevel.SUCCESS, f"Group [{group_identifier}] is excluded from the following Conditional Access policies:\n{details}")

        elif operation == "Location Lookup":
            if not location_identifiers_input:
                record_result(log, ResultLevel.WARNING, "Location identifier is required for location lookup operation")
                return

            location_identifier = location_identifiers_input.strip()
            named_locations = get_named_location(log, http_client, msgraph_base_url, location_identifier, graph_access_token)

            if not named_locations:
                record_result(log, ResultLevel.SUCCESS, f"No Conditional Access Named Location found for [{location_identifier}]")
                return

            named_locations_data = []
            for location in named_locations:
                countries = ",".join(location.get("countriesAndRegions", []))
                ip_ranges = ",".join([ip_range.get("cidrAddress", "") for ip_range in location.get("ipRanges", []) if ip_range.get("cidrAddress")])
                named_locations_data.append({
                    "id": location["id"],
                    "name": location["name"],
                    "type": location.get("type", ""),
                    "countryLookupMethod": location.get("countryLookupMethod", ""),
                    "countriesAndRegions": countries,
                    "ipRanges": ip_ranges
                })
            data_to_log["named_locations"] = named_locations_data

            details = "\n".join(f"- {loc['name']} ({loc['id']})" for loc in named_locations)
            record_result(log, ResultLevel.SUCCESS, f"Found Conditional Access Named Location(s):\n{details}")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
