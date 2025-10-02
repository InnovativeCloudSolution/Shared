import sys
import random
import re
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
vault_name = "mit-azu1-prod1-akv1"

data_to_log = {}
bot_name = "MIT-MSTeams - Replace teams owner"
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
    secret_url = f"https://{vault_name}.vault.azure.net/secrets/{secret_name}?api-version=7.3"
    response = execute_api_call(log, http_client, "get", secret_url, integration_name="custom_wf_oauth2_client_creds")
    if response:
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
        if response:
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

def get_access_token(log, http_client, tenant_id, client_id, client_secret, scope="https://graph.microsoft.com/.default", log_prefix="Token"):
    log.info(f"[{log_prefix}] Requesting access token for scope [{scope}]")
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    payload = urllib.parse.urlencode({
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": scope
    })
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
    log.error(f"[{log_prefix}] Failed to retrieve access token Status code: {response.status_code if response else 'N/A'}")
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
    if ticket_response and ticket_response.status_code == 200:
        ticket_data = ticket_response.json()
        company = ticket_data.get("company", {})
        company_id = company["id"]
        company_identifier = company["identifier"]
        company_name = company["name"]
        log.info(f"Company ID: [{company_id}], Identifier: [{company_identifier}], Name: [{company_name}]")
        company_endpoint = f"{cwpsa_base_url}/company/companies/{company_id}"
        company_response = execute_api_call(log, http_client, "get", company_endpoint, integration_name="cw_psa")
        company_types = []
        if company_response and company_response.status_code == 200:
            company_data = company_response.json()
            types = company_data.get("types", [])
            company_types = [t.get("name", "") for t in types if "name" in t]
            log.info(f"Company types for ID [{company_id}]: {company_types}")
        else:
            log.warning(f"Unable to retrieve company types for ID [{company_id}]")
        return company_identifier, company_name, company_id, company_types
    elif ticket_response:
        log.error(f"Failed to retrieve ticket [{ticket_number}] Status: {ticket_response.status_code}, Body: {ticket_response.text}")
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

def get_aad_user_data(log, http_client, msgraph_base_url, user_identifier, token):
    log.info(f"Resolving user ID and email for [{user_identifier}]")
    headers = {"Authorization": f"Bearer {token}"}
    if re.fullmatch(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}", user_identifier):
        endpoint = f"{msgraph_base_url}/users/{user_identifier}"
        response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
        if response:
            user = response.json()
            return user.get("id", ""), user.get("userPrincipalName", ""), user.get("onPremisesSamAccountName", ""), user.get("onPremisesSyncEnabled", False)
        return "", "", "", False
    filters = [
        f"startswith(displayName,'{user_identifier}')",
        f"startswith(userPrincipalName,'{user_identifier}')",
        f"startswith(mail,'{user_identifier}')"
    ]
    filter_query = " or ".join(filters)
    endpoint = f"{msgraph_base_url}/users?$filter={urllib.parse.quote(filter_query)}"
    response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
    if response:
        users = response.json().get("value", [])
        if len(users) > 1:
            log.error(f"Multiple users found for [{user_identifier}]")
            return users
        if users:
            user = users[0]
            return user.get("id", ""), user.get("userPrincipalName", ""), user.get("onPremisesSamAccountName", ""), user.get("onPremisesSyncEnabled", False)
    log.error(f"Failed to resolve user ID and email for [{user_identifier}]")
    return "", "", "", False

def get_user_owned_teams(log, http_client, msgraph_base_url, user_id, token):
    log.info(f"Retrieving teams where user [{user_id}] is an owner")
    endpoint = f"{msgraph_base_url}/users/{user_id}/memberOf/microsoft.graph.group?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName"
    headers = {"Authorization": f"Bearer {token}"}
    response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
    if response:
        teams = response.json().get("value", [])
        owned_teams = []
        for team in teams:
            team_id = team.get("id")
            owners_endpoint = f"{msgraph_base_url}/groups/{team_id}/owners"
            owners_response = execute_api_call(log, http_client, "get", owners_endpoint, headers=headers)
            if owners_response and owners_response.status_code == 200:
                owners = owners_response.json().get("value", [])
                if any(owner.get("id") == user_id for owner in owners):
                    owned_teams.append({"id": team_id, "displayName": team.get("displayName")})
        log.info(f"User [{user_id}] owns [{len(owned_teams)}] teams")
        return owned_teams
    log.error(f"Failed to retrieve user teams for [{user_id}]")
    return []

def remove_user_as_owner(log, http_client, msgraph_base_url, team_id, user_id, token):
    log.info(f"Removing user [{user_id}] as owner from team [{team_id}]")
    endpoint = f"{msgraph_base_url}/groups/{team_id}/owners/{user_id}/$ref"
    headers = {"Authorization": f"Bearer {token}"}
    response = execute_api_call(log, http_client, "delete", endpoint, headers=headers)
    if response and response.status_code in [204, 200]:
        log.info(f"User [{user_id}] removed as owner from team [{team_id}]")
        return True
    log.error(f"Failed to remove user [{user_id}] as owner from team [{team_id}]")
    return False

def replace_team_owner(log, http_client, msgraph_base_url, team_id, new_owner_id, token):
    log.info(f"Checking team [{team_id}] for existing owners")
    headers = {"Authorization": f"Bearer {token}"}
    endpoint = f"{msgraph_base_url}/groups/{team_id}/owners"
    response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
    if response:
        owners = response.json().get("value", [])
        if owners:
            log.info(f"Team [{team_id}] already has an owner")
            return False
        add_endpoint = f"{msgraph_base_url}/groups/{team_id}/owners/$ref"
        payload = {"@odata.id": f"https://graph.microsoft.com/v1.0/users/{new_owner_id}"}
        add_response = execute_api_call(log, http_client, "post", add_endpoint, data=payload, headers=headers)
        if add_response and add_response.status_code in [204, 200]:
            log.info(f"New owner [{new_owner_id}] added to team [{team_id}]")
            return True
        log.error(f"Failed to add new owner [{new_owner_id}] to team [{team_id}]")
        return False
    log.error(f"Failed to fetch owners for team [{team_id}]")
    return False

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1743192826772")
            user_identifier = input.get_value("User_1738886551777")
            new_owner_identifier = input.get_value("NewOwner_1738886608412")
            auth_code = input.get_value("AuthCode_1743025274034")
            graph_token = input.get_value("GraphToken_xxxxxxxxxxxxx")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        user_identifier = user_identifier.strip() if user_identifier else ""
        new_owner_identifier = new_owner_identifier.strip() if new_owner_identifier else ""
        auth_code = auth_code.strip() if auth_code else ""
        graph_token = graph_token.strip() if graph_token else ""

        log.info(f"Ticket Number = [{ticket_number}]")
        log.info(f"Received input user = [{user_identifier}]")
        log.info(f"New owner = [{new_owner_identifier}]")

        if not user_identifier:
            record_result(log, ResultLevel.WARNING, "User identifier is empty or invalid")
            return
        if not new_owner_identifier:
            record_result(log, ResultLevel.WARNING, "New owner identifier is empty or invalid")
            return

        log.info(f"Retrieving company data for ticket [{ticket_number}]")
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

        aad_user_result = get_aad_user_data(log, http_client, msgraph_base_url, user_identifier, graph_access_token)
        if isinstance(aad_user_result, list):
            details = "\n".join([f"- {u.get('displayName')} | {u.get('userPrincipalName')} | {u.get('id')}" for u in aad_user_result])
            record_result(log, ResultLevel.WARNING, f"Multiple users found for [{user_identifier}]\n{details}")
            return
        user_id, user_email, user_sam, user_onpremisessyncenabled = aad_user_result

        if not user_id:
            record_result(log, ResultLevel.WARNING, f"Failed to resolve user ID for [{user_identifier}]")
            return
        if not user_email:
            record_result(log, ResultLevel.WARNING, f"Unable to resolve user principal name for [{user_identifier}]")
            return

        new_owner_result = get_aad_user_data(log, http_client, msgraph_base_url, new_owner_identifier, graph_access_token)
        if isinstance(new_owner_result, list):
            details = "\n".join([f"- {u.get('displayName')} | {u.get('userPrincipalName')} | {u.get('id')}" for u in new_owner_result])
            record_result(log, ResultLevel.WARNING, f"Multiple users found for new owner [{new_owner_identifier}]\n{details}")
            return
        new_owner_id, new_owner_email, new_owner_sam, new_owner_onpremisessyncenabled = new_owner_result

        if not new_owner_id:
            record_result(log, ResultLevel.WARNING, f"Failed to resolve user ID for [{user_identifier}]")
            return
        if not new_owner_email:
            record_result(log, ResultLevel.WARNING, f"Unable to resolve user principal name for [{user_identifier}]")
            return

        owned_teams = get_user_owned_teams(log, http_client, msgraph_base_url, user_id, graph_access_token)
        if not owned_teams:
            record_result(log, ResultLevel.SUCCESS, f"User [{user_email}] has no owner roles to remove")
            return

        updated_teams = []
        for team in owned_teams:
            removed = remove_user_as_owner(log, http_client, msgraph_base_url, team["id"], user_id, graph_access_token)
            if removed:
                replaced = replace_team_owner(log, http_client, msgraph_base_url, team["id"], new_owner_id, graph_access_token)
                updated_teams.append((team["displayName"], replaced))
                data_to_log[f"Team_{team['displayName']}"] = "Owner replaced" if replaced else "Owner added"

        success_summary = "\n".join([f"- {name}: {'New owner assigned' if r else 'Another owner already exists'}" for name, r in updated_teams])
        record_result(log, ResultLevel.SUCCESS, f"Processed Teams for [{user_email}]:\n{success_summary}")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
