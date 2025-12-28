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
msgraph_base_url_beta = "https://graph.microsoft.com/beta"
vault_name = "PLACEHOLDER-akv1"

data_to_log = {}
bot_name = "AAD - License Management"
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

    token = get_access_token(
        log, http_client, tenant_id, client_id, client_secret,
        scope="https://graph.microsoft.com/.default", log_prefix="Graph"
    )
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
        log.error(
            f"Failed to retrieve ticket [{ticket_number}] "
            f"Status: {ticket_response.status_code}, Body: {ticket_response.text}"
        )
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

    if re.fullmatch(
        r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}",
        user_identifier,
    ):
        endpoint = f"{msgraph_base_url}/users/{user_identifier}"
        response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
        if response:
            user = response.json()
            return (
                user.get("id", ""),
                user.get("userPrincipalName", ""),
                user.get("onPremisesSamAccountName", ""),
                user.get("onPremisesSyncEnabled", False)
            )
        return "", "", "", False

    filters = [
        f"startswith(displayName,'{user_identifier}')",
        f"startswith(userPrincipalName,'{user_identifier}')",
        f"startswith(mail,'{user_identifier}')",
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
            return (
                user.get("id", ""),
                user.get("userPrincipalName", ""),
                user.get("onPremisesSamAccountName", ""),
                user.get("onPremisesSyncEnabled", False)
            )
    log.error(f"Failed to resolve user ID and email for [{user_identifier}]")
    return "", "", "", False

def get_aad_group(log, http_client, msgraph_base_url, group_identifier, token):
    log.info(f"Resolving group name for [{group_identifier}]")
    headers = {"Authorization": f"Bearer {token}"}

    filters = [
        f"startswith(displayName,'{group_identifier}')",
        f"startswith(mail,'{group_identifier}')"
    ]
    filter_query = " or ".join(filters)
    endpoint = f"{msgraph_base_url}/groups?$filter={urllib.parse.quote(filter_query)}"

    response = execute_api_call(log, http_client, "get", endpoint, headers=headers)

    if response:
        group = response.json().get("value", [])
        if len(group) > 1:
            log.error(f"Multiple group found for [{group_identifier}]")
            return group
        if group:
            group = group[0]
            group_name = group.get("displayName", group.get("mail", ""))
            group_id = group.get("id", "")
            log.info(f"Group found for [{group_identifier}] - Name: {group_name}, ID: {group_id}")
            return group_name, group_id

    log.error(f"Failed to resolve group for [{group_identifier}]")
    return "", ""

def get_sku_name_map(log, http_client, msgraph_base_url, token):
    log.info("Fetching SKU ID to Name mapping with service plans")
    headers = {"Authorization": f"Bearer {token}"}
    endpoint = f"{msgraph_base_url}/subscribedSkus"

    response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
    sku_map = {}

    if response:
        for sku in response.json().get("value", []):
            sku_id = sku.get("skuId")
            sku_name = sku.get("skuPartNumber")
            service_plans = sku.get("servicePlans", [])
            plan_dict = {
                plan["servicePlanId"]: plan["servicePlanName"]
                for plan in service_plans
                if "servicePlanId" in plan and "servicePlanName" in plan
            }
            if sku_id and sku_name:
                sku_map[sku_id] = {
                    "name": sku_name,
                    "servicePlans": plan_dict
                }
        log.info(f"Retrieved {len(sku_map)} SKU entries with service plans")
    else:
        log.error(f"Failed to fetch SKU ID to Name mapping with service plans Status: {response.status_code if response else 'N/A'}")

    return sku_map

def get_aad_group_license(log, http_client, msgraph_base_url, group_identifier, token):
    sku_map = get_sku_name_map(log, http_client, msgraph_base_url, token)

    result = get_aad_group(log, http_client, msgraph_base_url, group_identifier, token)
    if isinstance(result, list):
        log.error(f"Multiple group found for [{group_identifier}]")
        return None
    group_name, group_id = result
    if not group_id:
        log.error(f"Could not resolve group [{group_identifier}]")
        return None

    headers = {"Authorization": f"Bearer {token}"}
    assigned_sku_id = ""
    assigned_license_name = ""
    service_plans = {}

    group_endpoint = f"{msgraph_base_url}/groups/{group_id}?$select=assignedLicenses"
    group_resp = execute_api_call(log, http_client, "get", group_endpoint, headers=headers)
    if group_resp and group_resp.status_code == 200:
        plans = group_resp.json().get("assignedLicenses", [])
        if plans:
            assigned_sku_id = plans[0].get("skuId", "")
            if assigned_sku_id in sku_map:
                assigned_license_name = sku_map[assigned_sku_id]["name"]
                service_plans = sku_map[assigned_sku_id]["servicePlans"]
                log.info(f"Group [{group_id}] assigned license: {assigned_license_name}")
            else:
                assigned_license_name = assigned_sku_id
        else:
            log.info(f"No licenses assigned to group [{group_id}]")
    else:
        log.error(f"Failed to fetch license assignment for group [{group_id}]")

    members_endpoint = f"{msgraph_base_url}/groups/{group_id}/members?$select=id"
    group_members = []
    while members_endpoint:
        member_resp = execute_api_call(log, http_client, "get", members_endpoint, headers=headers)
        if member_resp and member_resp.status_code == 200:
            data = member_resp.json()
            group_members += [m.get("id") for m in data.get("value", []) if "id" in m]
            members_endpoint = data.get("@odata.nextLink")
        else:
            log.error(f"Failed to fetch group members for group [{group_id}]")
            break

    group_member_set = set(group_members or [])
    member_count = len(group_member_set)
    log.info(f"Group [{group_id}] has [{member_count}] members")

    assigned_total = 0
    assigned_users_set = set()
    if assigned_sku_id:
        sku_resp = execute_api_call(log, http_client, "get", f"{msgraph_base_url}/subscribedSkus", headers=headers)
        if sku_resp and sku_resp.status_code == 200:
            for sku in sku_resp.json().get("value", []):
                if sku.get("skuId") == assigned_sku_id:
                    assigned_total = sku.get("prepaidUnits", {}).get("enabled", 0)
                    break
        else:
            log.error("Failed to retrieve license total count from M365")

        user_filter = f"assignedLicenses/any(s:s/skuId eq {assigned_sku_id})"
        user_endpoint = f"{msgraph_base_url}/users?$filter={urllib.parse.quote(user_filter)}&$select=id"
        while user_endpoint:
            user_resp = execute_api_call(log, http_client, "get", user_endpoint, headers=headers)
            if user_resp and user_resp.status_code == 200:
                data = user_resp.json()
                assigned_users_set.update({u.get("id") for u in data.get("value", []) if "id" in u})
                user_endpoint = data.get("@odata.nextLink")
            else:
                log.error(f"Failed to fetch users with SKU [{assigned_sku_id}]")
                break

    try:
        manual_user_ids = assigned_users_set - group_member_set
    except Exception as e:
        log.warning(f"Failed to compute manual license assignments: {str(e)}")
        manual_user_ids = set()

    manual_assigned = len(manual_user_ids)
    assigned = member_count + manual_assigned
    available_licenses = max(0, assigned_total - assigned)
    licenses_to_purchase = max(0, assigned - assigned_total + 1)

    return {
        "group": group_name,
        "group_id": group_id,
        "assigned_license": assigned_license_name,
        "member_count": member_count,
        "total_licenses": assigned_total,
        "available_licenses": available_licenses,
        "licenses_to_purchase": licenses_to_purchase,
        "manual_assigned": manual_assigned,
        "assigned": assigned,
        "assigned_sku_id": assigned_sku_id,
        "service_plans": list(service_plans.values())
    }

def remove_aad_user_licenses(log, http_client, msgraph_base_url, user_id, token, sku_map):
    log.info(f"Removing all licenses from user [{user_id}]")

    license_endpoint = f"{msgraph_base_url}/users/{user_id}/licenseDetails"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    response = execute_api_call(log, http_client, "get", license_endpoint, headers=headers)
    if response:
        license_data = response.json().get("value", [])
        if not license_data:
            log.info(f"No licenses assigned to user [{user_id}]")
            return True, ["No licenses assigned"]

        removed_skus = []
        removed_names = []

        for license in license_data:
            sku_id = license.get("skuId")
            if sku_id:
                removed_skus.append(sku_id)
                removed_names.append(sku_map.get(sku_id, {}).get("name", sku_id))

        if not removed_skus:
            log.info(f"No removable licenses found for user [{user_id}]")
            return True, ["No licenses assigned"]

        payload = {
            "addLicenses": [],
            "removeLicenses": removed_skus
        }

        assign_endpoint = f"{msgraph_base_url}/users/{user_id}/assignLicense"
        response = execute_api_call(log, http_client, "post", assign_endpoint, data=payload, headers=headers)

        if response:
            log.info(f"Successfully removed licenses from user [{user_id}] - Names: {removed_names}")
            return True, removed_names
        elif response and response.status_code == 400 and "license is inherited from a group" in response.text:
            log.warning(f"User [{user_id}] is assigned licenses via group membership and cannot be removed directly")
            return False, ["Group-based assignment"]
        else:
            
            return False, []
    else:
        
        return False, []

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1755468112824")
            operation = input.get_value("Operation_1755468114262")
            user_identifier = input.get_value("User_1755468118558")
            group = input.get_value("Groups_1755468123015")
            auth_code = input.get_value("AuthenticationCode_1755468119813")
            graph_token = input.get_value("GraphToken_1755468124859")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        operation = operation.strip() if operation else ""
        user_identifier = user_identifier.strip() if user_identifier else ""
        group = group.strip() if group else ""
        auth_code = auth_code.strip() if auth_code else ""
        graph_token = graph_token.strip() if graph_token else ""

        log.info(f"Received input user = [{user_identifier}], Ticket = [{ticket_number}], Action = [{operation}]")

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number is required but missing")
            return
        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation value is missing or invalid")
            return
        if not user_identifier and operation != "get group license":
            record_result(log, ResultLevel.WARNING, "User identifier is empty or invalid")
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

        if operation == "Get Group License Information":
            group_list = [g.strip() for g in group.split(",") if g.strip()]
            for group in group_list:
                if not group:
                    record_result(log, ResultLevel.WARNING, "Group identifier is empty or invalid")
                    continue

                result = get_aad_group_license(log, http_client, msgraph_base_url, group, graph_access_token)
                if result:
                    data_to_log.update({
                        "group_name": result["group"],
                        "group_license": result["assigned_license"],
                        "member_count": result["member_count"],
                        "total_licenses": result["total_licenses"],
                        "available_licenses": result["available_licenses"],
                        "licenses_to_purchase": result["licenses_to_purchase"],
                        "manual_assigned": result["manual_assigned"],
                        "assigned": result["assigned"]
                    })

                    group = result["group"]
                    license_name = result["assigned_license"]
                    members = result["member_count"]
                    total = result["total_licenses"]
                    manual = result["manual_assigned"]
                    assigned = result["assigned"]
                    available = result["available_licenses"]
                    purchase = result["licenses_to_purchase"]

                    if not result["assigned_sku_id"]:
                        record_result(log, ResultLevel.WARNING, f"Group [{group}] has no license assigned. Cannot determine provisioning status.")
                    elif assigned + 1 > total:
                        record_result(log, ResultLevel.WARNING, f"Group [{group}] is assigned license [{license_name}] with {total} total and {assigned} assigned. There are {members} members in the group and {manual} manually assigned licenses. Adding 1 user will require purchasing {purchase} license(s).")
                    elif available == 1:
                        record_result(log, ResultLevel.INFO, f"Group [{group}] is assigned license [{license_name}] with {total} total and {assigned} assigned. There are {members} members in the group and {manual} manually assigned licenses. Next user will consume the last available license.")
                    elif available > 1:
                        record_result(log, ResultLevel.INFO, f"Group [{group}] is assigned license [{license_name}] with {total} total and {assigned} assigned. There are {members} members in the group and {manual} manually assigned licenses. {available} spare licenses available.")
                else:
                    record_result(log, ResultLevel.WARNING, f"Failed to analyze license group [{group}]")

        elif operation == "Remove All Licenses":
            user_id, user_email, user_sam, user_sync = get_aad_user_data(log, http_client, msgraph_base_url, user_identifier, graph_access_token)
            if not user_id:
                record_result(log, ResultLevel.WARNING, f"Could not resolve user [{user_identifier}]")
                return

            sku_map = get_sku_name_map(log, http_client, msgraph_base_url, graph_access_token)
            success, removed = remove_aad_user_licenses(log, http_client, msgraph_base_url, user_id, graph_access_token, sku_map)

            if success:
                removed_str = ", ".join(removed) if removed else "No licenses were assigned"
                record_result(log, ResultLevel.SUCCESS, f"Successfully removed licenses from user [{user_identifier}]: {removed_str}")
            elif removed and "Group-based assignment" in removed:
                record_result(log, ResultLevel.WARNING, f"User [{user_identifier}] is assigned licenses via group membership and they cannot be removed directly")
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to remove licenses from user [{user_identifier}]")
        
        else:
            record_result(log, ResultLevel.WARNING, f"Unknown operation [{operation}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()