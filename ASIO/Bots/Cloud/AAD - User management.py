import sys
import json
import random
import re
import os
import time
import urllib.parse
import requests
from datetime import datetime, timedelta, timezone
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
bot_name = "AAD - User management"
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

def post_ticket_note(log, http_client, cwpsa_base_url, ticket_number, note_type, note):
    log.info(f"Posting {note_type} note to ticket [{ticket_number}]")
    note_endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}/notes"
    payload = {
        "text": note,
        "detailDescriptionFlag": False,
        "internalAnalysisFlag": False,
        "resolutionFlag": False,
        "issueFlag": False,
        "internalFlag": False,
        "externalFlag": False,
        "contact": {
            "id": 15655
        }
    }
    if note_type == "discussion":
        payload["detailDescriptionFlag"] = True
    elif note_type == "internal":
        payload["internalAnalysisFlag"] = True
    note_response = execute_api_call(log, http_client, "post", note_endpoint, integration_name="cw_psa", data=payload)
    if note_response and note_response.status_code == 200:
        log.info(f"{note_type} note posted successfully to ticket [{ticket_number}]")
        return True
    else:
        log.error(f"Failed to post {note_type} note to ticket [{ticket_number}] Status: {note_response.status_code}, Body: {note_response.text}")
    return False

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

def get_random_word_list(log, http_client) -> list:
    log.info("Fetching word list from safe wordlist")
    dictionary_url = "https://mitazu1pubfilestore.blob.core.windows.net/automation/Password_safe_wordlist.txt"
    response = execute_api_call(log, http_client, "get", dictionary_url)

    if not response:
        log.error("API call to fetch word list failed. No response received.")
        return []

    if response.status_code != 200:
        log.error(f"Failed to retrieve word list, HTTP Status: {response.status_code}, Response: {response.text}")
        return []

    try:
        lines = response.text.splitlines()
        words = [line.strip().lower() for line in lines if line.strip()]
    except Exception as e:
        log.exception("Failed to parse safe word list")
        return []

    log.info(f"Total words retrieved: {len(words)}")

    if not words:
        log.error("No words found, cannot generate password.")
        return []

    return words

def generate_secure_password(log, http_client, word_count: int) -> str:
    log.info(f"Generating password with [{word_count}] words")
    word_list = get_random_word_list(log, http_client)

    if len(word_list) < word_count:
        log.error("Not enough words to generate password")
        return ""

    words = random.sample(word_list, word_count)
    cap_index = random.randint(0, word_count - 1)
    words[cap_index] = words[cap_index].capitalize()

    password_base = "-".join(words)
    digits = f"{random.randint(0, 99):02d}"
    password = f"{password_base}-{digits}"

    log.info("Password generated successfully")
    return password

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
    return "", "", "", False

def create_aad_user(log, http_client, msgraph_base_url, user_details, access_token):
    log.info("Creating new user in Azure AD.")

    sanitized_user_details = {k: v for k, v in user_details.items() if "password" not in k.lower()}
    log.info("User details received (passwords redacted).")
    log.info(json.dumps(sanitized_user_details, indent=2))
    force_password_change = True if user_details.get("user_password_reset_required") == "Yes" else False
    user_hire_raw = user_details.get("user_employee_hiredate", "")

    try:
        if "T" in user_hire_raw:
            formatted_user_employee_hiredate = user_hire_raw
        else:
            formatted_user_employee_hiredate = datetime.strptime(user_hire_raw, "%Y-%m-%d").strftime("%Y-%m-%dT%H:%M:%SZ")
    except Exception:
        formatted_user_employee_hiredate = None
        log.warning("Invalid or missing employee hire date format.")

    custom_displayname = user_details.get("user_displayname", "").strip()
    if custom_displayname:
        display_name = custom_displayname
    else:
        display_name = f"{user_details.get('user_firstname', '')} {user_details.get('user_lastname', '')}".strip()
    
    payload = {
        "accountEnabled": True,
        "displayName": display_name,
        "userPrincipalName": user_details.get("user_upn", ""),
        "mailNickname": user_details.get("user_mailnickname", ""),
        "passwordProfile": {
            "forceChangePasswordNextSignIn": force_password_change,
            "password": user_details.get("user_password", ""),
        },
        "userType": user_details.get("user_type", "Member")
    }

    if user_details.get("exchange_primary_smtp_address"):
        payload["mail"] = user_details.get("exchange_primary_smtp_address")
    if user_details.get("user_firstname"):
        payload["givenName"] = user_details.get("user_firstname")
    if user_details.get("user_lastname"):
        payload["surname"] = user_details.get("user_lastname")
    if user_details.get("organisation_department"):
        payload["department"] = user_details.get("organisation_department")
    if user_details.get("organisation_title"):
        payload["jobTitle"] = user_details.get("organisation_title")
    if user_details.get("user_employee_id"):
        payload["employeeId"] = user_details.get("user_employee_id")
    if user_details.get("user_employee_type"):
        payload["employeeType"] = user_details.get("user_employee_type")
    if user_details.get("organisation_company"):
        payload["companyName"] = user_details.get("organisation_company")
    if user_details.get("organisation_site_office"):
        payload["officeLocation"] = user_details.get("organisation_site_office")
    if user_details.get("organisation_site_streetaddress"):
        payload["streetAddress"] = user_details.get("organisation_site_streetaddress")
    if user_details.get("organisation_site_city"):
        payload["city"] = user_details.get("organisation_site_city")
    if user_details.get("organisation_site_state"):
        payload["state"] = user_details.get("organisation_site_state")
    if user_details.get("organisation_site_zip"):
        payload["postalCode"] = user_details.get("organisation_site_zip")
    if user_details.get("organisation_site_country"):
        payload["country"] = user_details.get("organisation_site_country")
    if user_details.get("user_mobile_personal"):
        payload["mobilePhone"] = user_details.get("user_mobile_personal")
    if user_details.get("user_business_phone"):
        payload["businessPhones"] = [user_details.get("user_business_phone")]
    if user_details.get("user_fax_number"):
        payload["faxNumber"] = user_details.get("user_fax_number")
    if formatted_user_employee_hiredate:
        payload["employeeHireDate"] = formatted_user_employee_hiredate
    
    # Sanitise payload to avoid logging sensitive information
    sanitized_payload = {k: v for k, v in payload.items() if k != "passwordProfile"}
    log.info("Creating user with payload (password redacted).")
    log.info(json.dumps(sanitized_payload, indent=2))

    # Generaet headers for the API call
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    endpoint = f"{msgraph_base_url}/users"
    response = execute_api_call(log, http_client, "post", endpoint, data=payload, headers=headers)

    if response and response.status_code == 201:
        user_id = response.json().get("id")
        user_upn = response.json().get("userPrincipalName")
        log.info(f"User [{user_id}]:[{payload.get('userPrincipalName')}] created successfully")
        return response, payload, user_id, user_upn
    else:
        
        return None, None, None, None

def update_aad_user(log, http_client, msgraph_base_url, user_id, user_details, access_token):
    log.info(f"Updating Azure AD user with ID [{user_id}]")

    payload = {}

    custom_displayname = user_details.get("user_displayname", "").strip()
    if custom_displayname:
        payload["displayName"] = custom_displayname
    elif user_details.get("user_firstname") or user_details.get("user_lastname"):
        payload["displayName"] = f"{user_details.get('user_firstname', '')} {user_details.get('user_lastname', '')}".strip()

    if user_details.get("exchange_primary_smtp_address"):
        payload["mail"] = user_details.get("exchange_primary_smtp_address")
    if user_details.get("user_firstname"):
        payload["givenName"] = user_details.get("user_firstname")
    if user_details.get("user_lastname"):
        payload["surname"] = user_details.get("user_lastname")
    if user_details.get("organisation_department"):
        payload["department"] = user_details.get("organisation_department")
    if user_details.get("organisation_title"):
        payload["jobTitle"] = user_details.get("organisation_title")
    if user_details.get("user_employee_id"):
        payload["employeeId"] = user_details.get("user_employee_id")
    if user_details.get("user_employee_type"):
        payload["employeeType"] = user_details.get("user_employee_type")
    if user_details.get("organisation_company"):
        payload["companyName"] = user_details.get("organisation_company")
    if user_details.get("organisation_site_office"):
        payload["officeLocation"] = user_details.get("organisation_site_office")
    if user_details.get("organisation_site_streetaddress"):
        payload["streetAddress"] = user_details.get("organisation_site_streetaddress")
    if user_details.get("organisation_site_city"):
        payload["city"] = user_details.get("organisation_site_city")
    if user_details.get("organisation_site_state"):
        payload["state"] = user_details.get("organisation_site_state")
    if user_details.get("organisation_site_zip"):
        payload["postalCode"] = user_details.get("organisation_site_zip")
    if user_details.get("organisation_site_country"):
        payload["country"] = user_details.get("organisation_site_country")
    if user_details.get("user_mobile_personal"):
        payload["mobilePhone"] = user_details.get("user_mobile_personal")
    if user_details.get("user_business_phone"):
        payload["businessPhones"] = [user_details.get("user_business_phone")]
    if user_details.get("user_fax_number"):
        payload["faxNumber"] = user_details.get("user_fax_number")

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    endpoint = f"{msgraph_base_url}/users/{user_id}"

    log.info("Updating user with the following attributes:")
    log.info(json.dumps(payload, indent=2))

    response = execute_api_call(log, http_client, "patch", endpoint, data=payload, headers=headers)
    if response:
        log.info(f"User [{user_id}]:[{payload.get('userPrincipalName')}] updated successfully")
        return True
    return False

def update_aad_user_manager(log, http_client, msgraph_base_url, user_id, manager_upn, access_token):
    log.info(f"Updating manager for user ID [{user_id}] to [{manager_upn}]")
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    endpoint = f"{msgraph_base_url}/users/{user_id}/manager/$ref"

    manager_result = get_aad_user_data(log, http_client, msgraph_base_url, manager_upn, access_token)
    manager_id, manager_email, manager_sam, manager_onpremisessyncenabled = manager_result
    if not manager_id:
        log.error(f"Manager user [{manager_upn}] not found") 
        return False
    else:
        log.info(f"Manager user found: ID [{manager_id}], Email [{manager_email}], SAM [{manager_sam}]")
        log.info(f"{msgraph_base_url_beta}/users/{manager_id}")
        log.info(f"{msgraph_base_url}/users/{user_id}/manager/$ref")

    payload = {
        "@odata.id": f"{msgraph_base_url_beta}/users/{manager_id}"
    }
    response = execute_api_call(log, http_client, "put", endpoint, data=payload, headers=headers)
    if response:
        log.info(f"Manager updated successfully for user ID [{user_id}]")
        return True
    return False

def disable_aad_user(log, http_client, msgraph_base_url, user_id, access_token):
    log.info(f"Disabling Azure AD user with ID [{user_id}]")

    endpoint = f"{msgraph_base_url}/users/{user_id}"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    payload = {
        "accountEnabled": False
    }

    response = execute_api_call(log, http_client, "patch", endpoint, data=payload, headers=headers)
    if response:
        log.info(f"User [{user_id}] disabled successfully")
        return True
    return False

def revoke_aad_user_sessions(log, http_client, msgraph_base_url, user_id, access_token):
    log.info(f"Revoking sessions for user ID [{user_id}]")

    endpoint = f"{msgraph_base_url}/users/{user_id}/revokeSignInSessions"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = execute_api_call(log, http_client, "post", endpoint, headers=headers)
    if response:
        log.info(f"Sessions successfully revoked for user ID [{user_id}]")
        return True
    return False

def reset_aad_user_password(log, http_client, msgraph_base_url, user_id, new_password, force_reset, access_token):
    log.info(f"Resetting password for user ID [{user_id}]")

    payload = {
        "passwordProfile": {
            "password": new_password,
            "forceChangePasswordNextSignIn": force_reset
        }
    }

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    endpoint = f"{msgraph_base_url}/users/{user_id}"

    log.info("Resetting user password with payload (password redacted)")
    log.info(json.dumps({"forceChangePasswordNextSignIn": force_reset}, indent=2))

    response = execute_api_call(log, http_client, "patch", endpoint, data=payload, headers=headers)
    if response:
        log.info(f"Password reset successfully for user ID [{user_id}]")
        return True
    return False

def revoke_mfa_aad_methods(log, http_client, msgraph_base_url, user_id, access_token):
    log.info(f"Revoking MFA methods for user ID [{user_id}]")
    headers = {"Authorization": f"Bearer {access_token}"}
    endpoint = f"{msgraph_base_url}/users/{user_id}/authentication/methods"

    response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
    if not response:
        
        return [], [f"Failed to retrieve MFA methods for user ID [{user_id}]"]

    methods = response.json().get("value", [])
    success_messages = []
    failure_messages = []

    type_segment_map = {
        "emailAuthenticationMethod": "emailMethods",
        "fido2AuthenticationMethod": "fido2Methods",
        "microsoftAuthenticatorAuthenticationMethod": "microsoftAuthenticatorMethods",
        "passwordAuthenticationMethod": "passwordMethods",
        "phoneAuthenticationMethod": "phoneMethods",
        "platformCredentialAuthenticationMethod": "passwordlessMicrosoftAuthenticatorMethods",
        "softwareOathAuthenticationMethod": "softwareOathMethods",
        "temporaryAccessPassAuthenticationMethod": "temporaryAccessPassMethods",
        "windowsHelloForBusinessAuthenticationMethod": "windowsHelloForBusinessMethods"
    }

    for method in methods:
        method_id = method.get("id")
        odata_type = method.get("@odata.type", "").split(".")[-1]
        path_segment = type_segment_map.get(odata_type)

        if not method_id or not path_segment:
            msg = f"Skipped MFA method with unknown or unsupported type [{odata_type}]"
            log.warning(msg)
            success_messages.append(msg)
            continue

        if odata_type == "passwordAuthenticationMethod":
            msg = f"Skipped deletion of password method [{method_id}] - not supported by Graph"
            log.info(msg)
            success_messages.append(msg)
            continue

        del_endpoint = f"{msgraph_base_url_beta}/users/{user_id}/authentication/{path_segment}/{method_id}"
        del_response = execute_api_call(log, http_client, "delete", del_endpoint, headers=headers)

        if del_response is None:
            msg = f"Failed to delete MFA method [{method_id}] of type [{odata_type}]"
            log.warning(msg)
            failure_messages.append(msg)
        elif del_response.status_code == 204:
            msg = f"MFA method [{method_id}] of type [{odata_type}] deleted successfully"
            log.info(msg)
            success_messages.append(msg)
        elif del_response.status_code == 400 and "default authentication method" in del_response.text:
            msg = f"Skipped default MFA method [{method_id}] of type [{odata_type}]"
            log.warning(msg)
            success_messages.append(msg)
        else:
            msg = f"Failed to delete MFA method [{method_id}] of type [{odata_type}] Status: {del_response.status_code}"
            log.warning(msg)
            failure_messages.append(msg)

    return success_messages, failure_messages

def invite_aad_guest_user(log, http_client, msgraph_base_url, user_details, access_token):
    log.info("Creating guest user in Azure AD.")

    payload = {
        "invitedUserEmailAddress": user_details.get("user_email_personal"),
        "inviteRedirectUrl": "https://myapp.microsoft.com"
    }

    # Generaet headers for the API call
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    endpoint = f"{msgraph_base_url}/invitations"
    response = execute_api_call(log, http_client, "post", endpoint, data=payload, headers=headers)

    if response and response.status_code == 201:
        user_id = response.json().get("id")
        user_upn = response.json().get("userPrincipalName")
        log.info(f"User [{user_id}]:[{payload.get('userPrincipalName')}] created successfully")
        return response, payload, user_id, user_upn
    else:
        
        return None, None, None, None

def main():
    try:
        try:
            user_details = {
                "cwpsa_ticket": input.get_value("TicketNumber_1755229310571"),
                "operation": input.get_value("Operation_1755229313056"),
                "user_firstname": input.get_value("FirstName_1755229324604"),
                "user_lastname": input.get_value("LastName_1755229317310"),
                "user_displayname": input.get_value("DisplayName_1758751385535"),
                "user_upn": input.get_value("UserPrincipalName_1755229325857"),
                "user_mailnickname": input.get_value("MailNickname_1755229328967"),
                "exchange_primary_smtp_address": input.get_value("PrimarySMTPAddress_1755229327394"),
                "user_password": input.get_value("Password_1755229330554"),
                "user_password_reset_required": input.get_value("ForcePasswordReset_1755229725998"),
                "user_type": input.get_value("UserType_1755229723422"),
                "organisation_title": input.get_value("JobTitle_1755229690740"),
                "organisation_company": input.get_value("Company_1755230117512"),
                "organisation_department": input.get_value("Department_1755230119642"),
                "organisation_manager": input.get_value("Manager_1755230212729"),
                "organisation_site_office": input.get_value("Office_1755230248167"),
                "organisation_site_streetaddress": input.get_value("StreetAddress_1755230250480"),
                "organisation_site_city": input.get_value("City_1755230581747"),
                "organisation_site_state": input.get_value("State_1755230584231"),
                "organisation_site_zip": input.get_value("Postcode_1755230587179"),
                "organisation_site_country": input.get_value("Country_1755230593077"),
                "user_mobile_personal": input.get_value("PersonalMobile_1755230596035"),
                "user_email_personal": input.get_value("PersonalEmailAddress_1757310272266"),
                "user_business_phone": input.get_value("BusinessPhone_1755230613923"),
                "user_fax_number": input.get_value("Fax_1755230617062"),
                "user_employee_id": input.get_value("EmployeeID_1755230963053"),
                "user_employee_type": input.get_value("EmployeeType_1755230965799"),
                "user_employee_hiredate": input.get_value("EmployeeHireDate_1755230971132"),
                "post_discussion_note": input.get_value("PostDiscussionNote_1763353203051"),
                "auth_code": input.get_value("AuthenticationCode_1755230968264"),
                "graph_token": input.get_value("GraphToken_1755230974175")
            }
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        for k in user_details:
            user_details[k] = str(user_details[k]).strip() if user_details[k] is not None else ""

        user_details["user_password_reset_required"] = "Yes" if user_details["user_password_reset_required"] == "Yes" else "No"

        user_identifier = user_details["user_upn"]
        operation = user_details["operation"]
        cwpsa_ticket = user_details["cwpsa_ticket"]
        auth_code = user_details["auth_code"]
        graph_token = user_details["graph_token"]

        log.info(f"Received input user = [{user_identifier}], Ticket = [{cwpsa_ticket}], Operation = [{operation}]")

        if not cwpsa_ticket:
            record_result(log, ResultLevel.WARNING, "Ticket number is required but missing")
            return
        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation value is missing or invalid")
            return
        if operation == "Invite Guest User":
            if not user_details.get("user_email_personal"):
                record_result(log, ResultLevel.WARNING, "User email personal is empty or invalid")
                return

        log.info(f"Retrieving company data for ticket [{cwpsa_ticket}]")
        company_identifier, company_name, company_id, company_types = get_company_data_from_ticket(log, http_client, cwpsa_base_url, cwpsa_ticket)
        if not company_identifier:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve company identifier from ticket [{cwpsa_ticket}]")
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

        if operation != "Invite Guest User":
            user_result = get_aad_user_data(log, http_client, msgraph_base_url, user_identifier, graph_access_token)
            if isinstance(user_result, list):
                details = "\n".join([f"- {u.get('displayName')} | {u.get('userPrincipalName')} | {u.get('id')}" for u in user_result])
                record_result(log, ResultLevel.WARNING, f"Multiple users found for [{user_identifier}]\n{details}")
                return
            user_id, user_email, user_sam, user_onpremisessyncenabled = user_result
        else:
            user_email = user_details.get("user_email_personal")


        if operation == "Create User":
            if user_id:
                record_result(log, ResultLevel.WARNING, f"AAD user [{user_email}] already exists")
                return

            response, created_payload, created_aad_user, created_aad_user_upn = create_aad_user(log, http_client, msgraph_base_url, user_details, graph_access_token)
            if response:
                if user_details.get("organisation_manager"):
                    updated = update_aad_user_manager(log, http_client, msgraph_base_url, created_aad_user, user_details.get("organisation_manager"), graph_access_token)
                    if updated:
                        post_ticket_note(log, http_client, cwpsa_base_url, cwpsa_ticket, "discussion", f"AAD user [{user_identifier}] created and manager updated successfully.") if user_details.get("post_discussion_note") == "Yes" else None
                        record_result(log, ResultLevel.SUCCESS, f"AAD user [{user_identifier}] created and manager updated successfully.")
                    else:
                        post_ticket_note(log, http_client, cwpsa_base_url, cwpsa_ticket, "discussion", f"AAD user [{user_identifier}] created but failed to update manager.") if user_details.get("post_discussion_note") == "Yes" else None
                        record_result(log, ResultLevel.SUCCESS, f"AAD user [{user_identifier}] created but failed to update manager.")
                else:
                    post_ticket_note(log, http_client, cwpsa_base_url, cwpsa_ticket, "discussion", f"AAD user [{user_identifier}] created successfully.") if user_details.get("post_discussion_note") == "Yes" else None
                    record_result(log, ResultLevel.SUCCESS, f"AAD user [{user_identifier}] created successfully.")
                data_to_log['operation'] = operation
                data_to_log['user_id'] = created_aad_user
                data_to_log['user_upn'] = created_aad_user_upn
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to create AAD user [{user_identifier}].")
                data_to_log['operation'] = operation

        elif operation == "Update User":
            if not user_id:
                record_result(log, ResultLevel.WARNING, f"Cannot update user: [{user_identifier}] does not exist")
            else:
                updated = update_aad_user(log, http_client, msgraph_base_url, user_id, user_details, graph_access_token)
                manager_upn = user_details.get("organisation_manager", "")

                if updated:
                    if manager_upn:
                        manager_updated = update_aad_user_manager(log, http_client, msgraph_base_url, user_id, manager_upn, graph_access_token)
                        if manager_updated:
                            post_ticket_note(log, http_client, cwpsa_base_url, cwpsa_ticket, "discussion", f"AAD user [{user_identifier}] updated and manager set successfully.") if user_details.get("post_discussion_note") == "Yes" else None
                            record_result(log, ResultLevel.SUCCESS, f"AAD user [{user_identifier}] updated and manager set successfully")
                        else:
                            post_ticket_note(log, http_client, cwpsa_base_url, cwpsa_ticket, "discussion", f"AAD user [{user_identifier}] updated but failed to set manager.") if user_details.get("post_discussion_note") == "Yes" else None
                            record_result(log, ResultLevel.SUCCESS, f"AAD user [{user_identifier}] updated but failed to set manager")
                    else:
                        post_ticket_note(log, http_client, cwpsa_base_url, cwpsa_ticket, "discussion", f"AAD user [{user_identifier}] updated successfully.") if user_details.get("post_discussion_note") == "Yes" else None
                        record_result(log, ResultLevel.SUCCESS, f"AAD user [{user_identifier}] updated successfully")
                else:
                    record_result(log, ResultLevel.WARNING, f"Failed to update AAD user [{user_identifier}]")

            data_to_log['operation'] = operation

        elif operation == "Disable User":
            if not user_id:
                record_result(log, ResultLevel.WARNING, f"Cannot disable user: [{user_identifier}] does not exist")
            else:
                disabled = disable_aad_user(log, http_client, msgraph_base_url, user_id, graph_access_token)
                if disabled:
                    post_ticket_note(log, http_client, cwpsa_base_url, cwpsa_ticket, "discussion", f"AAD user [{user_identifier}] disabled successfully.") if user_details.get("post_discussion_note") == "Yes" else None
                    record_result(log, ResultLevel.SUCCESS, f"AAD user [{user_identifier}] disabled successfully")
                else:
                    record_result(log, ResultLevel.WARNING, f"Failed to disable AAD user [{user_identifier}]")
            data_to_log['operation'] = operation

        elif operation == "Revoke Sessions":
            if not user_id:
                record_result(log, ResultLevel.WARNING, f"Cannot revoke sessions: [{user_identifier}] does not exist")
            else:
                revoked = revoke_aad_user_sessions(log, http_client, msgraph_base_url, user_id, graph_access_token)
                if revoked:
                    post_ticket_note(log, http_client, cwpsa_base_url, cwpsa_ticket, "discussion", f"Sessions revoked for AAD user [{user_identifier}].") if user_details.get("post_discussion_note") == "Yes" else None
                    record_result(log, ResultLevel.SUCCESS, f"Sessions revoked for AAD user [{user_identifier}]")
                else:
                    record_result(log, ResultLevel.WARNING, f"Failed to revoke sessions for AAD user [{user_identifier}]")
            data_to_log['operation'] = operation

        elif operation == "Reset User Password":
            if not user_id:
                record_result(log, ResultLevel.WARNING, f"Cannot reset password: [{user_identifier}] does not exist")
            else:
                new_password = user_details.get("user_password", "")
                force_reset = user_details["user_password_reset_required"] == "Yes"

                if not new_password:
                    new_password = generate_secure_password(log, http_client, 3)
                    if not new_password:
                        record_result(log, ResultLevel.WARNING, "Password generation failed")
                        data_to_log['operation'] = operation
                        return

                reset = reset_aad_user_password(log, http_client, msgraph_base_url, user_id, new_password, force_reset, graph_access_token)

                if reset:
                    post_discussion_note(log, http_client, cwpsa_base_url, cwpsa_ticket, f"Password reset for AAD user [{user_identifier}] successfully.") if user_details.get("post_discussion_note") == "Yes" else None
                    record_result(log, ResultLevel.SUCCESS, f"Password reset for AAD user [{user_identifier}]")
                else:
                    record_result(log, ResultLevel.WARNING, f"Failed to reset password for AAD user [{user_identifier}]")

            data_to_log['operation'] = operation

        elif operation == "Reset MFA":
            if not user_id:
                record_result(log, ResultLevel.WARNING, f"Cannot revoke MFA: [{user_identifier}] does not exist")
            else:
                success_msgs, failure_msgs = revoke_mfa_aad_methods(log, http_client, msgraph_base_url, user_id, graph_access_token)
                for msg in success_msgs:
                    post_discussion_note(log, http_client, cwpsa_base_url, cwpsa_ticket, f"MFA methods revoked for AAD user [{user_identifier}] successfully.") if user_details.get("post_discussion_note") == "Yes" else None
                    record_result(log, ResultLevel.SUCCESS, msg)
                for msg in failure_msgs:
                    record_result(log, ResultLevel.WARNING, msg)

        elif operation == "Invite Guest User":
            if not user_email:
                record_result(log, ResultLevel.WARNING, f"Cannot invite guest user as personal email address was not provided.")
            else:
                response, invited_payload, invited_aad_user, invited_aad_user_upn = invite_aad_guest_user(log, http_client, msgraph_base_url, user_details, graph_access_token)
                if response:
                    post_discussion_note(log, http_client, cwpsa_base_url, cwpsa_ticket, f"Guest user [{user_email}] invited successfully.") if user_details.get("post_discussion_note") == "Yes" else None
                    record_result(log, ResultLevel.SUCCESS, f"Guest user [{user_email}] invited successfully")
                else:
                    record_result(log, ResultLevel.WARNING, f"Failed to invite guest user [{user_email}]")
            data_to_log['operation'] = operation
            data_to_log['user_id'] = invited_aad_user
            data_to_log['user_upn'] = invited_aad_user_upn

        else:
            record_result(log, ResultLevel.WARNING, f"Unknown operation [{operation}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
