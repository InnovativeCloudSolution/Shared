import sys
import random
import re
import os
import base64
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
dictionary_url = "https://mitazu1pubfilestore.blob.core.windows.net/automation/Password_safe_wordlist.txt"
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

def get_user_data(log, http_client, msgraph_base_url, user_identifier, token):
    log.info(f"Resolving user ID and email for [{user_identifier}]")
    headers = {"Authorization": f"Bearer {token}"}

    if re.fullmatch(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}", user_identifier):
        endpoint = f"{msgraph_base_url}/users/{user_identifier}"
        response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
        if response and response.status_code == 200:
            user = response.json()
            user_id = user.get("id")
            user_email = user.get("userPrincipalName", "")
            user_sam = user.get("onPremisesSamAccountName", "")
            log.info(f"User found by object ID [{user_id}] with email [{user_email}] and SAM [{user_sam}]")
            return user_id, user_email, user_sam
        log.error(f"No user found with object ID [{user_identifier}]")
        return "", "", ""

    filters = [
        f"startswith(displayName,'{user_identifier}')",
        f"startswith(userPrincipalName,'{user_identifier}')",
        f"startswith(mail,'{user_identifier}')"
    ]
    filter_query = " or ".join(filters)
    endpoint = f"{msgraph_base_url}/users?$filter={urllib.parse.quote(filter_query)}"
    response = execute_api_call(log, http_client, "get", endpoint, headers=headers)

    if response and response.status_code == 200:
        users = response.json().get("value", [])
        if len(users) > 1:
            log.error(f"Multiple users found for [{user_identifier}]")
            return users
        if users:
            user = users[0]
            user_id = user.get("id")
            user_email = user.get("userPrincipalName", "")
            user_sam = user.get("onPremisesSamAccountName", "")
            log.info(f"User found for [{user_identifier}] - ID: {user_id}, Email: {user_email}, SAM: {user_sam}")
            return user_id, user_email, user_sam

    log.error(f"Failed to resolve user ID and email for [{user_identifier}]")
    return "", "", ""

def get_random_word_list(log, http_client) -> list:
    log.info("Fetching word list from safe wordlist")

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

def reset_user_password(log, http_client, msgraph_base_url, user_id, token, new_password):
    log.info(f"Resetting password for user [{user_id}] in Azure AD")

    endpoint = f"{msgraph_base_url}/users/{user_id}"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    payload = {
        "passwordProfile": {
            "password": new_password,
            "forceChangePasswordNextSignIn": True
        }
    }

    response = execute_api_call(log, http_client, "patch", endpoint, data=payload, headers=headers)

    if response and response.status_code == 204:
        log.info(f"Password reset successfully for user [{user_id}]")
        return True
    else:
        log.error(f"Failed to reset password for user [{user_id}] Status code: {response.status_code if response else 'N/A'}")
        return False

def send_sms(log, http_client, recipient_number: str, country_name: str, password: str) -> bool:
    try:
        sender_id = "Mangano IT"
        username = get_secret_value(log, http_client, vault_name, "MIT-TransmitSMS-Username")
        password_secret = get_secret_value(log, http_client, vault_name, "MIT-TransmitSMS-Password")

        if not username or not password_secret:
            log.error("Missing TransmitSMS credentials from Key Vault")
            return False

        base_url = "https://api.transmitsms.com/send-sms.json"
        message = f"Your password has been set to the following:\n\n{password}\n\nIf you did not request this change, please contact Mangano IT service desk immediately on (07) 3151 9000"

        params = {
            "message": message,
            "to": recipient_number,
            "from": sender_id
        }

        if country_name:
            params["countrycode"] = country_name

        auth_str = f"{username}:{password_secret}"
        encoded_auth = base64.b64encode(auth_str.encode("utf-8")).decode("utf-8")
        headers = {
            "Authorization": f"Basic {encoded_auth}",
            "Content-Type": "application/x-www-form-urlencoded"
        }

        response = execute_api_call(
            log,
            http_client,
            method="post",
            endpoint=base_url,
            data=params,
            headers=headers
        )

        if response is None:
            log.error("Failed to send SMS: No response from API")
            return False

        try:
            resp_json = response.json()
            error_block = resp_json.get("error", {})

            if error_block.get("code") == "SUCCESS":
                log.info(f"SMS sent successfully: {resp_json}")
                return True
            else:
                description = error_block.get("description", "Unknown error")
                log.error(f"Failed to send SMS: {description}")
                return False
        except Exception:
            log.exception("Failed to parse response from TransmitSMS")
            return False

    except Exception as e:
        log.exception(f"Exception occurred while sending SMS: {str(e)}")
        return False

def main():
    try:
        try:
            user_identifier = input.get_value("User_1740658833537")
            word_count_str = input.get_value("WordCount_1741310987394")
            ticket_number = input.get_value("TicketNumber_1742953836504")
            mobile_number = input.get_value("MobileNumber_1741310437739")
            country_name = input.get_value("Country_1741360570964")
            auth_code = input.get_value("AuthCode_1743025274034")
            provided_token = input.get_value("AccessToken_1742865275651")
        except Exception as e:
            log.exception(e, "Failed to fetch input values")
            log.result_message(ResultLevel.FAILED, "Failed to fetch input values")
            return
        
        user_identifier = user_identifier.strip() if user_identifier else ""
        word_count_str = word_count_str.strip() if word_count_str else ""
        ticket_number = ticket_number.strip() if ticket_number else ""
        mobile_number = mobile_number.strip() if mobile_number else ""
        country_name = country_name.strip() if country_name else ""
        auth_code = auth_code.strip() if auth_code else ""
        provided_token = provided_token.strip() if provided_token else ""

        log.info(f"Received input user = [{user_identifier}], Country = [{country_name}], Mobile = [{mobile_number}], Word Count = [{word_count_str}]")

        if not user_identifier or not user_identifier.strip():
            log.error("User identifier is empty or invalid")
            log.result_message(ResultLevel.FAILED, "User identifier is empty or invalid")
            return

        if not word_count_str or not word_count_str.isdigit():
            log.error("Invalid or missing word count input")
            log.result_message(ResultLevel.FAILED, "Invalid or missing word count input")
            return

        word_count = int(word_count_str)
        password = generate_secure_password(log, http_client, word_count)

        if not password:
            log.error("Password generation failed")
            log.result_message(ResultLevel.FAILED, "Failed to generate password")
            return

        if provided_token and provided_token.strip():
            access_token = provided_token.strip()
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

        user_result = get_user_data(log, http_client, msgraph_base_url, user_identifier, access_token)

        if isinstance(user_result, list):
            details = "\n".join([f"- {u.get('displayName')} | {u.get('userPrincipalName')} | {u.get('id')}" for u in user_result])
            log.result_message(ResultLevel.FAILED, f"Multiple users found for [{user_identifier}]\n{details}")
            return

        user_id, user_email, _ = user_result
        if not user_id:
            log.result_message(ResultLevel.FAILED, f"Failed to resolve user ID for [{user_identifier}]")
            return
        if not user_email:
            log.result_message(ResultLevel.FAILED, f"Unable to resolve user email for [{user_identifier}]")
            return

        success = reset_user_password(log, http_client, msgraph_base_url, user_id, access_token, password)
        if success:
            log.result_message(ResultLevel.SUCCESS, f"Password reset successfully for user [{user_email}]")
            data_to_log = {"password": password}

            if mobile_number:
                sms_success = send_sms(log, http_client, mobile_number, country_name, password)
                if sms_success:
                    log.result_message(ResultLevel.SUCCESS, f"Password sent successfully to [{mobile_number}]")
                    data_to_log["mobile_number"] = mobile_number
                else:
                    log.result_message(ResultLevel.FAILED, f"Failed to send password to [{mobile_number}]")

            log.result_data(data_to_log)
        else:
            log.result_message(ResultLevel.FAILED, f"Failed to reset password for user [{user_email}]")

    except Exception:
        log.exception("An error occurred while processing")
        log.result_message(ResultLevel.FAILED, "Process failed")

if __name__ == "__main__":
    main()