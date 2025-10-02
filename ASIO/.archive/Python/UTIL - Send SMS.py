import sys
import random
import re
import os
import base64
import time
import urllib.parse
import requests
import codecs
from datetime import datetime, timedelta, timezone
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

cwpsa_base_url = "https://au.myconnectwise.net/v4_6_release/apis/3.0"
msgraph_base_url = "https://graph.microsoft.com/v1.0"
vault_name = "mit-azu1-prod1-akv1"
data_to_log = {}
log.info("Static variables set")

def record_result(log, level, message):
    log.result_message(level, message)

    if level == ResultLevel.WARNING:
        data_to_log["Result"] = "Fail"
    elif level == ResultLevel.SUCCESS:
        if "Result" not in data_to_log:
            data_to_log["Result"] = "Success"

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

def get_company_name_from_ticket(log, http_client, cwpsa_base_url, ticket_number):
    log.info(f"Retrieving company name for ticket [{ticket_number}]")
    endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}"

    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")

    if response:
        if response.status_code == 200:
            data = response.json()
            company_name = data.get("company", {}).get("name", "")
            if company_name:
                log.info(f"Company name for ticket [{ticket_number}] is [{company_name}]")
                return company_name
            else:
                log.error(f"Company name not found in response for ticket [{ticket_number}]")
        else:
            log.error(f"Failed to retrieve company name for ticket [{ticket_number}] Status: {response.status_code}, Body: {response.text}")
    else:
        log.error(f"Failed to retrieve company name for ticket [{ticket_number}]: No response received")

    return ""

def format_transmitsms_datetime(raw_date: str, raw_time: str, log) -> str | None:
    try:
        if not raw_date or not re.match(r"^\d{4}-\d{2}-\d{2}$", raw_date):
            log.warning(f"Invalid or missing schedule date format: [{raw_date}] - skipping scheduling")
            return None

        dt = datetime.strptime(raw_date, "%Y-%m-%d")

        if raw_time and re.match(r"^\d{4}$", raw_time.strip()):
            hour = int(raw_time[:2])
            minute = int(raw_time[2:])
            dt = dt.replace(hour=hour, minute=minute)
        else:
            dt = dt.replace(hour=0, minute=0, second=0)

        dt_utc = dt - timedelta(hours=10)
        formatted = dt_utc.strftime("%Y-%m-%d %H:%M:%S")
        log.info(f"Scheduled time converted to UTC: {formatted}")
        return formatted

    except Exception:
        log.warning("Failed to convert schedule date/time - skipping scheduling")
        return None

def build_user_onboarding_sms(first_name, company_name, password):
    return (
        f"Hi {first_name}! Mangano IT here,\n\n"
        f"We are {company_name}'s IT service provider. Here is your password:\n\n"
        f"{password}\n\n"
        f"Please keep this handy as you will need it on your first day!\n\n"
        f"Thanks"
    )

def build_password_reset_sms(password):
    return (
        f"Your password has been set to the following:\n\n"
        f"{password}\n\n"
        f"If you did not request this change, please contact Mangano IT service desk immediately on (07) 3151 9000"
    )

def send_sms(log, http_client, message, mobile_number, country_name, scheduled_time, suppress_log=False) -> bool:
    try:
        sender_id = "Mangano IT"
        username = get_secret_value(log, http_client, vault_name, "MIT-TransmitSMS-Username")
        password_secret = get_secret_value(log, http_client, vault_name, "MIT-TransmitSMS-Password")

        if not username or not password_secret:
            log.error("Missing TransmitSMS credentials from Key Vault")
            return False

        base_url = "https://api.transmitsms.com/send-sms.json"

        params = {
            "from": sender_id,
            "message": message,
            "to": mobile_number,
            "countrycode": country_name
        }

        if scheduled_time:
            params["send_at"] = scheduled_time

        auth_str = f"{username}:{password_secret}"
        encoded_auth = base64.b64encode(auth_str.encode("utf-8")).decode("utf-8")
        headers = {
            "Authorization": f"Basic {encoded_auth}",
            "Content-Type": "application/x-www-form-urlencoded"
        }

        params = {k: str(v) for k, v in params.items() if v}

        if suppress_log:
            safe_log = {**params, "message": "[REDACTED FOR SECURITY]"}
            log.info(f"Sending SMS with payload: {safe_log}")
        else:
            log.info(f"Sending SMS with payload: {params}")

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
                log.info(f"SMS sent successfully to [{mobile_number}]")
                log.result_data({"mobile_number": mobile_number})
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
            mobile_number = input.get_value("MobileNumber_1743126613581")
            country_name = input.get_value("Country_1743126620084")
            raw_schedule_date = input.get_value("ScheduleDate_1743129088388")
            raw_schedule_time = input.get_value("ScheduleTime_1743129095293")

            message_input = input.get_value("Message_1743126607975")
            first_name = input.get_value("FirstName_1746497930643")
            password = input.get_value("Password_1746500047781")

            ticket_number = input.get_value("TicketNumber_1746500954779")
            provided_token = input.get_value("AccessToken_1746500961880")

            sms_input = input.get_value("Notification_1749634789057")
        except:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        mobile_number = mobile_number.strip() if mobile_number else ""
        country_name = country_name.strip() if country_name else ""
        raw_schedule_date = raw_schedule_date.strip() if raw_schedule_date else ""
        raw_schedule_time = raw_schedule_time.strip() if raw_schedule_time else ""

        message_input = message_input.strip() if message_input else ""
        first_name = first_name.strip() if first_name else ""
        password = password.strip() if password else ""

        ticket_number = ticket_number.strip() if ticket_number else ""
        provided_token = provided_token.strip() if provided_token else ""

        sms_input = sms_input.strip() if sms_input else ""

        log.info(f"Received sms: [{sms_input}]")
        if not sms_input:
            record_result(log, ResultLevel.SUCCESS, "Bot completed with no sms selected. No action taken")
            return

        log.info(f"Received SMS number input = [{mobile_number}], Country code = [{country_name}]")
        if not mobile_number or not country_name:
            record_result(log, ResultLevel.WARNING, "Missing required input: Mobile Number or Country")
            return

        company_name = ""
        if ticket_number:
            company_identifier = get_company_identifier_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
            company_name = get_company_name_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
            log.info(f"Resolved company from ticket [{ticket_number}]: identifier=[{company_identifier}], name=[{company_name}]")

        if provided_token:
            access_token = provided_token
            log.info("Using provided access token")
            if not isinstance(access_token, str) or "." not in access_token:
                record_result(log, ResultLevel.WARNING, "Provided access token is malformed (missing dots)")
                return
        else:
            client_id = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientID")
            client_secret = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientSecret")
            azure_domain = get_secret_value(log, http_client, vault_name, "MIT-PrimaryDomain")

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

        scheduled_time = None
        if raw_schedule_date:
            scheduled_time = format_transmitsms_datetime(raw_schedule_date, raw_schedule_time, log)

        if sms_input == "Plain":
            message = codecs.decode(message, "unicode_escape")
            success = send_sms(log, http_client, message, mobile_number, country_name, scheduled_time)
            record_result(log, ResultLevel.SUCCESS if success else ResultLevel.WARNING,
                          f"SMS {'sent' if success else 'failed'} to {mobile_number}")

        elif sms_input == "User onboarding sms for user":
            message = build_user_onboarding_sms(first_name, company_name, password)
            message = codecs.decode(message, "unicode_escape")
            success = send_sms(log, http_client, message, mobile_number, country_name, scheduled_time, suppress_log=True)
            record_result(log, ResultLevel.SUCCESS if success else ResultLevel.WARNING,
                        f"Onboarding SMS {'sent' if success else 'failed'} to {mobile_number}")

        elif sms_input == "Password reset":
            message = build_password_reset_sms(password)
            message = codecs.decode(message, "unicode_escape")
            success = send_sms(log, http_client, message, mobile_number, country_name, scheduled_time, suppress_log=True)
            record_result(log, ResultLevel.SUCCESS if success else ResultLevel.WARNING,
                        f"Password reset SMS {'sent' if success else 'failed'} to {mobile_number}")

    except:
        record_result(log, ResultLevel.WARNING, "Process failed")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()