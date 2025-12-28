import sys
import random
import re
import os
import base64
import time
import codecs
import urllib.parse
import requests
from datetime import datetime, timedelta
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
sender_email = "support@PLACEHOLDER.com.au"

data_to_log = {}
bot_name = "UTIL - Send SMS"
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

    if response:
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
    token = get_access_token(log, http_client, tenant_id, client_id, client_secret, scope="https://graph.microsoft.com/.default")
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

def format_transmitsms_datetime(raw_date: str, raw_time: str, log) -> str | None:
    try:
        if not raw_date:
            log.warning("Missing schedule date - skipping scheduling")
            return None

        dt = None
        
        if raw_date.isdigit():
            try:
                epoch_value = int(raw_date)
                if epoch_value > 10000000000:
                    dt = datetime.fromtimestamp(epoch_value / 1000)
                    log.info(f"Successfully parsed epoch milliseconds [{raw_date}]")
                else:
                    dt = datetime.fromtimestamp(epoch_value)
                    log.info(f"Successfully parsed epoch seconds [{raw_date}]")
            except (ValueError, OSError) as e:
                log.warning(f"Invalid epoch timestamp: [{raw_date}] - {str(e)}")
                dt = None
        
        if dt is None:
            datetime_formats = [
                "%Y-%m-%dT%H:%M:%S",
                "%d/%m/%Y %I:%M %p",
                "%d/%m/%Y %H:%M",
                "%d-%m-%Y %I:%M %p",
                "%d-%m-%Y %H:%M",
                "%d/%m/%YT%H:%M:%S",
                "%d-%m-%YT%H:%M:%S",
                "%d/%m/%Y",
                "%d-%m-%Y", 
                "%Y-%m-%d"
            ]
            
            for datetime_format in datetime_formats:
                try:
                    dt = datetime.strptime(raw_date, datetime_format)
                    log.info(f"Successfully parsed datetime [{raw_date}] using format [{datetime_format}]")
                    break
                except ValueError:
                    continue
        
        if dt is None:
            log.warning(f"Invalid datetime format: [{raw_date}]. Expected formats: YYYY-MM-DDTHH:MM:SS, DD/MM/YYYY H:MM AM/PM, DD/MM/YYYY HH:MM, or epoch timestamp")
            return None

        if raw_time and raw_time.strip() and "T" not in raw_date:
            raw_time = raw_time.strip()
            
            if re.match(r"^\d{1,2}:\d{2}$", raw_time):
                try:
                    time_dt = datetime.strptime(raw_time, "%H:%M")
                    dt = dt.replace(hour=time_dt.hour, minute=time_dt.minute)
                    log.info(f"Successfully parsed time [{raw_time}] using HH:MM format")
                except ValueError:
                    log.warning(f"Invalid time format: [{raw_time}] - using default time (00:00)")
                    dt = dt.replace(hour=0, minute=0, second=0)
            
            elif re.match(r"^\d{4}$", raw_time):
                try:
                    hour = int(raw_time[:2])
                    minute = int(raw_time[2:])
                    if 0 <= hour <= 23 and 0 <= minute <= 59:
                        dt = dt.replace(hour=hour, minute=minute)
                        log.info(f"Successfully parsed time [{raw_time}] using HHMM format")
                    else:
                        log.warning(f"Invalid time values: hour={hour}, minute={minute} - using default time (00:00)")
                        dt = dt.replace(hour=0, minute=0, second=0)
                except ValueError:
                    log.warning(f"Invalid time format: [{raw_time}] - using default time (00:00)")
                    dt = dt.replace(hour=0, minute=0, second=0)
            else:
                log.warning(f"Unrecognized time format: [{raw_time}]. Expected formats: HH:MM or HHMM - using default time (00:00)")
                dt = dt.replace(hour=0, minute=0, second=0)
        else:
            dt = dt.replace(hour=0, minute=0, second=0)
            log.info("No time provided, using default time (00:00)")

        dt_utc = dt - timedelta(hours=10)
        formatted = dt_utc.strftime("%Y-%m-%d %H:%M:%S")
        log.info(f"Scheduled time converted to UTC: {formatted}")
        return formatted

    except Exception as e:
        log.warning(f"Failed to convert schedule date/time - skipping scheduling. Error: {str(e)}")
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
            ticket_number = input.get_value("TicketNumber_1746500954779")
            mobile_number = input.get_value("MobileNumber_1743126613581")
            country_name = input.get_value("Country_1758574417139")
            schedule_datetime = input.get_value("ScheduleDateandTime_1743129095293")
            message_input = input.get_value("Message_1743126607975")
            first_name = input.get_value("FirstName_1746497930643")
            password = input.get_value("Password_1746500047781")
            operation = input.get_value("Operation_1749634789057")
            graph_token = input.get_value("GraphToken_1746500961880")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        mobile_number = mobile_number.strip() if mobile_number else ""
        country_name = country_name.strip() if country_name else ""
        schedule_datetime = schedule_datetime.strip() if schedule_datetime else ""
        message_input = message_input.strip() if message_input else ""
        first_name = first_name.strip() if first_name else ""
        password = password.strip() if password else ""
        operation = operation.strip() if operation else ""
        graph_token = graph_token.strip() if graph_token else ""

        log.info(f"Ticket Number = [{ticket_number}]")
        log.info(f"Requested operation = [{operation}]")
        log.info(f"Received SMS number input = [{mobile_number}], Country code = [{country_name}]")

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number is required but missing")
            return
        if not operation:
            record_result(log, ResultLevel.SUCCESS, "Bot completed with no operation selected. No action taken")
            return
        if not mobile_number or not country_name:
            record_result(log, ResultLevel.WARNING, "Missing required input: Mobile Number or Country")
            return

        log.info(f"Retrieving company data for ticket [{ticket_number}]")
        company_identifier, company_name, company_id, company_type = get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
        if not company_identifier:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve company identifier from ticket [{ticket_number}]")
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

        scheduled_time = None
        if schedule_datetime:
            log.info(f"Using merged schedule datetime: [{schedule_datetime}]")
            scheduled_time = format_transmitsms_datetime(schedule_datetime, "", log)

        if operation == "Send Plain SMS":
            message = codecs.decode(message_input, "unicode_escape")
            success = send_sms(log, http_client, message, mobile_number, country_name, scheduled_time)
            record_result(log, ResultLevel.SUCCESS if success else ResultLevel.WARNING,
                          f"SMS {'sent' if success else 'failed'} to {mobile_number}")

        elif operation == "Send user onboarding SMS for user":
            message = build_user_onboarding_sms(first_name, company_name, password)
            message = codecs.decode(message, "unicode_escape")
            success = send_sms(log, http_client, message, mobile_number, country_name, scheduled_time, suppress_log=True)
            record_result(log, ResultLevel.SUCCESS if success else ResultLevel.WARNING,
                        f"Onboarding SMS {'sent' if success else 'failed'} to {mobile_number}")

        elif operation == "Send password reset SMS":
            message = build_password_reset_sms(password)
            message = codecs.decode(message, "unicode_escape")
            success = send_sms(log, http_client, message, mobile_number, country_name, scheduled_time, suppress_log=True)
            record_result(log, ResultLevel.SUCCESS if success else ResultLevel.WARNING,
                        f"Password reset SMS {'sent' if success else 'failed'} to {mobile_number}")

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
