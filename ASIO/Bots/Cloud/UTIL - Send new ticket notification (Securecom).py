import sys
import random
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

cwpsa_base_url = "https://aus.myconnectwise.net/v4_6_release/apis/3.0"
msgraph_base_url = "https://graph.microsoft.com/v1.0"
msgraph_base_url_beta = "https://graph.microsoft.com/beta"
vault_name = "asio-test"

data_to_log = {}
bot_name = "UTIL - Send new ticket notification (Securecom)"
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
    try:
        log.info(f"Retrieving secret [{secret_name}] from vault [{vault_name}]")
        endpoint = f"https://{vault_name}.vault.azure.net/secrets/{secret_name}?api-version=7.3"
        response = execute_api_call(log, http_client, "get", endpoint, integration_name="azure_keyvault")
        if response:
            secret_value = response.json().get("value", "")
            log.info(f"Successfully retrieved secret [{secret_name}]")
            return secret_value
        log.error(f"Failed to retrieve secret [{secret_name}] from vault [{vault_name}]")
        return ""
    except Exception as e:
        log.exception(e, f"Exception while retrieving secret [{secret_name}]")
        return ""

def get_tenant_id_from_domain(log, http_client, azure_domain):
    try:
        config_url = f"https://login.windows.net/{azure_domain}/.well-known/openid-configuration"
        log.info(f"Fetching OpenID configuration for domain [{azure_domain}]")
        response = execute_api_call(log, http_client, "get", config_url)
        if response:
            token_endpoint = response.json().get("token_endpoint", "")
            tenant_id = token_endpoint.split("/")[3] if token_endpoint else ""
            if tenant_id:
                log.info(f"Successfully extracted tenant ID for domain [{azure_domain}]")
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
        if not isinstance(access_token, str) or "." not in access_token:
            log.error(f"[{log_prefix}] Access token is invalid or malformed")
            return ""
        log.info(f"[{log_prefix}] Successfully retrieved access token")
        return access_token
    return ""

def get_graph_token_email(log, http_client, vault_name):
    sender_email = get_secret_value(log, http_client, vault_name, "Sender-Email")
    client_id = get_secret_value(log, http_client, vault_name, "App-ClientID")
    client_secret = get_secret_value(log, http_client, vault_name, "App-ClientSecret")
    if not all([sender_email, client_id, client_secret]):
        log.error("Failed to retrieve required secrets for email sending")
        return "", "", ""
    azure_domain = sender_email.split("@")[1] if "@" in sender_email else ""
    if not azure_domain:
        log.error("Failed to extract Azure domain from sender email")
        return "", "", ""
    tenant_id = get_tenant_id_from_domain(log, http_client, azure_domain)
    if not tenant_id:
        log.error("Failed to resolve tenant ID for email sending")
        return "", "", ""
    token = get_access_token(log, http_client, tenant_id, client_id, client_secret, scope="https://graph.microsoft.com/.default", log_prefix="EmailGraph")
    if not isinstance(token, str) or "." not in token:
        log.error("Email Graph access token is malformed")
        return "", "", ""
    log.info("Successfully obtained MS Graph token for email sending")
    return tenant_id, token, sender_email

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

def get_ticket_details(log, http_client, cwpsa_base_url, ticket_number):
    try:
        log.info(f"Retrieving ticket details for ticket [{ticket_number}]")
        endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}"
        response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
        if response:
            ticket = response.json()
            ticket_summary = ticket.get("summary", "")
            ticket_detail = ticket.get("initialDescription", "")
            if not ticket_detail:
                ticket_detail = "No additional details provided."
            log.info(f"Ticket Summary: [{ticket_summary}]")
            return ticket_summary, ticket_detail
        return "", ""
    except Exception as e:
        log.exception(e, f"Exception while retrieving ticket details for [{ticket_number}]")
        return "", ""

def generate_notification_new(ticket_number, ticket_summary, ticket_detail):
    html_body = f"""<html>
    <head>
    <title>New ticket submitted</title>
    </head>
    <body>
    <table align="center" border="0" cellpadding="10px 15px" cellspacing="1" style="width: 100%; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 150%">
    <tbody>
        <tr>
        <td>
            <table align="left" border="0" cellpadding="10px 15px" cellspacing="1" style="max-width: 850px; width: 100%; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 150%">
            <tbody>
                <tr>
                <td style="text-align: left; vertical-align: middle; width: 201px"><img alt="" border="0" height="53" hspace="0" src="https://mcusercontent.com/98076c486b38ae32852eedf52/images/abfdcd10-08bc-789a-f10f-e6ae14ed7a1c.png" style="width: 201px; height: 53px; margin: 0; border: 0 solid #000000" vspace="0" width="201" /></td>
                <td style="text-align: left; vertical-align: middle; width: 201px;">&nbsp;</td>
                <td style="text-align: center; vertical-align: middle; width: 201px;"><span style="color:#1e3f76; font-size: 20px;"><strong>Ticket #{ticket_number}</strong></span></td>
                </tr>
                <tr>
                <td colspan="3" style="text-align: center; height: 80px; vertical-align: middle; background-color: #3ca900; padding: 15px 0px 0px;"><span style="color: #ffffff; font-size: 26px"><strong>New ticket submitted</strong>
    {ticket_summary}</span></td>
                </tr>
                <tr>
                <td colspan="3" style="text-align: center; padding-top: 20px; padding-bottom: 20px"><span style="color: #1e3f76; font-size: 16px"><strong>Your request has been received</strong></span></td>
                </tr>
                <tr>
                <td colspan="3">
                <p><span>We understand it is important and will get to work on it. If we have more information we can resolve your issue quicker!
    If you have not already supplied this, please respond with the following:</span></p>
                <ol>
                    <li><strong>Error message</strong>. Detailed error message, name of application you were using and the impacted user (Important if you are logging on behalf of someone else).</li>
                    <li><strong>More details</strong>. Details of what you were doing at the time. (What did you expect to happen and what instead did happen? What does &quot;resolved&quot; look like to you?).</li>
                    <li><strong>Screenshots</strong>. Screenshots of your screen and/or error messages are often very helpful in resolving issues.</li>
                </ol>
                <p>Thank you in advance for your assistance and cooperation. Our goal is to provide you <strong>excellent support</strong>. We would appreciate any feedback on how we are doing.</p>
                </td>
                </tr>
            </tbody>
            </table>
        </td>
        </tr>
        <tr>
        <td>
            <table align="left" border="0" cellpadding="10px 15px" cellspacing="1" style="width: 100%; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 150%">
            <tbody>
                <tr>
                <td colspan="3" style="text-align: left"><span>{ticket_detail}</span></td>
                </tr>
            </tbody>
            </table>
        </td>
        </tr>
        <tr>
        <td>
            <table align="left" border="0" cellpadding="10px 15px" cellspacing="1" style="max-width: 850px; width: 100%; font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 150%">
            <tbody>
                <tr>
                <td colspan="3"><span style="color: #1e3f76; font-size: 11px"><i>We live and breathe our values:</i></span></td>
                </tr>
                <tr>
                <td colspan="3" style="text-align: center;"><img alt="" border="0" height="26" hspace="0" src="https://www.securecom.co.nz/wp-content/uploads/2022/05/SC_Values_700x26.png" style="width: 700px; height: 26px; margin: 0; border: 0 solid #000000" vspace="0" width="700" /></td>
                </tr>
                <tr>
                <td colspan="3" style="white-space: nowrap; text-align: center; height: 60px; vertical-align: middle; background-color: #1e3f76;"><span style="color: #ffffff; font-size: 10px">Securecom Ltd. | Level 2 , Building A, 600 Great South Road, Ellerslie, Auckland 1051, New Zealand | 2022 Securecom Ltd. All Rights Reserved.</span></td>
                </tr>
                <tr>
                <td colspan="3"><span style="color: #777777; font-style: italic; font-size: 9px; line-height: normal">This communication may contain information which is confidential and/or privileged to Securecom Limited. Unless you are the intended recipient you may not disclose, copy or use it; please notify the sender immediately and delete it and any copies. You should protect your system from viruses etc; we accept no responsibility for damage that may be caused by them.</span></td>
                </tr>        
            </tbody>
            </table>
        </td>
        </tr>      
    </tbody>
    </table>
    </body>
    </html>"""
    return html_body

def send_email(log, http_client, msgraph_base_url, access_token, sender_email, recipient_emails, subject, html_body):
    try:
        log.info(f"Preparing to send email from [{sender_email}] to [{recipient_emails}] with subject [{subject}]")
        endpoint = f"{msgraph_base_url}/users/{sender_email}/sendMail"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }
        if isinstance(recipient_emails, str):
            recipient_emails = [email.strip() for email in recipient_emails.split(",") if email.strip()]
        to_recipients = [{"emailAddress": {"address": email}} for email in recipient_emails]
        email_message = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "HTML",
                    "content": html_body
                },
                "toRecipients": to_recipients,
            },
            "saveToSentItems": "true",
        }
        response = execute_api_call(log, http_client, "post", endpoint, data=email_message, headers=headers)
        if response and response.status_code == 202:
            log.info(f"Successfully sent email to [{', '.join(recipient_emails)}] with subject [{subject}]")
            return True
        else:
            log.error(f"Failed to send email to [{', '.join(recipient_emails)}] Status: {response.status_code if response else 'N/A'} Response: {response.text if response else 'N/A'}")
            return False
    except Exception as e:
        log.exception(e, "Exception occurred while sending email via Graph")
        return False

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_xxxxxxxxxxxxx")
            recipient_emails = input.get_value("Recipients_xxxxxxxxxxxxx")
            email_graph_token = input.get_value("EmailGraphToken_xxxxxxxxxxxxx")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        recipient_emails = recipient_emails.strip() if recipient_emails else ""
        email_graph_token = email_graph_token.strip() if email_graph_token else ""

        log.info(f"Ticket Number = [{ticket_number}]")
        log.info(f"Recipients = [{recipient_emails}]")

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number is required but missing")
            return
        if not recipient_emails:
            record_result(log, ResultLevel.WARNING, "Recipient email address is required but missing")
            return

        log.info(f"Retrieving company data for ticket [{ticket_number}]")
        company_identifier, company_name, company_id, company_type = get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
        if not company_identifier:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve company identifier from ticket [{ticket_number}]")
            return
        data_to_log["Company"] = company_identifier

        log.info("Acquiring MS Graph token for email sending")
        if email_graph_token:
            email_graph_access_token = email_graph_token
            email_graph_tenant_id = ""
            sender_email = get_secret_value(log, http_client, vault_name, "Sender-Email")
        else:
            email_graph_tenant_id, email_graph_access_token, sender_email = get_graph_token_email(log, http_client, vault_name)
        
        if not email_graph_access_token:
            record_result(log, ResultLevel.WARNING, "Failed to obtain email MS Graph access token")
            return
        if not sender_email:
            record_result(log, ResultLevel.WARNING, "Failed to retrieve sender email address")
            return

        log.info(f"Retrieving ticket details for ticket [{ticket_number}]")
        ticket_summary, ticket_detail = get_ticket_details(log, http_client, cwpsa_base_url, ticket_number)
        if not ticket_summary:
            ticket_summary = f"Support Request #{ticket_number}"

        html_body = generate_notification_new(ticket_number, ticket_summary, ticket_detail)
        subject = f"Ticket #{ticket_number} - New"

        if send_email(log, http_client, msgraph_base_url, email_graph_access_token, sender_email, recipient_emails, subject, html_body):
            record_result(log, ResultLevel.SUCCESS, f"Successfully sent New notification to [{recipient_emails}]")
        else:
            record_result(log, ResultLevel.WARNING, f"Failed to send New notification to [{recipient_emails}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()

