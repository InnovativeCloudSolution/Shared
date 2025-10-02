import sys
import random
import os
import time
import urllib.parse
import requests
from datetime import datetime
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

cwpsa_base_url = "https://au.myconnectwise.net/v4_6_release/apis/3.0"
msgraph_base_url = "https://graph.microsoft.com/v1.0"
msgraph_base_url_beta = "https://graph.microsoft.com/beta"
vault_name = "mit-azu1-prod1-akv1"
sender_email = "support@manganoit.com.au"

data_to_log = {}
bot_name = "MIT-UTIL - Send user offboarding notifications"
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
        log.error(f"Failed to extract tenant ID from domain [{azure_domain}]")
        return ""
    except Exception as e:
        log.exception(e, "Exception while extracting tenant ID from domain")
        return ""

def get_graph_token_MIT(log, http_client, vault_name):
    log.info("Fetching MS Graph token for MIT domain")
    
    client_id = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientID")
    client_secret = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientSecret")
    azure_domain = get_secret_value(log, http_client, vault_name, "MIT-PrimaryDomain")
    
    if not all([client_id, client_secret, azure_domain]):
        log.error("Failed to retrieve required secrets for MIT domain")
        return "", ""
    
    tenant_id = get_tenant_id_from_domain(log, http_client, azure_domain)
    if not tenant_id:
        log.error(f"Failed to resolve tenant ID for domain [{azure_domain}]")
        return "", ""
    
    token = get_access_token(log, http_client, tenant_id, client_id, client_secret, scope="https://graph.microsoft.com/.default")
    if not isinstance(token, str) or "." not in token:
        log.error("MS Graph access token is malformed for MIT domain")
        return "", ""
    
    log.info("Successfully obtained MS Graph token for MIT domain")
    return tenant_id, token

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
    log.error(f"[{log_prefix}] Failed to retrieve access token Status code: {response.status_code if response else 'N/A'}")
    return ""

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

def send_email(log, http_client, msgraph_base_url, access_token, sender_email, recipient_emails, subject, html_body):
    try:
        log.info(f"Preparing to send email from [{sender_email}] to [{recipient_emails}] with subject [{subject}]")

        endpoint = f"{msgraph_base_url}/users/{sender_email}/sendMail"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
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
                "toRecipients": to_recipients
            },
            "saveToSentItems": "true"
        }

        response = execute_api_call(log, http_client, "post", endpoint, data=email_message, headers=headers)

        if response and response.status_code == 202:
            log.info(f"Successfully sent email to [{', '.join(recipient_emails)}] with subject [{subject}]")
            return True
        else:
            
            return False

    except Exception as e:
        log.exception(e, "Exception occurred while sending email via Graph")
        return False

def format_date(date_raw):
    """
    Standardize date format to DD/MM/YYYY hh:mm AM/PM.
    Handles input formats: DD/MM/YYYY, YYYY-MM-DD, YYYY-MM-DDT00:00
    """
    if not date_raw:
        return ""
    
    try:
        if "/" in str(date_raw):
            try:
                parsed = datetime.strptime(str(date_raw), "%d/%m/%Y")
                return parsed.strftime("%d/%m/%Y %I:%M %p")
            except ValueError:
                pass
        
        if "-" in str(date_raw) and "T" not in str(date_raw):
            try:
                parsed = datetime.strptime(str(date_raw), "%Y-%m-%d")
                return parsed.strftime("%d/%m/%Y %I:%M %p")
            except ValueError:
                pass
        
        if "T" in str(date_raw):
            try:
                parsed = datetime.fromisoformat(str(date_raw))
                return parsed.strftime("%d/%m/%Y %I:%M %p")
            except ValueError:
                pass
        
        try:
            parsed = datetime.strptime(str(date_raw), "%Y-%m-%d")
            return parsed.strftime("%d/%m/%Y %I:%M %p")
        except ValueError:
            pass
            
    except Exception:
        pass
    
    return str(date_raw)
    
def generate_user_offboarding_asset_collection_email(ticket_number, user_name, user_enddate, devices=""):
    subject = f"{user_name} - User Offboarding Asset Collection - Ticket #{ticket_number}"

    def row(label, value):
        if value:
            return f"""
            <tr>
                <td style="padding:8px 0; font-size:16px; width:30%; font-weight:600;">{label}</td>
                <td style="padding:8px 0; font-size:16px; width:70%; text-align:right;">{value}</td>
            </tr>"""
        return ""
    
    device_html = ""
    if devices:
        items = [d.strip() for d in devices.split(",") if d.strip()]
        if items:
            device_html = "<br>".join(items)
            
    html_body = f"""
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>{user_name} - User Offboarding</title>
      </head>
      <body style="margin:0; padding:20px; background-color:#d3d3d3; font-family:'Montserrat','Segoe UI','Roboto','Helvetica Neue','Calibri',sans-serif; color:#343a40; width:100% !important; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%;">
        <div style="display:none; max-height:0; overflow:hidden;">
          {user_name} - User Offboarding
        </div>
        <center>
          <table role="presentation" cellpadding="0" cellspacing="0" border="0"
                 style="width:100%; max-width:960px; margin:0 auto; background-color:#ffffff; border-radius:12px; overflow:hidden; box-shadow:0 0 10px #000000; box-sizing:border-box; padding:32px;">
            <tbody>
              <tr>
                <td style="text-align:center; padding-bottom:20px;">
                  <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-OffboardingBanner.jpg"
                       alt="Mangano IT - Service Update Banner"
                       style="width:100%; height:auto; border-radius:12px; display:block;" />
                </td>
              </tr>
              <tr>
                <td style="text-align:center; padding:20px 0; font-size:28px; font-weight:700; color:#212529;">
                  {user_name} - User Offboarding Asset Collection - Ticket #{ticket_number}
                </td>
              </tr>
              <tr>
                <td style="text-align:center; padding:0 20px 20px; font-size:16px; font-weight:600; color:#343a40;">
                  You have been assigned the task of collecting hardware from a user who has been recently offboarded.
                </td>
              </tr>
              <tr>
                <td>
                  <table role="presentation" cellpadding="0" cellspacing="0" border="0"
                         style="width:100%; padding:0 20px 20px;">
                    {row("User Name:", user_name)}
                    {row("End date and time:", user_enddate)}
                    {row("Devices:", device_html)}
                  </table>
                </td>
              </tr>
              <tr>
                <td style="padding:0 20px 20px; font-size:16px; line-height:1.5; color:#343a40;">
                  Please retrieve all hardware from the user and notify us so this can be prepared for a future user.<br><br>
                  If you need any assistance with the above please call us urgently.
                </td>
              </tr>
              <tr>
                <td style="padding:20px; font-size:16px; line-height:1.5; color:#343a40;">
                  <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%;">
                    <tr>
                      <td style="width:30%; font-weight:600; font-size:16px; color:#212529;">Any issues</td>
                      <td style="width:70%; font-size:16px;">
                        Service Portal:
                        <a href="https://portal.manganoit.com.au"
                           style="color:#0066cc; text-decoration:none;">https://portal.manganoit.com.au</a><br />
                        Email:
                        <a href="mailto:support@manganoit.com.au"
                           style="color:#0066cc; text-decoration:none;">support@manganoit.com.au</a><br />
                        Phone:
                        <a href="tel:+61731519000" style="color:#0066cc; text-decoration:none;">(07) 3151 9000</a>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </tbody>
          </table>
        </center>
      </body>
      </html>
      """
    
    return subject, html_body
    
def generate_user_offboarding_email_complete(ticket_number, user_name, user_enddate, exchange_mailboxdelegate="", exchange_forwardto="", exchange_ooodelegate="", teams_delegateowner="", devices="", asset_delegate="", mobile_number_tag=""):
    subject = f"{user_name} - User Offboarding Complete - Ticket #{ticket_number}"

    def row(label, value):
        if value:
            return f"""
              <tr>
                <td style="padding:8px 0; font-size:16px; width:30%; font-weight:600;">{label}</td>
                <td style="padding:8px 0; font-size:16px; width:70%; text-align:right;">{value}</td>
              </tr>"""
        return ""

    device_html = ""
    if devices:
        items = [d.strip() for d in devices.split(",") if d.strip()]
        if items:
            device_html = "<br>".join(items)

    html_body = f"""
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>{user_name} - User Offboarding Complete</title>
      </head>
      <body style="margin:0; padding:20px; background-color:#d3d3d3; font-family:'Montserrat','Segoe UI','Roboto','Helvetica Neue','Calibri',sans-serif; color:#343a40; width:100% !important;">
        <div style="display:none; max-height:0; overflow:hidden;">
          {user_name} - User Offboarding Complete
        </div>
        <center>
          <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%; max-width:960px; margin:0 auto; background-color:#ffffff; border-radius:12px; overflow:hidden; box-shadow:0 0 10px #000000; box-sizing:border-box; padding:32px;">
            <tbody>
              <tr>
                <td style="text-align:center; padding-bottom:20px;">
                  <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-OffboardingBanner.jpg" alt="Mangano IT - Service Update Banner" style="width:100%; height:auto; border-radius:12px; display:block;" />
                </td>
              </tr>
              <tr>
                <td style="text-align:center; padding:20px 0; font-size:28px; font-weight:700; color:#212529;">
                  {user_name} - User Offboarding Complete - Ticket #{ticket_number}
                </td>
              </tr>
              <tr>
                <td style="text-align:center; padding:0 20px 20px; font-size:16px; font-weight:600; color:#343a40;">
                  The offboarding has been <b>completed</b> - details below:
                </td>
              </tr>
              <tr>
                <td>
                  <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%; padding:0 20px 20px;">
                    {row("User Name:", user_name)}
                    {row("End date and time:", user_enddate)}
                    {row("Email access given to:", exchange_mailboxdelegate)}
                    {row("Emails forwarded to:", exchange_forwardto)}
                    {row("Auto response directs to:", exchange_ooodelegate)}
                    {row("Teams ownership given to:", teams_delegateowner)}
                    {row("Assets currently assigned to this user:", device_html)}
                    {row("Nominee to collect assets:", asset_delegate)}
                    {row("Mobile numbers to be cancelled:", mobile_number_tag)}
                  </table>
                </td>
              </tr>
              <tr>
                <td style="padding:0 20px 20px; font-size:16px; line-height:1.5; color:#343a40;">
                  Please retrieve all hardware from the user and notify us so this can be prepared for a future user.<br>
                </td>
              </tr>
              <tr>
                <td style="padding:20px; font-size:16px; line-height:1.5; color:#343a40;">
                  <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%;">
                    <tr>
                      <td style="width:30%; font-weight:600; font-size:16px; color:#212529;">Any issues</td>
                      <td style="width:70%; font-size:16px;">
                        Service Portal:
                        <a href="https://portal.manganoit.com.au" style="color:#0066cc; text-decoration:none;">https://portal.manganoit.com.au</a><br />
                        Email:
                        <a href="mailto:support@manganoit.com.au" style="color:#0066cc; text-decoration:none;">support@manganoit.com.au</a><br />
                        Phone:
                        <a href="tel:+61731519000" style="color:#0066cc; text-decoration:none;">(07) 3151 9000</a>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </tbody>
          </table>
        </center>
      </body>
      </html>
      """

    return subject, html_body

def generate_user_offboarding_email_scheduled(ticket_number, user_name, user_enddate):
    html_body = f"""
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>{user_name} - User Offboarding</title>
      </head>
      <body style="margin:0; padding:20px; background-color:#d3d3d3; font-family:'Montserrat','Segoe UI','Roboto','Helvetica Neue','Calibri',sans-serif; color:#343a40; width:100% !important; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%;">
        <div style="display:none; max-height:0; overflow:hidden;">
          {user_name} - User Offboarding
        </div>
        <center>
          <table role="presentation" cellpadding="0" cellspacing="0" border="0"
                 style="width:100%; max-width:960px; margin:0 auto; background-color:#ffffff; border-radius:12px; overflow:hidden; box-shadow:0 0 10px #000000; box-sizing:border-box; padding:32px;">
            <tbody>
              <tr>
                <td style="text-align:center; padding-bottom:20px;">
                  <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-OffboardingBanner.jpg"
                       alt="Mangano IT - Service Update Banner"
                       style="width:100%; height:auto; border-radius:12px; display:block;" />
                </td>
              </tr>
              <tr>
                <td style="text-align:center; padding:20px 0; font-size:28px; font-weight:700; color:#212529;">
                  {user_name} - User Offboarding Details - Ticket #{ticket_number}
                </td>
              </tr>
              <tr>
                <td style="text-align:center; padding:0 20px 20px; font-size:16px; font-weight:600; color:#343a40;">
                  The offboarding has been <b>scheduled</b> - details below:
                </td>
              </tr>
              <tr>
                <td>
                  <table role="presentation" cellpadding="0" cellspacing="0" border="0"
                         style="width:100%; padding:0 20px 20px;">
                    <tr>
                      <td style="padding:8px 0; font-size:16px; width:30%; font-weight:600;">User Name:</td>
                      <td style="padding:8px 0; font-size:16px; width:70%; text-align:right;">{user_name}</td>
                    </tr>
                    <tr>
                      <td style="padding:8px 0; font-size:16px; width:30%; font-weight:600;">End date and time:</td>
                      <td style="padding:8px 0; font-size:16px; width:70%; text-align:right;">{user_enddate}</td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr>
                <td style="padding:0 20px 20px; font-size:16px; line-height:1.5; color:#343a40;">
                  <b>At the above scheduled time, the user will instantly be logged out and their account will no longer be valid.</b><br>
                  Please retrieve all hardware from the user and notify us so this can be prepared for a future user.<br><br>
                  If you need to make changes to the above please call us urgently.
                </td>
              </tr>
              <tr>
                <td style="padding:20px; font-size:16px; line-height:1.5; color:#343a40;">
                  <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%;">
                    <tr>
                      <td style="width:30%; font-weight:600; font-size:16px; color:#212529;">Any issues</td>
                      <td style="width:70%; font-size:16px;">
                        Service Portal:
                        <a href="https://portal.manganoit.com.au"
                           style="color:#0066cc; text-decoration:none;">https://portal.manganoit.com.au</a><br />
                        Email:
                        <a href="mailto:support@manganoit.com.au"
                           style="color:#0066cc; text-decoration:none;">support@manganoit.com.au</a><br />
                        Phone:
                        <a href="tel:+61731519000" style="color:#0066cc; text-decoration:none;">(07) 3151 9000</a>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </tbody>
          </table>
        </center>
      </body>
      </html>
      """

    subject = f"{user_name} - User Offboarding Scheduled - Ticket #{ticket_number}"
    return subject, html_body

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1749604627542")
            operation = input.get_value("Operation_1749604643029")
            recipient_emails = input.get_value("Recipients_1749604673149")
            user_name = input.get_value("UserName_1749604684081")
            end_date = input.get_value("EndDate_1749604945017")
            mit_graph_token = input.get_value("MITGraphToken_1758661738690")
            graph_token = input.get_value("GraphToken_1749604630419")
            
            mailbox_delegate = input.get_value("MailboxDelegate_1749604993235")
            forward_to = input.get_value("ForwardTo_1749605043314")
            ooo_delegate = input.get_value("OOODelegate_1749605041677")
            teams_delegate = input.get_value("TeamsOwnerDelegate_1749605048282")
            devices = input.get_value("Devices_1749605045347")
            asset_delegate = input.get_value("AssetDelegate_1749605050997")
            mobile_number = input.get_value("MobileNumber_1749605055653")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        operation = operation.strip() if operation else ""
        recipient_emails = recipient_emails.strip() if recipient_emails else ""
        user_name = user_name.strip() if user_name else ""
        end_date = format_date(end_date.strip()) if end_date else ""
        mit_graph_token = mit_graph_token.strip() if mit_graph_token else ""
        graph_token = graph_token.strip() if graph_token else ""
        
        mailbox_delegate = mailbox_delegate.strip() if mailbox_delegate else ""
        forward_to = forward_to.strip() if forward_to else ""
        ooo_delegate = ooo_delegate.strip() if ooo_delegate else ""
        teams_delegate = teams_delegate.strip() if teams_delegate else ""
        devices = devices.strip() if devices else ""
        asset_delegate = asset_delegate.strip() if asset_delegate else ""
        mobile_number = mobile_number.strip() if mobile_number else ""

        log.info(f"Ticket Number = [{ticket_number}]")
        log.info(f"Requested operation = [{operation}]")
        log.info(f"User Name = [{user_name}]")

        if not operation:
            record_result(log, ResultLevel.INFO, "Bot completed with no operation selected. No action taken")
            return

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number is required but missing")
            return

        log.info(f"Retrieving company data for ticket [{ticket_number}]")
        company_identifier, company_name, company_id, company_types = get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
        if not company_identifier:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve company identifier from ticket [{ticket_number}]")
            return
        
        data_to_log["Company"] = company_identifier
        log.info(f"Resolved company from ticket [{ticket_number}]: identifier=[{company_identifier}]")

        if mit_graph_token:
            log.info("Using provided MIT MS Graph token")
            mit_graph_access_token = mit_graph_token
            mit_graph_tenant_id = ""
        else:
            mit_graph_tenant_id, mit_graph_access_token = get_graph_token_MIT(log, http_client, vault_name)
            if not mit_graph_access_token:
                record_result(log, ResultLevel.WARNING, "Failed to obtain MIT MS Graph access token")
                return

        if graph_token:
            log.info("Using provided MS Graph token")
            graph_access_token = graph_token
            graph_tenant_id = ""
        else:
            graph_tenant_id, graph_access_token = get_graph_token(log, http_client, vault_name, "MIT")
            if not graph_access_token:
                record_result(log, ResultLevel.WARNING, "Failed to obtain MS Graph access token")
                return

        azure_domain = get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")
        if not azure_domain:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve Azure domain for [{company_identifier}]")
            return

        if operation == "Complete":
            log.info("Executing operation: Complete")
            subject, html_body = generate_user_offboarding_email_complete(
                ticket_number, user_name, end_date,
                exchange_mailboxdelegate=mailbox_delegate,
                exchange_forwardto=forward_to,
                exchange_ooodelegate=ooo_delegate,
                teams_delegateowner=teams_delegate,
                devices=devices,
                asset_delegate=asset_delegate,
                mobile_number_tag=mobile_number
            )
        
        elif operation == "Scheduled":
            log.info("Executing operation: Scheduled")
            subject, html_body = generate_user_offboarding_email_scheduled(
                ticket_number, user_name, end_date
            )
        
        elif operation == "Asset Collection":
            log.info("Executing operation: Asset Collection")
            subject, html_body = generate_user_offboarding_asset_collection_email(
                ticket_number, user_name, end_date, devices
            )
        
        else:
            record_result(log, ResultLevel.WARNING, f"Unknown operation [{operation}]")
            return

        log.info(f"Attempting to send user offboarding [{operation}] notification email")
        if send_email(log, http_client, msgraph_base_url, mit_graph_access_token, sender_email, recipient_emails, subject, html_body):
            record_result(log, ResultLevel.SUCCESS, f"Successfully sent user offboarding [{operation}] notification to [{recipient_emails}]")
        else:
            record_result(log, ResultLevel.WARNING, f"Failed to send user offboarding [{operation}] notification to [{recipient_emails}]")
            return

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
