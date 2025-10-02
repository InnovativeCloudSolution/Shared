import sys
import random
import os
import time
import urllib.parse
import requests
import re
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
vault_name = "mit-azu1-prod1-akv1"
sender_email = "support@manganoit.com.au"

data_to_log = {}
bot_name = "MIT-UTIL - Send ticket update notifications"
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
                    getattr(
                        http_client.third_party_integration(integration_name),
                        method
                    )(url=endpoint, json=data) if data else getattr(
                        http_client.third_party_integration(integration_name),
                        method
                    )(url=endpoint)
                )
            else:
                request_args = {"url": endpoint}
                if params:
                    request_args["params"] = params
                if headers:
                    request_args["headers"] = headers
                if data:
                    if (
                        headers and headers.get("Content-Type")
                        == "application/x-www-form-urlencoded"
                    ):
                        request_args["data"] = data
                    else:
                        request_args["json"] = data
                response = getattr(requests, method)(**request_args)

            if response.status_code in [200, 202, 204]:
                return response
            if response.status_code in [429, 503]:
                retry_after = response.headers.get("Retry-After")
                wait_time = (
                    int(retry_after) if retry_after else base_delay *
                    (2**attempt) + random.uniform(0, 3)
                )
                log.warning(
                    f"Rate limit exceeded. Retrying in {wait_time:.2f} seconds"
                )
                time.sleep(wait_time)
            elif response.status_code == 404:
                log.warning(f"Skipping non-existent resource [{endpoint}]")
                return None
            else:
                log.error(
                    f"API request failed Status: {response.status_code}, Response: {response.text}"
                )
                return response
        except Exception as e:
            log.exception(e, f"Exception during API call to {endpoint}")
            return None
    return None

def get_secret_value(log, http_client, vault_name, secret_name):
    log.info(f"Fetching secret [{secret_name}] from Key Vault [{vault_name}]")

    secret_url = (
        f"https://{vault_name}.vault.azure.net/secrets/{secret_name}?api-version=7.3"
    )
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

    log.error(
        f"Failed to retrieve secret [{secret_name}] Status code: {response.status_code if response else 'N/A'}"
    )
    return ""

def get_tenant_id_from_domain(log, http_client, azure_domain):
    try:
        config_url = (
            f"https://login.windows.net/{azure_domain}/.well-known/openid-configuration"
        )
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
    payload = urllib.parse.urlencode(
        {
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": scope,
        }
    )
    headers = {"Content-Type": "application/x-www-form-urlencoded"}

    response = execute_api_call(
        log,
        http_client,
        "post",
        token_url,
        data=payload,
        retries=3,
        headers=headers
    )

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

    log.error(
        f"Failed to retrieve access token Status code: {response.status_code if response else 'N/A'}"
    )
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

def get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number):
    log.info(f"Retrieving company details for ticket [{ticket_number}]")

    endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}"
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")

    if response:
        ticket_data = response.json()
        company = ticket_data.get("company", {})
        company_identifier = company.get("identifier", "")
        company_name = company.get("name", "")

        if company_identifier:
            log.info(f"Company identifier for ticket [{ticket_number}] is [{company_identifier}]")
        else:
            log.warning(f"Company identifier not found for ticket [{ticket_number}]")

        if company_name:
            log.info(f"Company name for ticket [{ticket_number}] is [{company_name}]")
        else:
            log.warning(f"Company name not found for ticket [{ticket_number}]")

        return company_identifier, company_name, company.get("id", ""), company.get("types", [])

    elif response:
        log.error(
            f"Failed to retrieve ticket [{ticket_number}] Status: {response.status_code}, Body: {response.text}"
        )
    return "", "", "", ""

def get_ticket_details(log, http_client, cwpsa_base_url, ticket_number):
    try:
        log.info(
            f"Retrieving full ticket details for ticket number [{ticket_number}]"
        )
        endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}"

        response = execute_api_call(
            log, http_client, "get", endpoint, integration_name="cw_psa"
        )
        if not response:
            log.error(
                f"Failed to retrieve ticket [{ticket_number}] - Status: {response.status_code if response else 'N/A'}"
            )
            return "", "", ""

        ticket = response.json()
        ticket_summary = ticket.get("summary", "")
        ticket_type = ticket.get("type", {}).get("name", "")
        priority_name = ticket.get("priority", {}).get("name", "")

        log.info(
            f"Ticket [{ticket_number}] Summary = [{ticket_summary}], Type = [{ticket_type}], Priority = [{priority_name}]"
        )
        return ticket_summary, ticket_type, priority_name

    except Exception as e:
        log.exception(
            e,
            f"Exception occurred while retrieving ticket details for [{ticket_number}]",
        )
        return "", "", ""

def is_business_hours(dt):
    if dt.weekday() >= 5:
        return False
    return 9 <= dt.hour < 17

def add_business_minutes(start_time, minutes_to_add):
    current = start_time
    remaining = minutes_to_add
    iterations = 0
    max_iterations = 1000

    while remaining > 0 and iterations < max_iterations:
        iterations += 1
        if is_business_hours(current):
            end_of_day = current.replace(
                hour=17, minute=0, second=0, microsecond=0
            )
            minutes_left_today = int(
                (end_of_day - current).total_seconds() // 60
            )

            if minutes_left_today >= remaining:
                current += timedelta(minutes=remaining)
                log.info(
                    f"Used [{remaining}] minutes at [{current}] Remaining: [0]"
                )
                remaining = 0
            else:
                current = end_of_day
                remaining -= minutes_left_today
                log.info(
                    f"Used [{minutes_left_today}] minutes at [{current}] Remaining: [{remaining}]"
                )
        else:
            if current.weekday() >= 5:
                days_until_monday = 7 - current.weekday()
                current = (current + timedelta(days=days_until_monday)
                           ).replace(hour=9, minute=0, second=0, microsecond=0)
            elif current.hour >= 17:
                current = (current + timedelta(days=1)
                           ).replace(hour=9, minute=0, second=0, microsecond=0)
            elif current.hour < 9:
                current = current.replace(
                    hour=9, minute=0, second=0, microsecond=0
                )
            else:
                current += timedelta(minutes=1)
    return current

def get_estimate_first_touch(log, http_client, cwpsa_base_url, priority_name, ticket_type):
    try:
        if priority_name.startswith("P1") or priority_name.startswith("P2"):
            log.info(
                f"Skipping SLA calculation for priority [{priority_name}] - handled directly in reviewed message"
            )
            return ""

        now_utc = datetime.now(timezone.utc)
        now_aest = now_utc.astimezone(timezone(timedelta(hours=10)))
        log.info(f"Python AEST Start: {now_aest.isoformat()}")

        end = now_utc.isoformat().split(".")[0] + "Z"
        start = (now_utc - timedelta(days=28)).isoformat().split(".")[0] + "Z"

        condition = (
            f"(source/name='Email' OR source/name='Desk Director') AND "
            f"company/identifier!='CQL' AND company/identifier!='CNS' AND company/identifier!='MIT' AND "
            f"resPlanMinutes!=0 AND dateResplan >= [{start}] AND dateResplan <= [{end}] AND "
            f"type/name='{ticket_type}' AND priority/name='{priority_name}'"
        )
        encoded_condition = urllib.parse.quote_plus(condition)
        endpoint = f"{cwpsa_base_url}/service/tickets?conditions={encoded_condition}&pageSize=1000"

        response = execute_api_call(
            log, http_client, "get", endpoint, integration_name="cw_psa"
        )
        if not response:
            log.warning(
                f"Ticket query failed or returned no data for {ticket_type} - {priority_name}"
            )
            return ""

        tickets = response.json()
        log.info(
            f"Total tickets retrieved: {len(tickets)} for condition: type/name='{ticket_type}', priority/name='{priority_name}'"
        )

        res_plan_values = sorted(
            [
                t.get("resPlanMinutes") for t in tickets
                if t.get("resPlanMinutes")
            ]
        )
        if not res_plan_values:
            log.warning("No valid resPlanMinutes found in ticket results")
            return ""

        index = int(0.75 * (len(res_plan_values) - 1))
        selected_minutes = res_plan_values[index]
        method = "75th percentile (28 days)"

        selected_minutes_ceiled = (
            int(selected_minutes) if selected_minutes == int(selected_minutes)
            else int(selected_minutes) + 1
        )
        log.info(
            f"Calculated {method} resPlanMinutes = {selected_minutes} (ceiled: {selected_minutes_ceiled}) for {ticket_type} - {priority_name}"
        )

        estimate_dt = add_business_minutes(now_aest, selected_minutes_ceiled)
        return estimate_dt.strftime("%d/%m/%Y")

    except Exception as e:
        log.exception(
            e,
            f"Failed to generate estimate_first_touch for {ticket_type} - {priority_name}",
        )
        return ""

def generate_ticket_reviewed(first_name, ticket_number, priority, ticket_summary, estimate_first_touch):
    priority_upper = priority.upper()
    if priority_upper.startswith("P1"):
        friendly_priority = "Priority 1"
    elif priority_upper.startswith("P2"):
        friendly_priority = "Priority 2"
    elif priority_upper.startswith("P3"):
        friendly_priority = "Priority 3"
    elif priority_upper.startswith("P4"):
        friendly_priority = "Priority 4"
    elif "VIP" in priority_upper:
        friendly_priority = "VIP"
    else:
        friendly_priority = priority

    if friendly_priority in ("Priority 1", "Priority 2"):
        eta_message = (
            "Due to the critical nature of this issue, we are currently organising an engineer "
            "to begin working on this - we will be in touch soon."
        )
    else:
        eta_message = (
            f"We're currently aiming to have one of the team working on this by <strong>{estimate_first_touch}</strong>.<br><br>"
            "This is our best estimate based on the information we have right now. If we need anything further from you "
            "along the way, we'll be sure to let you know, and we'll keep you up to date with any changes."
        )

    if friendly_priority == "VIP":
        priority_statement = (
            f"and we've classed it as a <strong>{friendly_priority}</strong> ticket."
        )
    else:
        priority_statement = f"and we've classed it as a <strong>{friendly_priority}</strong>, based on the impact and urgency of the ticket."

    subject = f"Your ticket #{ticket_number} has been reviewed: {ticket_summary}"

    html_body = f"""<!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>Reviewed Ticket</title>
    </head>
    <body style="margin:0; padding:20px; background-color:#d3d3d3; font-family:'Montserrat','Segoe UI','Roboto','Helvetica Neue','Calibri',sans-serif; color:#343a40; width:100% !important; text-align:center;">
    <center>
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%; max-width:960px; margin:0 auto; background-color:#ffffff; border-radius:12px; overflow:hidden; box-shadow:0 0 10px #000000; padding:0;">
        <tbody>
            <tr>
            <td align="center" style="padding: 20px">
                <a href="https://www.manganoit.com.au" target="_blank">
                <img
                    src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/Mangano%20IT%20(Logo-Horizontal).png"
                    alt="Mangano IT" width="215"
                    style="display: block; max-width: 215px; width: 100%; height: auto; border: 0; outline: none; text-decoration: none">
                </a>
            </td>
            </tr>
                <td>
                    <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%; background-color:#047cc2;">
                    <tr>
                        <td style="width:20%; padding:15px; color:#ffffff; font-weight:bold; text-align:center;">New</td>
                        <td style="width:20%; padding:20px; color:#ffffff; font-weight:bold; text-align:center; background-color:#9ccc50; border-radius:15px;">Reviewed</td>
                        <td style="width:20%; padding:15px; color:#ffffff; font-weight:bold; text-align:center;">Working Now</td>
                        <td style="width:20%; padding:15px; color:#ffffff; font-weight:bold; text-align:center;">Waiting</td>
                        <td style="width:20%; padding:15px; color:#ffffff; font-weight:bold; text-align:center;">Completed</td>
                    </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td style="padding:5px 15px 15px; text-align:center; background-color:#ffffff;">
                    <h1 style="color:#047cc2; font-size:24px;">Hi {first_name},</h1>
                    <h2 style="color:#333333; font-size:18px;">Your ticket has been reviewed: #{ticket_number}</h2>
                    <p style="color:#333333; font-size:16px;"><strong>Summary:</strong> {ticket_summary}</p>
                    <p style="color:#333333; font-size:16px;">Thanks for reaching out - your ticket (Ticket #{ticket_number}) has been reviewed by our Service Desk team {priority_statement}</p>
                    <p style="color:#333333; font-size:16px;">{eta_message}</p>
                    <p style="color:#333333; font-size:16px;">To view the progress of your ticket at any time or make changes visit our <a href="https://portal.manganoit.com.au" style="color:#9ccc50;">Mangano IT Service Desk Portal</a></p>
                    <p style="color:#333333; font-size:16px;"><strong>Need to make changes or things have escalated?</strong></p>
                    <p style="color:#333333; font-size:16px;">If the situation has become more urgent, or you think the priority should be reconsidered, just give us a call on <a href="tel:+61731519000" style="color:#9ccc50;">(07) 3151 9000</a>. We're here to help.</p>
                    <p style="color:#333333; font-size:16px;">Otherwise, you can <a href="https://manganoit.timezest.com/service-desk-analysts/phone-call-30/ticket/{ticket_number}" target="_blank" style="color:#9ccc50;">click here</a> to book a 30-minute call with a technician.</p>
                    <p style="color:#333333; font-size:16px;">Best regards,<br>The Mangano IT Team</p>
                </td>
            </tr>
            <tr>
                <td style="background-color:#047cc2; padding:10px; text-align:center; color:#ffffff; font-size:14px;">
                    <p style="margin:0 0 10px;">
                        Website: <a href="https://portal.manganoit.com.au" target="_blank" style="color:#ffffff;">https://portal.manganoit.com.au</a><br />
                        Email: <a href="mailto:support@manganoit.com.au" style="color:#ffffff;">support@manganoit.com.au</a><br />
                        Phone: <a href="tel:+61731519000" style="color:#ffffff;">(07) 3151 9000</a>
                    </p>
                    <div style="border-top:1px solid #ffffff; margin:10px 0;"></div>
                    <p style="margin:0;">&copy; Mangano IT 2025. All rights reserved.</p>
                </td>
            </tr>
        </tbody>
        </table>
    </center>
    </body>
    </html>
    """
    return subject, html_body

def generate_access_request_email(user_name, access_item, organisation_title="", organisation_department="", organisation_site_office="", organisation_manager_name=""):
    subject = f"{user_name} - {access_item}"

    details_html = f"""
      <tr>
        <td>
          <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%; padding:0 20px 20px;">
            <tr>
              <td style="padding:8px 0; font-size:16px; width:30%; font-weight:600;">Requestor:</td>
              <td style="padding:8px 0; font-size:16px; width:70%; text-align:right;">{user_name}</td>
            </tr>
    """

    if organisation_title:
        details_html += f"""
            <tr>
              <td style="padding:8px 0; font-size:16px; width:30%; font-weight:600;">Job Title:</td>
              <td style="padding:8px 0; font-size:16px; width:70%; text-align:right;">{organisation_title}</td>
            </tr>
        """
    if organisation_department:
        details_html += f"""
            <tr>
              <td style="padding:8px 0; font-size:16px; width:30%; font-weight:600;">Department:</td>
              <td style="padding:8px 0; font-size:16px; width:70%; text-align:right;">{organisation_department}</td>
            </tr>
        """
    if organisation_site_office:
        details_html += f"""
            <tr>
              <td style="padding:8px 0; font-size:16px; width:30%; font-weight:600;">Office location:</td>
              <td style="padding:8px 0; font-size:16px; width:70%; text-align:right;">{organisation_site_office}</td>
            </tr>
        """
    if organisation_manager_name:
        details_html += f"""
            <tr>
              <td style="padding:8px 0; font-size:16px; width:30%; font-weight:600;">Manager:</td>
              <td style="padding:8px 0; font-size:16px; width:70%; text-align:right;">{organisation_manager_name}</td>
            </tr>
        """

    details_html += "</table></td></tr>"

    html_body = f"""<!DOCTYPE html>
    <html lang="en">
    <head><meta charset="UTF-8" /><meta name="viewport" content="width=device-width, initial-scale=1.0" /><title>Internal IT Request</title></head>
    <body style="margin:0; padding:20px; background-color:#d3d3d3; font-family:'Montserrat','Segoe UI','Roboto','Helvetica Neue','Calibri',sans-serif; color:#343a40; width:100% !important; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%;">
    <div style="display:none; max-height:0; overflow:hidden;">Internal IT Request</div>
    <center>
    <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%; max-width:960px; margin:0 auto; background-color:#ffffff; border-radius:12px; overflow:hidden; box-shadow:0 0 10px #000000; box-sizing:border-box; padding:32px;">
      <tbody>
        <tr>
          <td style="text-align:center; padding-bottom:20px;">
            <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-ServiceDeskBanner.jpg" alt="Mangano IT - Service Update Banner" style="width:100%; height:auto; border-radius:12px; display:block;" />
          </td>
        </tr>
        <tr>
          <td style="text-align:center; padding:20px 0; font-size:28px; font-weight:700; color:#212529;">{user_name} - {access_item}</td>
        </tr>
        <tr>
          <td style="text-align:center; padding:0 20px 20px; font-size:16px; font-weight:600; color:#343a40;">
            The user listed above requires access to {access_item}.<br><br>Could you please arrange for the provisioning of
            this access at your earliest convenience? This may include installing software, granting system permissions,
            or configuring access to a SaaS platform.<br><br>If you encounter any issues or require further
            clarification, please don't hesitate to respond to this email.
          </td>
        </tr>
        {details_html}
        <tr>
          <td style="padding:20px; font-size:16px; line-height:1.5; color:#343a40;">
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%;">
              <tr>
                <td style="width:30%; font-weight:600; font-size:16px; color:#212529;">Any issues</td>
                <td style="width:70%; font-size:16px;">
                  Service Portal: <a href="https://portal.manganoit.com.au" style="color:#0066cc; text-decoration:none;">https://portal.manganoit.com.au</a><br />
                  Email: <a href="mailto:support@manganoit.com.au" style="color:#0066cc; text-decoration:none;">support@manganoit.com.au</a><br />
                  Phone: <a href="tel:+61731519000" style="color:#0066cc; text-decoration:none;">(07) 3151 9000</a>
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

def send_email(log, http_client, msgraph_base_url, access_token, sender_email, recipient_emails, subject, html_body):
    try:
        log.info(
            f"Preparing to send email from [{sender_email}] to [{recipient_emails}] with subject [{subject}]"
        )

        endpoint = f"{msgraph_base_url}/users/{sender_email}/sendMail"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

        if isinstance(recipient_emails, str):
            recipient_emails = [
                email.strip() for email in recipient_emails.split(",")
                if email.strip()
            ]

        to_recipients = [
            {
                "emailAddress": {
                    "address": email
                }
            } for email in recipient_emails
        ]

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

        response = execute_api_call(
            log,
            http_client,
            "post",
            endpoint,
            data=email_message,
            headers=headers
        )

        if response and response.status_code == 202:
            log.info(
                f"Successfully sent email to [{', '.join(recipient_emails)}] with subject [{subject}]"
            )
            return True
        else:
            log.error(
                f"Failed to send email to [{', '.join(recipient_emails)}] Status: {response.status_code if response else 'N/A'} Response: {response.text if response else 'N/A'}"
            )
            return False

    except Exception as e:
        log.exception(e, "Exception occurred while sending email via Graph")
        return False

def main():
    try:
        try:
            operation = input.get_value("Operation_1747968620394")
            recipient_emails = input.get_value("Recipients_1746937219795")
            first_name = input.get_value("FirstName_1746937221384")
            user_name = input.get_value("UserName_1749963752618")
            access = input.get_value("Access_1749963759640")
            ticket_number = input.get_value("TicketNumber_1746937195651")
            mit_graph_token = input.get_value("MITGraphToken_1746937205103")

            organisation_title = input.get_value("JobTitle_1750887478573")
            organisation_department = input.get_value("Department_1750887482045")
            organisation_site_office = input.get_value("Office_1750887485283")
            organisation_manager_name = input.get_value("ManagerName_1750887489661")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        log.info(f"Raw inputs received: ticket_number=[{ticket_number}], recipient_emails=[{recipient_emails}], first_name=[{first_name}], operation=[{operation}], user_name=[{user_name}], access=[{access}]")

        operation = operation.strip() if operation else ""
        recipient_emails = recipient_emails.strip() if recipient_emails else ""
        first_name = first_name.strip() if first_name else ""
        user_name = user_name.strip() if user_name else ""
        access = access.strip() if access else ""
        ticket_number = ticket_number.strip() if ticket_number else ""
        mit_graph_token = mit_graph_token.strip() if mit_graph_token else ""

        organisation_title = organisation_title.strip() if organisation_title else ""
        organisation_department = organisation_department.strip() if organisation_department else ""
        organisation_site_office = organisation_site_office.strip() if organisation_site_office else ""
        organisation_manager_name = organisation_manager_name.strip() if organisation_manager_name else ""

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
        log.info(f"Resolved company from ticket [{ticket_number}]: identifier=[{company_identifier}], name=[{company_name}]")

        if company_identifier == "MIT":
            if mit_graph_token:
                log.info("Using provided MIT MS Graph token")
                mit_graph_access_token = mit_graph_token
                mit_graph_tenant_id = ""
            else:
                mit_graph_tenant_id, mit_graph_access_token = get_graph_token_MIT(log, http_client, vault_name)
                if not mit_graph_access_token:
                    record_result(log, ResultLevel.WARNING, "Failed to obtain MIT MS Graph access token")
                    return
        else:
            mit_graph_tenant_id, mit_graph_access_token = get_graph_token_MIT(log, http_client, vault_name)
            if not mit_graph_access_token:
                record_result(log, ResultLevel.WARNING, "Failed to obtain MIT MS Graph access token")
                return

        if operation == "Send Reviewed Email":
            log.info("Executing operation: Reviewed")
            log.info(f"Retrieving ticket details for ticket number [{ticket_number}]")

            ticket_summary, ticket_type, ticket_priority = get_ticket_details(log, http_client, cwpsa_base_url, ticket_number)
            log.info(f"Ticket summary: [{ticket_summary}], Type: [{ticket_type}], Priority: [{ticket_priority}]")

            log.info(f"Calculating estimated first contact time for type [{ticket_type}], priority [{ticket_priority}]")
            estimate_first_touch = get_estimate_first_touch(log, http_client, cwpsa_base_url, ticket_priority, ticket_type)
            log.info(f"Estimated first contact by: {estimate_first_touch}")

            log.info("Preparing reviewed message content")
            log.info(f"Estimate used: [{estimate_first_touch}], Priority: [{ticket_priority}]")

            subject, html_body = generate_ticket_reviewed(first_name, ticket_number, ticket_priority, ticket_summary, estimate_first_touch)
            log.info("Reviewed message content generated successfully")

        elif operation == "Send Access Request Email":
            log.info("Executing operation: Access Request")

            if not user_name or not access or not recipient_emails:
                record_result(log, ResultLevel.WARNING, "Missing required input for access request email")
                return

            subject, html_body = generate_access_request_email(
                user_name,
                access,
                organisation_title=organisation_title,
                organisation_department=organisation_department,
                organisation_site_office=organisation_site_office,
                organisation_manager_name=organisation_manager_name,
            )
            log.info("Access request email content generated successfully")

        else:
            record_result(log, ResultLevel.INFO, f"No action defined for operation type [{operation}]. No action taken")

        log.info(f"Attempting to send {operation} email")
        subject_masked = re.sub(r"(\'|\")(eyJ[\w-]{5,}?\.[\w-]{5,}?\.([\w-]{5,})?)\1", r"\1***TOKEN-MASKED***\1", subject)
        subject_masked = re.sub(r"(\'|\")( [A-Za-z0-9]{8,}-(clientid|clientsecret|password))\1", r"\1***SECRET-MASKED***\1", subject_masked, flags=re.IGNORECASE)
        html_body_masked = re.sub(r"(\'|\")(eyJ[\w-]{5,}?\.[\w-]{5,}?\.([\w-]{5,})?)\1", r"\1***TOKEN-MASKED***\1", html_body)
        html_body_masked = re.sub(r"(\'|\")( [A-Za-z0-9]{8,}-(clientid|clientsecret|password))\1", r"\1***SECRET-MASKED***\1", html_body_masked, flags=re.IGNORECASE)
        if send_email(log, http_client, msgraph_base_url, mit_graph_access_token, sender_email, recipient_emails, subject_masked, html_body_masked):
            record_result(log, ResultLevel.SUCCESS, f"Successfully sent {operation} to [{recipient_emails}]")
        else:
            record_result(log, ResultLevel.WARNING, f"Failed to send {operation} to [{recipient_emails}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
