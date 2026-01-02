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

cwpsa_base_url = "https://aus.myconnectwise.net"
cwpsa_base_url_path = "/v4_6_release/apis/3.0"
msgraph_base_url_base = "https://graph.microsoft.com"
msgraph_base_url_path = "/v1.0"
msgraph_base_url_beta_base = "https://graph.microsoft.com"
msgraph_base_url_beta_path = "/beta"
vault_name = "dbit-azu1-prod1-akv1"
sender_email = "help@dropbear-it.com.au"

data_to_log = {}
bot_name = "UTIL - Send ticket update notifications"
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
    return ""

def get_graph_token_email(log, http_client, vault_name, email_domain_secret="EmailDomain"):
    log.info("Fetching MS Graph token for email sending")
    client_id = get_secret_value(log, http_client, vault_name, "PartnerApp-ClientID")
    client_secret = get_secret_value(log, http_client, vault_name, "PartnerApp-ClientSecret")
    azure_domain = get_secret_value(log, http_client, vault_name, email_domain_secret)
    if not all([client_id, client_secret, azure_domain]):
        log.error("Failed to retrieve required secrets for email Graph token")
        return "", ""
    tenant_id = get_tenant_id_from_domain(log, http_client, azure_domain)
    if not tenant_id:
        log.error(f"Failed to resolve tenant ID for domain [{azure_domain}]")
        return "", ""
    token = get_access_token(log, http_client, tenant_id, client_id, client_secret, scope="https://graph.microsoft.com/.default", log_prefix="EmailGraph")
    if not isinstance(token, str) or "." not in token:
        log.error("Email Graph access token is malformed")
        return "", ""
    log.info("Successfully obtained MS Graph token for email sending")
    return tenant_id, token

def get_company_data_from_ticket(log, http_client, cwpsa_base_url, cwpsa_base_url_path, ticket_number):
    log.info(f"Retrieving company details for ticket [{ticket_number}]")
    ticket_endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets/{ticket_number}"
    ticket_response = execute_api_call(log, http_client, "get", ticket_endpoint, integration_name="cw_psa")
    if ticket_response:
        ticket_data = ticket_response.json()
        company = ticket_data.get("company", {})
        company_id = company["id"]
        company_identifier = company["identifier"]
        company_name = company["name"]
        log.info(f"Company ID: [{company_id}], Identifier: [{company_identifier}], Name: [{company_name}]")
        company_endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/company/companies/{company_id}"
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
        endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets/{ticket_number}"
        response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
        if response:
            ticket = response.json()
            contact = ticket.get("contact", {})
            contact_first_name = contact.get("name", "").split()[0] if contact.get("name") else ""
            contact_email = contact.get("communicationItems", [{}])[0].get("value", "") if contact.get("communicationItems") else ""
            ticket_summary = ticket.get("summary", "")
            member = ticket.get("member", {})
            member_first_name = member.get("name", "").split()[0] if member.get("name") else ""
            member_last_name = " ".join(member.get("name", "").split()[1:]) if member.get("name") else ""
            log.info(f"Contact: [{contact_first_name}], Email: [{contact_email}], Summary: [{ticket_summary}], Member: [{member_first_name} {member_last_name}]")
            return contact_first_name, contact_email, ticket_summary, member_first_name, member_last_name
        return "", "", "", "", ""
    except Exception as e:
        log.exception(e, f"Exception while retrieving ticket details for [{ticket_number}]")
        return "", "", "", "", ""

def generate_html_base(content_html):
    return f"""
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8" />
            <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Ticket Update - DropBear IT</title>
        </head>
    <body style="margin:0; padding:20px; background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); font-family:'Inter','Segoe UI','Roboto','Helvetica Neue',sans-serif; color:#2d3436; width:100% !important; text-align:center; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%;">
        <center>
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%; max-width:680px; margin:0 auto; background-color:#ffffff; border-radius:16px; overflow:hidden; box-shadow:0 20px 60px rgba(0,0,0,0.3); box-sizing:border-box; padding:0;">
            <tbody>
                <tr>
            <td align="center" style="padding:0; position:relative;">
                <a href="https://dropbear-it.com.au/" target="_blank" style="display:block; position:relative;">
                <img src="https://cloud.katana.nexigen.digital/katana/fTgTpj2FrgOsSo5U8d3s8sSDb8ASdePnZUc8i4ak.webp" alt="DropBear IT" width="680" style="display:block; width:100%; max-width:680px; height:auto; border:0; outline:none; text-decoration:none; border-radius:16px 16px 0 0;">
                    </a>
                </td>
                </tr>
            {content_html}
            <tr>
            <td style="background:linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); padding:30px 35px; text-align:center; color:#ffffff; font-size:14px;">
                <div style="margin-bottom:20px;">
                <p style="margin:0 0 12px; font-size:15px; font-weight:600; color:#ffffff;">Get in Touch</p>
                <p style="margin:0 0 8px;"><span style="color:#b8bfce;">?? Website:</span> <a href="https://dropbear-it.com.au/" target="_blank" style="color:#667eea; text-decoration:none; font-weight:500; margin-left:5px;">dropbear-it.com.au</a></p>
                <p style="margin:0 0 8px;"><span style="color:#b8bfce;">?? Email:</span> <a href="mailto:help@dropbear-it.com.au" style="color:#667eea; text-decoration:none; font-weight:500; margin-left:5px;">help@dropbear-it.com.au</a></p>
                <p style="margin:0;"><span style="color:#b8bfce;">?? Phone:</span> <a href="tel:+611800573165" style="color:#667eea; text-decoration:none; font-weight:500; margin-left:5px;">1800 573 165</a></p>
                </div>
                <div style="border-top:1px solid rgba(255,255,255,0.2); margin:20px 0; padding-top:20px;">
                <p style="margin:0; color:#b8bfce; font-size:13px;">&copy; 2025 DropBear IT. All rights reserved.</p>
                </div>
                    </td>
                </tr>
            </tbody>
            </table>
        </center>
        </body>
    </html>
    """

def generate_notification_new(contact_first_name, ticket_number, ticket_summary):
    content = f"""
    <tr>
    <td style="padding:40px 35px; text-align:left; background-color:#ffffff;">
        <h1 style="color:#1a1a2e; font-size:28px; font-weight:700; margin:0 0 20px; line-height:1.3;">G'day {contact_first_name}! ??</h1>
        <div style="background:linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding:20px; border-radius:12px; margin-bottom:25px;">
        <h2 style="color:#ffffff; font-size:20px; font-weight:600; margin:0 0 8px;">Your Ticket Has Been Created</h2>
        <p style="color:#f0f0f0; font-size:16px; margin:0; font-weight:500;">Ticket #{ticket_number}</p>
        </div>
        <div style="background-color:#f8f9fa; padding:20px; border-radius:10px; border-left:4px solid #667eea; margin-bottom:25px;">
        <p style="color:#2d3436; font-size:15px; margin:0 0 5px; font-weight:600;">Summary:</p>
        <p style="color:#636e72; font-size:15px; margin:0; line-height:1.5;">{ticket_summary}</p>
        </div>
        <p style="color:#2d3436; font-size:16px; line-height:1.6; margin-bottom:20px;">Thanks for reaching out! We're on it. Our team will review your ticket and assign a priority based on the urgency and impact, then get back to you shortly with an estimated timeframe.</p>
        <div style="background-color:#d1ecf1; border-left:4px solid #17a2b8; padding:18px; border-radius:8px; margin:25px 0;">
        <p style="color:#0c5460; font-size:15px; margin:0 0 10px; font-weight:600;">How We Prioritise Tickets</p>
        <p style="color:#0c5460; font-size:14px; margin:0; line-height:1.5;">We assess tickets based on impact and urgency to ensure critical issues get attention first. You'll receive an update once your ticket has been reviewed.</p>
        </div>
        <div style="background-color:#fff3cd; border-left:4px solid #ffc107; padding:18px; border-radius:8px; margin:25px 0;">
        <p style="color:#856404; font-size:15px; margin:0 0 10px; font-weight:600;">Things Changed or Escalated?</p>
        <p style="color:#856404; font-size:14px; margin:0; line-height:1.5;">If the situation becomes urgent or business critical, give us a call immediately. We're here to help.</p>
        </div>
        <div style="text-align:center; margin:30px 0;">
        <a href="tel:+611800573165" style="display:inline-block; background:linear-gradient(135deg, #667eea 0%, #764ba2 100%); color:#ffffff; padding:14px 32px; border-radius:8px; text-decoration:none; font-size:16px; font-weight:600; box-shadow:0 4px 15px rgba(102,126,234,0.4);">?? Call Us: 1800 573 165</a>
        </div>
        <p style="color:#2d3436; font-size:16px; line-height:1.6; margin:30px 0 10px;">Cheers,<br><strong style="color:#667eea;">DropBear IT Team</strong></p>
    </td>
    </tr>
    """
    return generate_html_base(content)

def generate_notification_reviewed(contact_first_name, ticket_number, ticket_summary):
    content = f"""
    <tr>
    <td style="padding:40px 35px; text-align:left; background-color:#ffffff;">
        <h1 style="color:#1a1a2e; font-size:28px; font-weight:700; margin:0 0 20px; line-height:1.3;">G'day {contact_first_name}! ??</h1>
        <div style="background:linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding:20px; border-radius:12px; margin-bottom:25px;">
        <h2 style="color:#ffffff; font-size:20px; font-weight:600; margin:0 0 8px;">Your Ticket Has Been Reviewed</h2>
        <p style="color:#f0f0f0; font-size:16px; margin:0; font-weight:500;">Ticket #{ticket_number}</p>
        </div>
        <div style="background-color:#f8f9fa; padding:20px; border-radius:10px; border-left:4px solid #667eea; margin-bottom:25px;">
        <p style="color:#2d3436; font-size:15px; margin:0 0 5px; font-weight:600;">Summary:</p>
        <p style="color:#636e72; font-size:15px; margin:0; line-height:1.5;">{ticket_summary}</p>
        </div>
        <p style="color:#2d3436; font-size:16px; line-height:1.6; margin-bottom:20px;">Thanks for your patience! Your ticket has been reviewed by our Service Desk team and prioritized. We'll keep you updated with any changes and reach out if we need anything else from you.</p>
        <div style="background-color:#fff3cd; border-left:4px solid #ffc107; padding:18px; border-radius:8px; margin:25px 0;">
        <p style="color:#856404; font-size:15px; margin:0 0 10px; font-weight:600;">Priority Changed or More Urgent?</p>
        <p style="color:#856404; font-size:14px; margin:0; line-height:1.5;">If the situation has escalated or you believe the priority needs reconsidering, give us a call. We're here to help.</p>
        </div>
        <div style="text-align:center; margin:30px 0;">
        <a href="tel:+611800573165" style="display:inline-block; background:linear-gradient(135deg, #667eea 0%, #764ba2 100%); color:#ffffff; padding:14px 32px; border-radius:8px; text-decoration:none; font-size:16px; font-weight:600; box-shadow:0 4px 15px rgba(102,126,234,0.4);">?? Call Us: 1800 573 165</a>
        </div>
        <p style="color:#2d3436; font-size:16px; line-height:1.6; margin:30px 0 10px;">Cheers,<br><strong style="color:#667eea;">DropBear IT Team</strong></p>
    </td>
    </tr>
    """
    return generate_html_base(content)

def generate_notification_working(contact_first_name, ticket_number, ticket_summary, member_first_name, member_last_name):
    content = f"""
    <tr>
    <td style="padding:40px 35px; text-align:left; background-color:#ffffff;">
        <h1 style="color:#1a1a2e; font-size:28px; font-weight:700; margin:0 0 20px; line-height:1.3;">G'day {contact_first_name}! ??</h1>
        <div style="background:linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding:20px; border-radius:12px; margin-bottom:25px;">
        <h2 style="color:#ffffff; font-size:20px; font-weight:600; margin:0 0 8px;">Your Ticket is Being Worked On</h2>
        <p style="color:#f0f0f0; font-size:16px; margin:0; font-weight:500;">Ticket #{ticket_number}</p>
        </div>
        <div style="background-color:#f8f9fa; padding:20px; border-radius:10px; border-left:4px solid #667eea; margin-bottom:25px;">
        <p style="color:#2d3436; font-size:15px; margin:0 0 5px; font-weight:600;">Summary:</p>
        <p style="color:#636e72; font-size:15px; margin:0; line-height:1.5;">{ticket_summary}</p>
        </div>
        <p style="color:#2d3436; font-size:16px; line-height:1.6; margin-bottom:20px;">Great news! Your ticket has been assigned to <strong style="color:#667eea;">{member_first_name} {member_last_name}</strong> and they're working on it now. {member_first_name} will reach out if they need any additional information from you.</p>
        <div style="background-color:#fff3cd; border-left:4px solid #ffc107; padding:18px; border-radius:8px; margin:25px 0;">
        <p style="color:#856404; font-size:15px; margin:0 0 10px; font-weight:600;">Need to Update Your Ticket?</p>
        <p style="color:#856404; font-size:14px; margin:0; line-height:1.5;">If the situation has become more urgent or you need to provide additional details, give us a call. We're here to help!</p>
        </div>
        <div style="text-align:center; margin:30px 0;">
        <a href="tel:+611800573165" style="display:inline-block; background:linear-gradient(135deg, #667eea 0%, #764ba2 100%); color:#ffffff; padding:14px 32px; border-radius:8px; text-decoration:none; font-size:16px; font-weight:600; box-shadow:0 4px 15px rgba(102,126,234,0.4);">?? Call Us: 1800 573 165</a>
        </div>
        <p style="color:#2d3436; font-size:16px; line-height:1.6; margin:30px 0 10px;">Cheers,<br><strong style="color:#667eea;">DropBear IT Team</strong></p>
    </td>
    </tr>
    """
    return generate_html_base(content)

def generate_notification_waiting(contact_first_name, ticket_number, ticket_summary):
    content = f"""
    <tr>
    <td style="padding:40px 35px; text-align:left; background-color:#ffffff;">
        <h1 style="color:#1a1a2e; font-size:28px; font-weight:700; margin:0 0 20px; line-height:1.3;">G'day {contact_first_name}! ??</h1>
        <div style="background:linear-gradient(135deg, #ffc107 0%, #ff9800 100%); padding:20px; border-radius:12px; margin-bottom:25px;">
        <h2 style="color:#ffffff; font-size:20px; font-weight:600; margin:0 0 8px;">We're Waiting to Hear From You</h2>
        <p style="color:#f0f0f0; font-size:16px; margin:0; font-weight:500;">Ticket #{ticket_number}</p>
        </div>
        <div style="background-color:#f8f9fa; padding:20px; border-radius:10px; border-left:4px solid #ffc107; margin-bottom:25px;">
        <p style="color:#2d3436; font-size:15px; margin:0 0 5px; font-weight:600;">Summary:</p>
        <p style="color:#636e72; font-size:15px; margin:0; line-height:1.5;">{ticket_summary}</p>
        </div>
        <p style="color:#2d3436; font-size:16px; line-height:1.6; margin-bottom:20px;">Just a friendly reminder that we haven't heard back from you yet regarding this ticket.</p>
        <div style="background-color:#d1ecf1; border-left:4px solid #17a2b8; padding:18px; border-radius:8px; margin:25px 0;">
        <p style="color:#0c5460; font-size:15px; margin:0 0 10px; font-weight:600;">Is Everything Working Now?</p>
        <p style="color:#0c5460; font-size:14px; margin:0; line-height:1.5;">If the issue is resolved and everything's working as expected, that's great! Just let us know and we'll close off the ticket.</p>
        </div>
        <div style="background-color:#fff3cd; border-left:4px solid #ffc107; padding:18px; border-radius:8px; margin:25px 0;">
        <p style="color:#856404; font-size:15px; margin:0 0 10px; font-weight:600;">Still Having Issues?</p>
        <p style="color:#856404; font-size:14px; margin:0; line-height:1.5;">If the problem persists, please reply to this email or give us a call and we'll continue investigating for you.</p>
        </div>
        <div style="text-align:center; margin:30px 0;">
        <a href="tel:+611800573165" style="display:inline-block; background:linear-gradient(135deg, #667eea 0%, #764ba2 100%); color:#ffffff; padding:14px 32px; border-radius:8px; text-decoration:none; font-size:16px; font-weight:600; box-shadow:0 4px 15px rgba(102,126,234,0.4);">?? Call Us: 1800 573 165</a>
        </div>
        <div style="background-color:#f8d7da; border-left:4px solid #dc3545; padding:18px; border-radius:8px; margin:25px 0;">
        <p style="color:#721c24; font-size:14px; margin:0; line-height:1.5;"><strong>Important:</strong> If we don't hear from you within 48 hours, we'll close this ticket and send you a completion email. You can always respond to reopen it if needed.</p>
        </div>
        <p style="color:#2d3436; font-size:16px; line-height:1.6; margin:30px 0 10px;">Cheers,<br><strong style="color:#667eea;">DropBear IT Team</strong></p>
    </td>
    </tr>
    """
    return generate_html_base(content)

def generate_notification_closed(contact_first_name, ticket_number, ticket_summary):
    content = f"""
    <tr>
    <td style="padding:40px 35px; text-align:left; background-color:#ffffff;">
        <h1 style="color:#1a1a2e; font-size:28px; font-weight:700; margin:0 0 20px; line-height:1.3;">G'day {contact_first_name}! ??</h1>
        <div style="background:linear-gradient(135deg, #28a745 0%, #20c997 100%); padding:20px; border-radius:12px; margin-bottom:25px;">
        <h2 style="color:#ffffff; font-size:20px; font-weight:600; margin:0 0 8px;">Your Ticket Has Been Completed</h2>
        <p style="color:#f0f0f0; font-size:16px; margin:0; font-weight:500;">Ticket #{ticket_number}</p>
        </div>
        <div style="background-color:#f8f9fa; padding:20px; border-radius:10px; border-left:4px solid #28a745; margin-bottom:25px;">
        <p style="color:#2d3436; font-size:15px; margin:0 0 5px; font-weight:600;">Summary:</p>
        <p style="color:#636e72; font-size:15px; margin:0; line-height:1.5;">{ticket_summary}</p>
        </div>
        <p style="color:#2d3436; font-size:16px; line-height:1.6; margin-bottom:25px;">We've wrapped up your ticket and hope everything is working smoothly for you now! If there's anything else you need help with, just reply to this email or reach out to us.</p>
        <div style="background-color:#d4edda; border-left:4px solid #28a745; padding:20px; border-radius:8px; margin:30px 0; text-align:center;">
        <h3 style="color:#155724; font-size:18px; margin:0 0 15px; font-weight:600;">How Did We Do?</h3>
        <p style="color:#155724; font-size:14px; margin:0 0 20px; line-height:1.5;">We'd love to hear about your experience. Your feedback helps us improve!</p>
        <table border="0" cellpadding="0" cellspacing="0" role="presentation" style="margin:0 auto; border-collapse:collapse;">
        <tbody>
            <tr>
                <td style="padding:10px 15px; text-align:center;">
                <a href="https://feedback.smileback.io/r/7/Sr2ajccj_vCkD-4T3KHVUQ/{ticket_number}/1/" target="_blank">
                    <img alt="Positive" height="55" src="https://feedback.smileback.io/v/7/Sr2ajccj_vCkD-4T3KHVUQ/{ticket_number}/1/" style="margin:0 auto; display:block;" width="55" />
                </a>
                <p style="margin:8px 0 0; color:#155724; font-size:13px; font-weight:500;">Positive</p>
            </td>
                <td style="padding:10px 15px; text-align:center;">
                <a href="https://feedback.smileback.io/r/7/Sr2ajccj_vCkD-4T3KHVUQ/{ticket_number}/0/" target="_blank">
                    <img alt="Neutral" height="55" src="https://feedback.smileback.io/v/7/Sr2ajccj_vCkD-4T3KHVUQ/{ticket_number}/0/" style="margin:0 auto; display:block;" width="55" />
                </a>
                <p style="margin:8px 0 0; color:#155724; font-size:13px; font-weight:500;">Neutral</p>
            </td>
                <td style="padding:10px 15px; text-align:center;">
                <a href="https://feedback.smileback.io/r/7/Sr2ajccj_vCkD-4T3KHVUQ/{ticket_number}/-1/" target="_blank">
                    <img alt="Negative" height="55" src="https://feedback.smileback.io/v/7/Sr2ajccj_vCkD-4T3KHVUQ/{ticket_number}/-1/" style="margin:0 auto; display:block;" width="55" />
                </a>
                <p style="margin:8px 0 0; color:#155724; font-size:13px; font-weight:500;">Negative</p>
            </td>
            </tr>
        </tbody>
        </table>
        <p style="color:#155724; font-size:13px; margin:15px 0 0; line-height:1.4;">Be honest - whether positive or negative, we appreciate your feedback.</p>
        </div>
        <div style="background-color:#d1ecf1; border-left:4px solid #17a2b8; padding:18px; border-radius:8px; margin:25px 0;">
        <p style="color:#0c5460; font-size:14px; margin:0; line-height:1.5;">Still need help with this issue? Just reply to this email and we'll reopen your ticket right away.</p>
        </div>
        <p style="color:#2d3436; font-size:16px; line-height:1.6; margin:30px 0 10px;">Cheers,<br><strong style="color:#667eea;">DropBear IT Team</strong></p>
    </td>
    </tr>
    """
    return generate_html_base(content)

def send_email(log, http_client, msgraph_base_url, access_token, sender_email, recipient_emails, subject, html_body):
    try:
        log.info(f"Preparing to send email from [{sender_email}] to [{recipient_emails}] with subject [{subject}]")
        endpoint = f"{msgraph_base_url_base}{msgraph_base_url_path}/users/{sender_email}/sendMail"
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
            operation = input.get_value("Operation_xxxxxxxxxxxxx")
            email_graph_token = input.get_value("EmailGraphToken_xxxxxxxxxxxxx")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        operation = operation.strip() if operation else ""
        email_graph_token = email_graph_token.strip() if email_graph_token else ""

        log.info(f"Ticket Number = [{ticket_number}]")
        log.info(f"Requested operation = [{operation}]")

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number is required but missing")
            return
        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation value is missing or invalid")
            return

        log.info(f"Retrieving company data for ticket [{ticket_number}]")
        company_identifier, company_name, company_id, company_type = get_company_data_from_ticket(log, http_client, cwpsa_base_url, cwpsa_base_url_path, ticket_number)
        if not company_identifier:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve company identifier from ticket [{ticket_number}]")
            return
        data_to_log["Company"] = company_identifier

        if email_graph_token:
            log.info("Using provided email MS Graph token")
            email_graph_access_token = email_graph_token
            email_graph_tenant_id = ""
        else:
            email_graph_tenant_id, email_graph_access_token = get_graph_token_email(log, http_client, vault_name)
            if not email_graph_access_token:
                record_result(log, ResultLevel.WARNING, "Failed to obtain email MS Graph access token")
                return

        log.info(f"Retrieving ticket details for ticket [{ticket_number}]")
        contact_first_name, contact_email, ticket_summary, member_first_name, member_last_name = get_ticket_details(log, http_client, cwpsa_base_url, ticket_number)
        if not contact_email:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve contact email for ticket [{ticket_number}]")
            return
        if not ticket_summary:
            ticket_summary = f"Support Request #{ticket_number}"

        if operation == "New":
            html_body = generate_notification_new(contact_first_name, ticket_number, ticket_summary)
            subject = f"Ticket #{ticket_number} - New"
        elif operation == "Reviewed":
            html_body = generate_notification_reviewed(contact_first_name, ticket_number, ticket_summary)
            subject = f"Ticket #{ticket_number} - Reviewed"
        elif operation == "Working":
            html_body = generate_notification_working(contact_first_name, ticket_number, ticket_summary, member_first_name, member_last_name)
            subject = f"Ticket #{ticket_number} - Being Worked On"
        elif operation == "Waiting":
            html_body = generate_notification_waiting(contact_first_name, ticket_number, ticket_summary)
            subject = f"Ticket #{ticket_number} - Waiting for Response"
        elif operation == "Closed":
            html_body = generate_notification_closed(contact_first_name, ticket_number, ticket_summary)
            subject = f"Ticket #{ticket_number} - Completed"
        else:
            record_result(log, ResultLevel.WARNING, f"Unknown operation [{operation}]")
            return

        subject_masked = re.sub(r"('|\")(eyJ[a-zA-Z0-9_-]{5,}?\.[a-zA-Z0-9_-]{5,}?\.([a-zA-Z0-9_-]{5,})?)\1", r"\1***TOKEN-MASKED***\1", subject)
        subject_masked = re.sub(r"('|\")([a-zA-Z0-9]{8,}-(clientid|clientsecret|password))\1", r"\1***SECRET-MASKED***\1", subject_masked, flags=re.IGNORECASE)
        html_body_masked = re.sub(r"('|\")(eyJ[a-zA-Z0-9_-]{5,}?\.[a-zA-Z0-9_-]{5,}?\.([a-zA-Z0-9_-]{5,})?)\1", r"\1***TOKEN-MASKED***\1", html_body)
        html_body_masked = re.sub(r"('|\")([a-zA-Z0-9]{8,}-(clientid|clientsecret|password))\1", r"\1***SECRET-MASKED***\1", html_body_masked, flags=re.IGNORECASE)

        if send_email(log, http_client, msgraph_base_url, email_graph_access_token, sender_email, contact_email, subject_masked, html_body_masked):
            record_result(log, ResultLevel.SUCCESS, f"Successfully sent {operation} notification to [{contact_email}]")
        else:
            record_result(log, ResultLevel.WARNING, f"Failed to send {operation} notification to [{contact_email}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
