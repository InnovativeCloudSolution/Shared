import sys
import traceback
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
sender_email = "support@PLACEHOLDER.com.au"

data_to_log = {}
bot_name = "UTIL - Send user onboarding notifications"
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

    log.error(f"Failed to retrieve access token Status code: {response.status_code if response else 'N/A'}")
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
        log.error(f"Failed to retrieve ticket [{ticket_number}] Status: {ticket_response.status_code}, Body: {ticket_response.text}")
    return "", "", 0, []

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
    token = get_access_token(log, http_client, tenant_id, client_id, client_secret, scope="https://graph.microsoft.com/.default", log_prefix="Graph")
    if not isinstance(token, str) or "." not in token:
        log.error("MS Graph access token is malformed (missing dots)")
        return "", ""
    return tenant_id, token

def format_date(date_raw):
    """
    Standardize date format to DD/MM/YYYY.
    Handles input formats: DD/MM/YYYY, YYYY-MM-DD, YYYY-MM-DDT00:00
    """
    if not date_raw:
        return ""
    
    try:
        if "/" in str(date_raw):
            try:
                parsed = datetime.strptime(str(date_raw), "%d/%m/%Y")
                return parsed.strftime("%d/%m/%Y")
            except ValueError:
                pass
        
        if "-" in str(date_raw) and "T" not in str(date_raw):
            try:
                parsed = datetime.strptime(str(date_raw), "%Y-%m-%d")
                return parsed.strftime("%d/%m/%Y")
            except ValueError:
                pass
        
        if "T" in str(date_raw):
            try:
                parsed = datetime.fromisoformat(str(date_raw))
                return parsed.strftime("%d/%m/%Y")
            except ValueError:
                pass
        
        try:
            parsed = datetime.strptime(str(date_raw), "%Y-%m-%d")
            return parsed.strftime("%d/%m/%Y")
        except ValueError:
                pass
            
    except Exception:
        pass
    
    return str(date_raw)

def generate_user_onboarding_email(first_name, company_name, user_email):
    subject = f"Welcome {first_name}! Let's get your IT set up!"

    html_body = f"""
    <!DOCTYPE html>
    <html lang="en">

    <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Welcome {first_name}</title>
    </head>

    <body style="margin:0; padding:20px; background-color:#d3d3d3; font-family:'Montserrat', 'Segoe UI', 'Roboto', 'Helvetica Neue', 'Calibri', sans-serif; color:#343a40; width:100% !important; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%;">

    <div style="display:none; max-height:0; overflow:hidden;">
        Welcome {first_name}
    </div>

    <center>
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%; max-width:960px; margin:0 auto; background-color:#ffffff; border-radius:12px; overflow:hidden; box-shadow:0 0 10px #000000; box-sizing:border-box; padding:32px;">
        <tbody>

            <tr>
            <td style="text-align:center; padding-bottom:20px;">
                <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-OnboardingBanner.jpg"
                alt="Mangano IT - Welcome Banner"
                style="width:100%; height:auto; border-radius:12px; display:block;" />
            </td>
            </tr>

            <tr>
            <td style="text-align:center; padding:20px 0; font-size:28px; font-weight:700; color:#212529;">
                Welcome {first_name}
            </td>
            </tr>

            <tr>
            <td style="padding:0 20px 20px; font-size:16px; line-height:1.5; color:#343a40;">
                Congratulations on your new role at {company_name} - we're excited to help you get up and running smoothly!<br><br>
                Your account is set up and you should have received a text message with your temporary password.<br><br>
                Here are your login details to get started:
            </td>
            </tr>

            <tr>
            <td style="padding:0 20px 20px; font-size:16px; line-height:1.5; color:#343a40;">
                <b>Login:</b> {user_email}<br>
                <b>Password:</b> (This is in the text message you have received)
            </td>
            </tr>

            <tr>
            <td style="padding:0 20px 20px; font-size:16px; line-height:1.5; color:#343a40;">
                Once you have logged in, please make sure you can access:
                <ul style="margin-top:10px; padding-left:20px;">
                <li style="margin-bottom:8px;">The company network</li>
                <li style="margin-bottom:8px;">Your Outlook mailbox</li>
                <li style="margin-bottom:8px;">SharePoint</li>
                <li style="margin-bottom:8px;">Microsoft Teams</li>
                </ul>
                If anything isn't working quite right, or you didn't receive your password, please give our Service Desk team a call - we're here to help.<br><br>
                We're looking forward to supporting you in your new role - welcome aboard!<br><br>
                Warm regards,<br>
                The Mangano IT Team
            </td>
            </tr>

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

def generate_manager_onboarding_email(user_name, start_date, user_email, groups, devices, ticket_number):
    subject = f"{user_name} - New Starter Onboarding Details - Ticket #{ticket_number}"

    formatted_groups = "<br>".join(groups) if isinstance(groups, list) else "<br>".join([g.strip() for g in groups.split(",") if g.strip()])

    formatted_devices = "<br>".join(devices) if isinstance(devices, list) else "<br>".join([g.strip() for g in devices.split(",") if g.strip()])

    html_body = f"""
    <!DOCTYPE html>
    <html lang="en">

    <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>{user_name} - New Starter Onboarding</title>
    </head>

    <body style="margin:0; padding:20px; background-color:#d3d3d3; font-family:'Montserrat','Segoe UI','Roboto','Helvetica Neue','Calibri',sans-serif; color:#343a40; width:100% !important; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%;">

    <div style="display:none; max-height:0; overflow:hidden;">
        {user_name} - New Starter Onboarding
    </div>

    <center>
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%; max-width:960px; margin:0 auto; background-color:#ffffff; border-radius:12px; overflow:hidden; box-shadow:0 0 10px #000000; box-sizing:border-box; padding:32px;">
        <tbody>

            <tr>
            <td style="text-align:center; padding-bottom:20px;">
                <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-OnboardingBanner.jpg"
                alt="Mangano IT - Service Update Banner"
                style="width:100%; height:auto; border-radius:12px; display:block;" />
            </td>
            </tr>

            <tr>
            <td style="text-align:center; padding:20px 0; font-size:28px; font-weight:700; color:#212529;">
                {user_name} - New Starter Onboarding Details - Ticket #{ticket_number}
            </td>
            </tr>

            <tr>
            <td style="text-align:center; padding:0 20px 20px; font-size:16px; font-weight:600; color:#343a40;">
                The user has received their password via text message to their listed mobile.
            </td>
            </tr>

            <tr>
            <td>
                <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%; padding:0 20px 20px;">
                <tr>
                    <td style="padding:8px 0; font-size:16px; width:30%; font-weight:600;">User Name:</td>
                    <td style="padding:8px 0; font-size:16px; width:70%; text-align:right;">{user_name}</td>
                </tr>
                <tr>
                    <td style="padding:8px 0; font-size:16px; width:30%; font-weight:600;">Start Date:</td>
                    <td style="padding:8px 0; font-size:16px; width:70%; text-align:right;">{start_date}</td>
                </tr>
                <tr>
                    <td style="padding:8px 0; font-size:16px; width:30%; font-weight:600;">Email Address:</td>
                    <td style="padding:8px 0; font-size:16px; width:70%; text-align:right;">{user_email}</td>
                </tr>
                <tr>
                    <td style="padding:8px 0; font-size:16px; width:30%; font-weight:600;">Added to Groups:</td>
                    <td style="padding:8px 0; font-size:16px; width:70%; text-align:right;">{formatted_groups}</td>
                </tr>
                <tr>
                    <td style="padding:8px 0; font-size:16px; width:30%; font-weight:600;">Device/s Allocated:</td>
                    <td style="padding:8px 0; font-size:16px; width:70%; text-align:right;">{formatted_devices}</td>
                </tr>
                </table>
            </td>
            </tr>

            <tr>
            <td style="padding:0 20px 20px; font-size:16px; line-height:1.5; color:#343a40;">
                <b>When the user starts, please ensure they can:</b>
                <ul style="margin-top:10px; padding-left:20px;">
                <li style="margin-bottom:8px;">Log into their device</li>
                <li style="margin-bottom:8px;">Connect to the network</li>
                <li style="margin-bottom:8px;">Access all relevant systems</li>
                </ul>
                You will shortly receive an email with the users login details for their first day. Please print this out and hand to them with their device.
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

def generate_first_day_guide():
    subject = "First Day Guide"

    html_body = f"""
    <!DOCTYPE html>
    <html lang="en">

    <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>First Day Guide</title>
    </head>

    <body style="margin:0; padding:20px; background-color:#f4f7fb; font-family:'Montserrat','Calibri',sans-serif; color:#343a40; width:100% !important; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%;">
    <center>
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="max-width:960px; background-color:#ffffff; margin:0 auto; box-shadow:0 6px 18px rgba(0,0,0,0.08); padding:32px; box-sizing:border-box;">
        <tbody>
            <tr>
            <td style="text-align:center;">
                <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-First-Day-Guide.jpg"
                alt="Application Control"
                style="width:100%; height:auto; display:block;" />
            </td>
            </tr>
            <tr>
            <td style="background-color:#9ccc50; text-align:center; padding:40px 20px;">
                <div style="font-size:40px; font-weight:700; color:#35383a; letter-spacing:0.2px; font-family:'Montserrat','Calibri',sans-serif;">
                Where to go if I need help?
                </div>
                <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="max-width:800px; margin:20px auto 0; font-family:'Montserrat','Calibri',sans-serif;">
                <tbody>
                    <tr>
                    <td style="width:50%; padding:20px 30px; font-size:20px; text-align:left; color:#35383a;">
                        Should you have any questions or need assistance, please reach out to the Service Desk via the
                        following options:
                    </td>
                    <td style="width:50%; padding:20px 30px; font-size:20px; text-align:left; color:#35383a;">
                        Service Portal:
                        <a href="https://portal.manganoit.com.au" style="color:#35383a; font-weight:700; text-decoration:underline;">Mangano IT Portal</a><br>
                        Email:
                        <a href="mailto:support@manganoit.com.au" style="color:#35383a; font-weight:700; text-decoration:underline;">support@manganoit.com.au</a><br>
                        Phone:
                        <a href="tel:+61731519000" style="color:#35383a; font-weight:700; text-decoration:underline;">(07) 3151 9000</a>
                    </td>
                    </tr>
                </tbody>
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

def generate_mit_introduction(first_name, company_name):
    subject = "Meet Mangano IT - Your Friendly IT Support Team"

    html_body = f"""
    <!DOCTYPE html>
    <html lang="en">

    <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Meet Mangano IT - Your Friendly IT Support Team</title>
    </head>

    <body style="margin:0; padding:20px; background-color:#d3d3d3; font-family:'Montserrat', 'Segoe UI', 'Roboto', 'Helvetica Neue', 'Calibri', sans-serif; color:#343a40; width:100% !important; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%;">
    <div style="display:none; max-height:0; overflow:hidden;">Meet your IT team at Mangano IT - how to contact us and what to expect.</div>
    <center>
        <table cellpadding="0" cellspacing="0" border="0" role="presentation" style="width:100%; max-width:960px; margin:0 auto; background-color:#ffffff; border-radius:12px; overflow:hidden; box-shadow:0 0 10px #000000; box-sizing:border-box; padding:32px;">
        <tbody>
            <tr>
            <td colspan="3" style="text-align:center;">
                <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-IntroductionToManganoBanner.jpg" alt="Meet Your Support Team - Mangano IT" style="width:100%; height:auto; border-radius:12px; display:block;" />
            </td>
            </tr>
            <tr>
            <td colspan="3" style="font-size:16px; line-height:1.6; padding:10px 20px;">
                Hi {first_name},<br><br>
                Welcome again to {company_name}!<br><br>
                We're Mangano IT, your IT Team. That means we're here behind the scenes making sure your technology works smoothly - and we're just a call or ticket away if you need help.<br><br>
                Whether it's a pesky password issue, a dodgy Wi-Fi connection, or something bigger, our team is ready to jump in and get it sorted. Below is everything you need to know about how to contact us and what to expect.<br><br>
                To view the progress of your ticket at any time or make changes visit our
                <a href="https://portal.manganoit.com.au" style="color:#0066cc; text-decoration:none;">Mangano IT Service Desk Portal</a>
            </td>
            </tr>
            <tr>
            <td colspan="3" style="padding:12px 20px; font-weight:bold; font-size:16px;">
                Print this out and keep it handy!
            </td>
            </tr>
            <tr>
            <td colspan="3" style="padding:12px 20px; font-size:18px; font-weight:600; color:#70ad47;">
                How things are prioritised and what to do
            </td>
            </tr>
            <tr>
            <td colspan="3">
                <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%; border-collapse:separate; border-spacing:0 6px;">
                <thead>
                    <tr>
                    <td style="background-color:#d9d9d9; padding:12px; text-align:center; font-weight:bold;">Priority</td>
                    <td style="background-color:#d9d9d9; padding:12px; text-align:center; font-weight:bold;">Criteria</td>
                    <td style="background-color:#d9d9d9; padding:12px; text-align:center; font-weight:bold;">What to do</td>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                    <td style="font-weight:bold; font-size:20px; text-align:center; padding:20px; color:#ffffff; background-color:#bc381d;">
                        Priority 1<br>
                        <small style="display:block; margin-top:6px; font-size:16px;">Aim to respond within 20mins<br>Resolve within 12hrs</small>
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px;">
                        <ul style="padding-left:20px; margin:0;">
                        <li style="margin-bottom:8px;">Whole company or site is affected and critical/major business processes are stopped</li>
                        </ul>
                        Example: The internet has gone down, there has been a security breach etc
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px; text-align:center;">
                        <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-Phone-Icon.png" alt="Phone Icon" style="width:30px; margin-bottom:10px;"><br>
                        Call us<br>
                        <a href="tel:+61731519000" style="color:#0066cc; text-decoration:none;">(07) 3151 9000</a><br>8am-5pm QLD time
                    </td>
                    </tr>
                    <tr>
                    <td style="font-weight:bold; font-size:20px; text-align:center; padding:20px; color:#ffffff; background-color:#e89600;">
                        Priority 2<br>
                        <small style="display:block; margin-top:6px; font-size:16px;">Aim to respond within 60mins<br>Resolve within 12hrs</small>
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px;">
                        <ul style="padding-left:20px; margin:0;">
                        <li style="margin-bottom:8px;">Whole company or site affected but workaround exists</li>
                        <li style="margin-bottom:8px;">Large group of users blocked on critical functions</li>
                        </ul>
                        Example: A key business app is broken, Wi-Fi is down but docks work
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px; text-align:center;">
                        <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-Phone-Icon.png" alt="Phone Icon" style="width:30px; margin-bottom:10px;"><br>
                        Call us<br>
                        <a href="tel:+61731519000" style="color:#0066cc; text-decoration:none;">(07) 3151 9000</a><br>8am-5pm QLD time
                    </td>
                    </tr>
                    <tr>
                    <td style="font-weight:bold; font-size:20px; text-align:center; padding:20px; color:#000000; background-color:#ffd890;">
                        Priority 3<br>
                        <small style="display:block; margin-top:6px; font-size:16px;">Aim to respond within 2hrs<br>Resolve within 3 days</small>
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px;">
                        <ul style="padding-left:20px; margin:0;">
                        <li style="margin-bottom:8px;">Issue is company-wide but there is a workaround</li>
                        <li style="margin-bottom:8px;">Business critical processes are blocked for one or a few users</li>
                        </ul>
                        Example: Invoices sent are being blocked
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px; text-align:center;">
                        <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-Laptop-Icon.png" alt="Laptop Icon" style="width:30px; margin-bottom:10px;"><br>
                        Log a ticket<br>
                        <a href="https://portal.manganoit.com.au" style="color:#0066cc; text-decoration:none;">Service Desk Portal</a>
                    </td>
                    </tr>
                    <tr>
                    <td style="font-weight:bold; font-size:20px; text-align:center; padding:20px; color:#000000; background-color:#9ccc50;">
                        Priority 4<br>
                        <small style="display:block; margin-top:6px; font-size:16px;">Aim to respond within 4hrs<br>Resolve within 5 days</small>
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px;">
                        <ul style="padding-left:20px; margin:0;">
                        <li style="margin-bottom:8px;">Minor issue for one user or small group</li>
                        </ul>
                        Example: My password has expired
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px; text-align:center;">
                        <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-Laptop-Icon.png" alt="Laptop Icon" style="width:30px; margin-bottom:10px;"><br>
                        Log a ticket<br>
                        <a href="https://portal.manganoit.com.au" style="color:#0066cc; text-decoration:none;">Service Desk Portal</a><br>or<br>
                        Email us<br>
                        <a href="mailto:support@manganoit.com.au" style="color:#0066cc; text-decoration:none;">support@manganoit.com.au</a>
                    </td>
                    </tr>
                    <tr>
                    <td style="font-weight:bold; font-size:20px; text-align:center; padding:20px; color:#000000; background-color:#d9d9d9;">
                        Requests<br>
                        <small style="display:block; margin-top:6px; font-size:16px;">Aim to resolve within 5 days</small>
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px;">
                        <ul style="padding-left:20px; margin:0;">
                        <li style="margin-bottom:8px;">Request for product or service change</li>
                        </ul>
                        Example: Add user to SharePoint group
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px; text-align:center;">
                        <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-Laptop-Icon.png" alt="Laptop Icon" style="width:30px; margin-bottom:10px; margin-top:10px;"><br>
                        Log a ticket<br>
                        <a href="https://portal.manganoit.com.au" style="color:#0066cc; text-decoration:none;">Service Desk Portal</a><br>or<br>
                        Email us<br>
                        <a href="mailto:support@manganoit.com.au" style="color:#0066cc; text-decoration:none;">support@manganoit.com.au</a>
                    </td>
                    </tr>
                </tbody>
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

def generate_qrl_mit_introduction(first_name, company_name):
    subject = "Meet Mangano IT - Your Friendly IT Support Team"

    html_body = f"""
    <!DOCTYPE html>
    <html lang="en">

    <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Meet Mangano IT - Your Friendly IT Support Team</title>
    </head>

    <body style="margin:0; padding:20px; background-color:#d3d3d3; font-family:'Montserrat', 'Segoe UI', 'Roboto', 'Helvetica Neue', 'Calibri', sans-serif; color:#343a40; width:100% !important; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%;">
    <div style="display:none; max-height:0; overflow:hidden;">Meet your IT team at Mangano IT - how to contact us and what to expect.</div>
    <center>
        <table cellpadding="0" cellspacing="0" border="0" role="presentation" style="width:100%; max-width:960px; margin:0 auto; background-color:#ffffff; border-radius:12px; overflow:hidden; box-shadow:0 0 10px #000000; box-sizing:border-box; padding:32px;">
        <tbody>
            <tr>
            <td colspan="3" style="text-align:center;">
                <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/QRL-IntroductionToManganoBanner.jpg" alt="Meet Your Support Team - Mangano IT" style="width:100%; height:auto; border-radius:12px; display:block;" />
            </td>
            </tr>
            <tr>
            <td colspan="3" style="font-size:16px; line-height:1.6; padding:10px 20px;">
                Hi {first_name},<br><br>
                Welcome again to {company_name}!<br><br>
                We're Mangano IT, your IT Team. That means we're here behind the scenes making sure your technology works smoothly - and we're just a call or ticket away if you need help.<br><br>
                Whether it's a pesky password issue, a dodgy Wi-Fi connection, or something bigger, our team is ready to jump in and get it sorted. Below is everything you need to know about how to contact us and what to expect.<br><br>
                To view the progress of your ticket at any time or make changes visit our
                <a href="https://portal.manganoit.com.au" style="color:#0066cc; text-decoration:none;">Mangano IT Service Desk Portal</a>
            </td>
            </tr>
            <tr>
            <td colspan="3" style="padding:12px 20px; font-weight:bold; font-size:16px;">
                Print this out and keep it handy!
            </td>
            </tr>
            <tr>
            <td colspan="3" style="padding:12px 20px; font-size:18px; font-weight:600; color:#70ad47;">
                How things are prioritised and what to do
            </td>
            </tr>
            <tr>
            <td colspan="3">
                <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="width:100%; border-collapse:separate; border-spacing:0 6px;">
                <thead>
                    <tr>
                    <td style="background-color:#d9d9d9; padding:12px; text-align:center; font-weight:bold;">Priority</td>
                    <td style="background-color:#d9d9d9; padding:12px; text-align:center; font-weight:bold;">Criteria</td>
                    <td style="background-color:#d9d9d9; padding:12px; text-align:center; font-weight:bold;">What to do</td>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                    <td style="font-weight:bold; font-size:20px; text-align:center; padding:20px; color:#ffffff; background-color:#bc381d;">
                        Priority 1<br>
                        <small style="display:block; margin-top:6px; font-size:16px;">Aim to respond within 20mins<br>Resolve within 12hrs</small>
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px;">
                        <ul style="padding-left:20px; margin:0;">
                        <li style="margin-bottom:8px;">Whole company or site is affected and critical/major business processes are stopped</li>
                        </ul>
                        Example: The internet has gone down, there has been a security breach etc
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px; text-align:center;">
                        <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-Phone-Icon.png" alt="Phone Icon" style="width:30px; margin-bottom:10px;"><br>
                        Call us<br>
                        <a href="tel:+61731519000" style="color:#0066cc; text-decoration:none;">(07) 3151 9000</a><br>8am-5pm QLD time
                    </td>
                    </tr>
                    <tr>
                    <td style="font-weight:bold; font-size:20px; text-align:center; padding:20px; color:#ffffff; background-color:#e89600;">
                        Priority 2<br>
                        <small style="display:block; margin-top:6px; font-size:16px;">Aim to respond within 60mins<br>Resolve within 12hrs</small>
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px;">
                        <ul style="padding-left:20px; margin:0;">
                        <li style="margin-bottom:8px;">Whole company or site affected but workaround exists</li>
                        <li style="margin-bottom:8px;">Large group of users blocked on critical functions</li>
                        </ul>
                        Example: A key business app is broken, Wi-Fi is down but docks work
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px; text-align:center;">
                        <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-Phone-Icon.png" alt="Phone Icon" style="width:30px; margin-bottom:10px;"><br>
                        Call us<br>
                        <a href="tel:+61731519000" style="color:#0066cc; text-decoration:none;">(07) 3151 9000</a><br>8am-5pm QLD time
                    </td>
                    </tr>
                    <tr>
                    <td style="font-weight:bold; font-size:20px; text-align:center; padding:20px; color:#000000; background-color:#ffd890;">
                        Priority 3<br>
                        <small style="display:block; margin-top:6px; font-size:16px;">Aim to respond within 2hrs<br>Resolve within 3 days</small>
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px;">
                        <ul style="padding-left:20px; margin:0;">
                        <li style="margin-bottom:8px;">Issue is company-wide but there is a workaround</li>
                        <li style="margin-bottom:8px;">Business critical processes are blocked for one or a few users</li>
                        </ul>
                        Example: Invoices sent are being blocked
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px; text-align:center;">
                        <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-Laptop-Icon.png" alt="Laptop Icon" style="width:30px; margin-bottom:10px;"><br>
                        Log a ticket<br>
                        <a href="https://portal.manganoit.com.au" style="color:#0066cc; text-decoration:none;">Service Desk Portal</a>
                    </td>
                    </tr>
                    <tr>
                    <td style="font-weight:bold; font-size:20px; text-align:center; padding:20px; color:#000000; background-color:#9ccc50;">
                        Priority 4<br>
                        <small style="display:block; margin-top:6px; font-size:16px;">Aim to respond within 4hrs<br>Resolve within 5 days</small>
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px;">
                        <ul style="padding-left:20px; margin:0;">
                        <li style="margin-bottom:8px;">Minor issue for one user or small group</li>
                        </ul>
                        Example: My password has expired
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px; text-align:center;">
                        <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-Laptop-Icon.png" alt="Laptop Icon" style="width:30px; margin-bottom:10px;"><br>
                        Log a ticket<br>
                        <a href="https://portal.manganoit.com.au" style="color:#0066cc; text-decoration:none;">Service Desk Portal</a><br>or<br>
                        Email us<br>
                        <a href="mailto:support@manganoit.com.au" style="color:#0066cc; text-decoration:none;">support@manganoit.com.au</a>
                    </td>
                    </tr>
                    <tr>
                    <td style="font-weight:bold; font-size:20px; text-align:center; padding:20px; color:#000000; background-color:#d9d9d9;">
                        Requests<br>
                        <small style="display:block; margin-top:6px; font-size:16px;">Aim to resolve within 5 days</small>
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px;">
                        <ul style="padding-left:20px; margin:0;">
                        <li style="margin-bottom:8px;">Request for product or service change</li>
                        </ul>
                        Example: Add user to SharePoint group
                    </td>
                    <td style="background-color:#ffffff; padding:0 0 0 20px; font-size:16px; text-align:center;">
                        <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-Laptop-Icon.png" alt="Laptop Icon" style="width:30px; margin-bottom:10px; margin-top:10px;"><br>
                        Log a ticket<br>
                        <a href="https://portal.manganoit.com.au" style="color:#0066cc; text-decoration:none;">Service Desk Portal</a><br>or<br>
                        Email us<br>
                        <a href="mailto:support@manganoit.com.au" style="color:#0066cc; text-decoration:none;">support@manganoit.com.au</a>
                    </td>
                    </tr>
                </tbody>
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
        log.info(f"Preparing to send email from [{sender_email}] to [{recipient_emails}] with subject [{subject}]")

        masked_subject = re.sub(r"('|\")(eyJ[a-zA-Z0-9_-]{5,}?\.[a-zA-Z0-9_-]{5,}?\.([a-zA-Z0-9_-]{5,})?)", r"\1***TOKEN-MASKED***\1", subject)
        masked_subject = re.sub(r"('|\")([a-zA-Z0-9]{8,}-(clientid|clientsecret|password))\1", r"\1***SECRET-MASKED***\1", masked_subject, flags=re.IGNORECASE)
        
        masked_html_body = re.sub(r"('|\")(eyJ[a-zA-Z0-9_-]{5,}?\.[a-zA-Z0-9_-]{5,}?\.([a-zA-Z0-9_-]{5,})?)", r"\1***TOKEN-MASKED***\1", html_body)
        masked_html_body = re.sub(r"('|\")([a-zA-Z0-9]{8,}-(clientid|clientsecret|password))\1", r"\1***SECRET-MASKED***\1", masked_html_body, flags=re.IGNORECASE)

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
                "subject": masked_subject,
                "body": {
                    "contentType": "HTML",
                    "content": masked_html_body
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

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1745975394407")
            operation = input.get_value("Operation_1749629031429")
            recipient_emails = input.get_value("Recipients_1745974383560")
            first_name = input.get_value("FirstName_1745974704285")
            user_upn = input.get_value("UserPrincipalName_1745974706233")
            user_name = input.get_value("UserFullName_1745974994826")
            user_email = input.get_value("UserEmail_1745974996616")
            start_date = input.get_value("StartDate_1749629966002")
            groups = input.get_value("Groups_1745975216435")
            devices = input.get_value("Devices_1745975226614")
            mit_graph_token = input.get_value("MITGraphToken_1758666595949")
            graph_token = input.get_value("GraphToken_1745975397197")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if isinstance(ticket_number, str) else ""
        recipient_emails = recipient_emails.strip() if recipient_emails else ""
        first_name = first_name.strip() if first_name else ""
        user_upn = user_upn.strip() if user_upn else ""
        user_name = user_name.strip() if user_name else ""
        user_email = user_email.strip() if user_email else ""
        start_date = format_date(start_date.strip()) if start_date else ""
        groups = groups.strip() if groups else ""
        devices = devices.strip() if devices else ""
        graph_token = graph_token.strip() if graph_token else ""
        mit_graph_token = mit_graph_token.strip() if mit_graph_token else ""
        operation = operation.strip() if isinstance(operation, str) else ""

        log.info(f"Ticket Number = [{ticket_number}]")
        log.info(f"Requested operation = [{operation}]")
        log.info(f"User Name = [{user_name}]")

        if not operation:
            record_result(log, ResultLevel.INFO, "No operation selected. No action taken")
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
            graph_tenant_id, graph_access_token = get_graph_token(log, http_client, vault_name, company_identifier)
            if not graph_access_token:
                record_result(log, ResultLevel.WARNING, "Failed to obtain MS Graph access token")
                return

        azure_domain = get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")
        if not azure_domain:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve Azure domain for [{company_identifier}]")
            return

        if operation == "Send User onboarding notification for users":
            log.info("Executing operation: Send User onboarding notification for users")
            subject, html_body = generate_user_onboarding_email(first_name, company_name, user_upn)
        
        elif operation == "Send User onboarding notification for manager":
            log.info("Executing operation: Send User onboarding notification for manager")
            subject, html_body = generate_manager_onboarding_email(user_name, start_date, user_email, groups, devices, ticket_number)
        
        elif operation == "Send Introducing Mangano IT":
            log.info("Executing operation: Send Introducing Mangano IT")
            subject, html_body = generate_mit_introduction(first_name, company_name)
        
        elif operation == "Send first day guide":
            log.info("Executing operation: Send first day guide")
            subject, html_body = generate_first_day_guide()
        
        elif operation == "QRL - Send Introducing Mangano IT":
            log.info("Executing operation: QRL - Send Introducing Mangano IT")
            subject, html_body = generate_qrl_mit_introduction(first_name, company_name)
        
        else:
            record_result(log, ResultLevel.WARNING, f"Unsupported operation: {operation}")
            return

        if send_email(log, http_client, msgraph_base_url, mit_graph_access_token, sender_email, recipient_emails, subject, html_body):
            record_result(log, ResultLevel.SUCCESS, f"Sent [{operation}] notification to [{recipient_emails}]")
        else:
            record_result(log, ResultLevel.WARNING, f"Failed to send [{operation}] notification to [{recipient_emails}]")

    except Exception as e:
        log.exception(e, "Unhandled error in main")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
