import sys
import random
import os
import time
import urllib.parse
import re
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
vault_name = "mit-azu1-prod1-akv1"
sender_email = "support@manganoit.com.au"

data_to_log = {}
bot_name = "MIT-UTIL - Send technology stack guides"
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

    log.error(f"[{log_prefix}] Failed to retrieve access token Status code: {response.status_code if response else 'N/A'}")
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
    
    token = get_access_token(log, http_client, tenant_id, client_id, client_secret, scope="https://graph.microsoft.com/.default", log_prefix="Graph")
    if not isinstance(token, str) or "." not in token:
        log.error("MS Graph access token is malformed for MIT domain")
        return "", ""
    
    log.info("Successfully obtained MS Graph token for MIT domain")
    return tenant_id, token

def generate_mit_application_control():
    subject = "Application Control"

    html_body ="""
    <!DOCTYPE html>
    <html lang="en">

    <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Email Signature Management</title>
    </head>

    <body
    style="margin:0; padding:20px; background-color:#f4f7fb; font-family:'Montserrat', 'Calibri', sans-serif; color:#343a40; width:100% !important; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%;">
    <center>
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
        style="max-width:960px; background-color:#ffffff; margin:0 auto; box-shadow:0 6px 18px rgba(0,0,0,0.08); padding:32px; box-sizing:border-box;">
        <tbody>

            <tr>
            <td style="text-align:center;">
                <img
                src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-Application-Control.jpg"
                alt="Application Control" style="width:100%; height:auto; display:block;" />
            </td>
            </tr>

            <tr>
            <td style="background-color: #e5e5e4; text-align:center; padding:40px 20px;">
                <div
                style="font-size:40px; font-weight:700; color:#35383a; letter-spacing:0.2px; font-family:'Montserrat','Calibri',sans-serif;">
                Where to go if I need help?
                </div>
                <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
                style="max-width:800px; margin:20px auto 0; font-family:'Montserrat','Calibri',sans-serif;">
                <tbody>
                    <tr>
                    <td style="width:50%; padding:20px 30px; font-size:20px; text-align:left; color:#35383a;">
                        Should you have any questions or need assistance, please reach out to the Service Desk via the
                        following options:</td>
                    <td style="width:50%; padding:20px 30px; font-size:20px; text-align:left; color:#35383a;">
                        Service Portal: <a href="https://portal.manganoit.com.au"
                        style="color:#35383a; font-weight:700; text-decoration:none; text-decoration: underline;">Mangano
                        IT Portal</a><br>
                        Email: <a href="mailto:support@manganoit.com.au"
                        style="color:#35383a; font-weight:700; text-decoration:none; text-decoration: underline;">support@manganoit.com.au</a><br>
                        Phone: <a href="tel:+61731519000"
                        style="color:#35383a; font-weight:700; text-decoration:none; text-decoration: underline;">(07)
                        3151 9000</a>
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

def generate_mit_email_signature_management():
    subject = "Email Signature Management"

    html_body ="""
    <!DOCTYPE html>
    <html lang="en">

    <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Email Signature Management</title>
    </head>

    <body
    style="margin:0; padding:20px; background-color:#f4f7fb; font-family:'Montserrat', 'Calibri', sans-serif; color:#343a40; width:100% !important; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%;">
    <center>
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
        style="max-width:960px; background-color:#ffffff; margin:0 auto; box-shadow:0 6px 18px rgba(0,0,0,0.08); padding:32px; box-sizing:border-box;">
        <tbody>

            <tr>
            <td style="text-align:center;">
                <img
                src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-Email-Signature-Management.jpg"
                alt="Application Control" style="width:100%; height:auto; display:block;" />
            </td>
            </tr>

            <tr>
            <td style="background-color: #e5e5e4; text-align:center; padding:40px 20px;">
                <div
                style="font-size:40px; font-weight:700; color:#35383a; letter-spacing:0.2px; font-family:'Montserrat','Calibri',sans-serif;">
                Where to go if I need help?
                </div>
                <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
                style="max-width:800px; margin:20px auto 0; font-family:'Montserrat','Calibri',sans-serif;">
                <tbody>
                    <tr>
                    <td style="width:50%; padding:20px 30px; font-size:20px; text-align:left; color:#35383a;">
                        Should you have any questions or need assistance, please reach out to the Service Desk via the
                        following options:</td>
                    <td style="width:50%; padding:20px 30px; font-size:20px; text-align:left; color:#35383a;">
                        Service Portal: <a href="https://portal.manganoit.com.au"
                        style="color:#35383a; font-weight:700; text-decoration:none; text-decoration: underline;">Mangano
                        IT Portal</a><br>
                        Email: <a href="mailto:support@manganoit.com.au"
                        style="color:#35383a; font-weight:700; text-decoration:none; text-decoration: underline;">support@manganoit.com.au</a><br>
                        Phone: <a href="tel:+61731519000"
                        style="color:#35383a; font-weight:700; text-decoration:none; text-decoration: underline;">(07)
                        3151 9000</a>
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

def generate_mit_company_portal():
    subject = "Company Portal"

    html_body ="""
    <!DOCTYPE html>
    <html lang="en">

    <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Company Portal</title>
    </head>

    <body
    style="margin:0; padding:20px; background-color:#f4f7fb; font-family:'Montserrat', 'Calibri', sans-serif; color:#343a40; width:100% !important; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%;">
    <center>
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
        style="max-width:960px; background-color:#ffffff; margin:0 auto; box-shadow:0 6px 18px rgba(0,0,0,0.08); padding:32px; box-sizing:border-box;">
        <tbody>

            <tr>
            <td style="text-align:center;">
                <img src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-Company-Portal.jpg"
                alt="Application Control" style="width:100%; height:auto; display:block;" />
                <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
                style="max-width: 710px; margin: 0 auto; background-color: #ffffff; border-radius: 30px; padding: 10px 0 0 0; letter-spacing:0.2px; font-family:'Montserrat','Calibri',sans-serif;">
                <tr>
                    <td>
                    <div style="font-size:40px; font-weight:700; color:#35383a; text-align: center;">
                        Device Setup Guides
                    </div>
                    <div style="text-align: left;">
                        <ul style="display: inline-block; font-size: 20px; color: #343a40; margin: 10px; padding: 0;">
                        <li><strong><a href="https://mitazu1pubfilestore.blob.core.windows.net/qrl/iOS-BYOD-Enrolment-Guide.pdf"
                                style="color: #9ccc50; text-decoration: underline;">Click here</a></strong> to access the step-by-step guide for set up your personal iPhone</li>
                        <li><strong><a href="https://mitazu1pubfilestore.blob.core.windows.net/qrl/iOS-Corporate-Setup-Guide.pdf"
                                style="color: #9ccc50; text-decoration: underline;">Click here</a></strong> to access the step-by-step guide for your company-issued iPhone</li>
                        <li><strong><a href="https://mitazu1pubfilestore.blob.core.windows.net/qrl/Android-BYOD-Enrolment-Guide.pdf"
                                style="color: #9ccc50; text-decoration: underline;">Click here</a></strong> to access the step-by-step guide to set up your personal Android device</li>
                        </ul>
                    </div>
                    </td>
                </tr>
                </table>
            </td>
            </tr>

            <tr>
            <td style="background-color: #047cc2; text-align:center; padding:40px 20px;">
                <div
                style="font-size:40px; font-weight:700; color:#ffffff; letter-spacing:0.2px; font-family:'Montserrat','Calibri',sans-serif;">
                Where to go if I need help?
                </div>
                <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
                style="max-width:800px; margin:20px auto 0; font-family:'Montserrat','Calibri',sans-serif;">
                <tbody>
                    <tr>
                    <td style="width:50%; padding:20px 30px; font-size:20px; text-align:left; color:#ffffff;">
                        Should you have any questions or need assistance, please reach out to the Service Desk via the
                        following options:</td>
                    <td style="width:50%; padding:20px 30px; font-size:20px; text-align:left; color:#ffffff;">
                        Service Portal: <a href="https://portal.manganoit.com.au"
                        style="color:#9ccc50; font-weight:700; text-decoration:none; text-decoration: underline;">Mangano
                        IT Portal</a><br>
                        Email: <a href="mailto:support@manganoit.com.au"
                        style="color:#9ccc50; font-weight:700; text-decoration:none; text-decoration: underline;">support@manganoit.com.au</a><br>
                        Phone: <a href="tel:+61731519000"
                        style="color:#9ccc50; font-weight:700; text-decoration:none; text-decoration: underline;">(07)
                        3151 9000</a>
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

def generate_mit_password_manager():
    subject = "Password Manager"

    html_body ="""
    <!DOCTYPE html>
    <html lang="en">

    <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Password Manager</title>
    </head>

    <body
    style="margin:0; padding:20px; background-color:#f4f7fb; font-family:'Montserrat', 'Calibri', sans-serif; color:#343a40; width:100% !important; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%;">
    <center>
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
        style="max-width:960px; background-color:#ffffff; margin:0 auto; box-shadow:0 6px 18px rgba(0,0,0,0.08); padding:32px; box-sizing:border-box;">
        <tbody>

            <tr>
            <td style="text-align:center;">
                <img
                src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-Password-Manager-01.jpg"
                alt="Application Control" style="width:100%; height:auto; display:block;" />
            </td>
            </tr>

            <tr>
            <td style="text-align:center;">
                <img
                src="https://mitazu1pubfilestore.blob.core.windows.net/connectwiseemailimages/MIT-Password-Manager-02.jpg"
                alt="Application Control" style="width:100%; height:auto; display:block;" />
            </td>
            </tr>

            <tr>
            <td style="background-color: #e5e5e4; text-align:center; padding:40px 20px;">
                <div
                style="font-size:40px; font-weight:700; color:#35383a; letter-spacing:0.2px; font-family:'Montserrat','Calibri',sans-serif;">
                Where to go if I need help?
                </div>
                <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
                style="max-width:800px; margin:20px auto 0; font-family:'Montserrat','Calibri',sans-serif;">
                <tbody>
                    <tr>
                    <td style="width:50%; padding:20px 30px; font-size:20px; text-align:left; color:#35383a;">
                        Should you have any questions or need assistance, please reach out to the Service Desk via the
                        following options:</td>
                    <td style="width:50%; padding:20px 30px; font-size:20px; text-align:left; color:#35383a;">
                        Service Portal: <a href="https://portal.manganoit.com.au"
                        style="color:#35383a; font-weight:700; text-decoration:none; text-decoration: underline;">Mangano
                        IT Portal</a><br>
                        Email: <a href="mailto:support@manganoit.com.au"
                        style="color:#35383a; font-weight:700; text-decoration:none; text-decoration: underline;">support@manganoit.com.au</a><br>
                        Phone: <a href="tel:+61731519000"
                        style="color:#35383a; font-weight:700; text-decoration:none; text-decoration: underline;">(07)
                        3151 9000</a>
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

        endpoint = f"{msgraph_base_url}/users/{sender_email}/sendMail"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }

        subject_masked = re.sub(r'(\"|\")(eyJ[a-zA-Z0-9_-]{5,}?\.[a-zA-Z0-9_-]{5,}?\.([a-zA-Z0-9_-]{5,})?)\1', r'\1***TOKEN-MASKED***\1', subject)
        subject_masked = re.sub(r'(\"|\")( [a-zA-Z0-9]{8,}-(clientid|clientsecret|password))\1', r'\1***SECRET-MASKED***\1', subject_masked, flags=re.IGNORECASE)
        html_body_masked = re.sub(r'(\"|\")(eyJ[a-zA-Z0-9_-]{5,}?\.[a-zA-Z0-9_-]{5,}?\.([a-zA-Z0-9_-]{5,})?)\1', r'\1***TOKEN-MASKED***\1', html_body)
        html_body_masked = re.sub(r'(\"|\")( [a-zA-Z0-9]{8,}-(clientid|clientsecret|password))\1', r'\1***SECRET-MASKED***\1', html_body_masked, flags=re.IGNORECASE)

        if isinstance(recipient_emails, str):
            recipient_emails = [email.strip() for email in recipient_emails.split(",") if email.strip()]

        to_recipients = [{"emailAddress": {"address": email}} for email in recipient_emails]

        email_message = {
            "message": {
                "subject": subject_masked,
                "body": {
                    "contentType": "HTML",
                    "content": html_body_masked
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
            recipient_emails = input.get_value("Recipients_1746145174359")
            operation = input.get_value("Operation_1749635761818")
            mit_graph_token = input.get_value("MITGraphToken_1758575271903")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        recipient_emails = recipient_emails.strip() if recipient_emails else ""
        operation = operation.strip() if operation else ""
        mit_graph_token = mit_graph_token.strip() if mit_graph_token else ""

        log.info(f"Recipients = [{recipient_emails}]")
        log.info(f"Requested operation = [{operation}]")

        if not recipient_emails or not operation:
            record_result(log, ResultLevel.WARNING, "Bot completed with no recipient or operation selected. No action taken")
            return

        if mit_graph_token:
            log.info("Using provided MIT MS Graph token")
            mit_graph_access_token = mit_graph_token
            mit_graph_tenant_id = ""
        else:
            mit_graph_tenant_id, mit_graph_access_token = get_graph_token_MIT(log, http_client, vault_name)
            if not mit_graph_access_token:
                record_result(log, ResultLevel.WARNING, "Failed to obtain MIT MS Graph access token")
                return

        if operation == "Send Application Control Email":
            log.info("Executing operation: Application Control")
            subject, html_body = generate_mit_application_control()
        
        elif operation == "Send Email Signature Management Email":
            log.info("Executing operation: Email Signature Management")
            subject, html_body = generate_mit_email_signature_management()
        
        elif operation == "Send Company Portal Email":
            log.info("Executing operation: Company Portal")
            subject, html_body = generate_mit_company_portal()
        
        elif operation == "Send Password Management Email":
            log.info("Executing operation: Password Management")
            subject, html_body = generate_mit_password_manager()
        
        else:
            record_result(log, ResultLevel.WARNING, f"Unrecognized operation: {operation}")
            return

        log.info(f"Attempting to send {operation} email")
        if send_email(log, http_client, msgraph_base_url, mit_graph_access_token, sender_email, recipient_emails, subject, html_body):
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
