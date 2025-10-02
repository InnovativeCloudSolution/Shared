import sys
import random
import os
import base64
import time
import urllib.parse
import requests
import pandas as pd
from datetime import datetime, timedelta, timezone
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

cwpsa_base_url = "https://au.myconnectwise.net/v4_6_release/apis/3.0"
msgraph_base_url = "https://graph.microsoft.com/v1.0"
vault_name = "mit-azu1-prod1-akv1"
sender_email = "support@manganoit.com.au"
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
    if response and response.status_code == 200:
        return response.json().get("value", "")
    return ""


def get_tenant_id_from_domain(log, http_client, azure_domain):
    config_url = f"https://login.windows.net/{azure_domain}/.well-known/openid-configuration"
    response = execute_api_call(log, http_client, "get", config_url)
    if response and response.status_code == 200:
        token_endpoint = response.json().get("token_endpoint", "")
        return token_endpoint.split("/")[3] if token_endpoint else ""
    return ""


def get_access_token(log, http_client, tenant_id, client_id, client_secret):
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    payload = urllib.parse.urlencode({
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default"
    })
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    response = execute_api_call(log, http_client, "post", token_url, data=payload, headers=headers)
    if response and response.status_code == 200:
        return response.json().get("access_token", "")
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
    else:
        log.error(f"No response received when retrieving ticket [{ticket_number}]")

    return "", "", 0, []


def get_ticket_data(log, http_client, cwpsa_base_url, ticket_number):
    try:
        log.info(f"Retrieving full ticket details for ticket number [{ticket_number}]")
        endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}"
        response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
        if not response or response.status_code != 200:
            log.error(
                f"Failed to retrieve ticket [{ticket_number}] - Status: {response.status_code if response else 'N/A'}"
            )
            return "", "", "", ""

        ticket = response.json()
        ticket_summary = ticket.get("summary", "")
        ticket_type = ticket.get("type", {}).get("name", "")
        priority_name = ticket.get("priority", {}).get("name", "")
        due_date = ticket.get("requiredDate", "")

        log.info(
            f"Ticket [{ticket_number}] Summary = [{ticket_summary}], Type = [{ticket_type}], Priority = [{priority_name}], Due = [{due_date}]"
        )
        return ticket_summary, ticket_type, priority_name, due_date

    except Exception as e:
        log.exception(e, f"Exception occurred while retrieving ticket details for [{ticket_number}]")
        return "", "", "", ""


def validate_mit_authentication(log, http_client, vault_name, auth_code):
    if not auth_code:
        log.result_message(ResultLevel.FAILED, "Authentication code input is required for MIT")
        return False
    expected_code = get_secret_value(log, http_client, vault_name, "MIT-AuthenticationCode")
    if not expected_code:
        log.result_message(
            ResultLevel.FAILED,
            "Failed to retrieve expected authentication code for MIT",
        )
        return False
    if auth_code.strip() != expected_code.strip():
        log.result_message(ResultLevel.FAILED, "Provided authentication code is incorrect")
        return False
    return True


def send_email(log, http_client, msgraph_base_url, access_token, sender_email, recipient_emails, subject, html_body, attachment_path=None):
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

        message = {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": html_body
            },
            "toRecipients": to_recipients
        }

        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, "rb") as file:
                content_bytes = base64.b64encode(file.read()).decode("utf-8")
                attachment = {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": os.path.basename(attachment_path),
                    "contentBytes": content_bytes
                }
                message["attachments"] = [attachment]
                log.info(f"Attached file [{attachment_path}] to email")

        email_message = {
            "message": message,
            "saveToSentItems": "true"
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


def generate_user_license_report(log, http_client, msgraph_base_url, access_token):
    try:
        log.info("Generating user license report")

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }

        sku_url = "https://mitazu1pubfilestore.blob.core.windows.net/automation/M365SKU.csv"
        sku_response = execute_api_call(log, http_client, "get", sku_url)

        sku_map = {}
        if sku_response and sku_response.status_code == 200:
            from io import StringIO
            import pandas as pd
            sku_df = pd.read_csv(StringIO(sku_response.text))
            sku_map = {row["SKU"]: row["PN"] for _, row in sku_df.iterrows() if row.get("SKU") and row.get("PN")}
        else:
            log.warning("Could not load SKU mapping. License names will default to raw SKU")

        users = []
        next_link = f"{msgraph_base_url}/users?$select=displayName,userPrincipalName,id&$top=999"
        while next_link:
            response = execute_api_call(log, http_client, "get", next_link, headers=headers)
            if response and response.status_code == 200:
                data = response.json()
                users.extend(data.get("value", []))
                next_link = data.get("@odata.nextLink", None)
            else:
                log.warning("Failed to fetch users")
                break

        if not users:
            log.warning("No users found")
            return ""

        rows = []
        for user in users:
            display_name = user.get("displayName", "")
            user_id = user.get("id", "")
            user_upn = user.get("userPrincipalName", "")
            upn_domain = user_upn.split("@")[-1].replace(".com.au", "") if "@" in user_upn else ""

            if not user_id:
                continue

            license_endpoint = f"{msgraph_base_url}/users/{user_id}/licenseDetails"
            license_response = execute_api_call(log, http_client, "get", license_endpoint, headers=headers)

            if not license_response or license_response.status_code != 200:
                continue

            license_data = license_response.json().get("value", [])
            for entry in license_data:
                sku_id = entry.get("skuPartNumber", "").strip()
                friendly_name = sku_map.get(sku_id, sku_id)

                rows.append({
                    "Display Name": display_name,
                    "UPN Domain": upn_domain,
                    "License Name": friendly_name
                })

        if not rows:
            log.warning("No license data found")
            return ""

        import pandas as pd
        df = pd.DataFrame(rows)
        report_path = "/opt/app/RPA-execute/UserLicenseReport.csv"
        df.to_csv(report_path, index=False)
        log.info(f"User license report saved to [{report_path}]")
        return report_path

    except Exception as e:
        log.exception(e, "Exception occurred while generating license report")
        return ""


def generate_license_email_body(log, company_name):
    try:
        today = datetime.utcnow().strftime("%d %b %Y")
        html = f"""
        <html>
        <body>
            <p>Hi,</p>
            <p>Please find attached the Microsoft 365 user license report for <strong>{company_name}</strong>.</p>
            <p>If you have any questions, please contact the Mangano IT team.</p>
            <p>Regards,<br>Mangano IT</p>
        </body>
        </html>
        """
        return html
    except Exception as e:
        log.exception(e, "Failed to generate email body")
        return "<p>Failed to generate email body.</p>"


def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1752442372951")
            auth_code = input.get_value("AuthCode_1752442375576")
            recipient_emails = input.get_value("Recipient_1752442374360")
            provided_token = input.get_value("AccessToken_1752442377047")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        auth_code = auth_code.strip() if auth_code else ""
        recipient_emails = recipient_emails.strip() if recipient_emails else ""
        provided_token = provided_token.strip() if provided_token else ""

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number is missing or invalid")
            return
        if not recipient_emails:
            record_result(log, ResultLevel.WARNING, "Recipient email(s) are missing")
            return

        company_identifier, company_name, _, _ = get_company_data_from_ticket(
            log, http_client, cwpsa_base_url, ticket_number
        )

        if not company_identifier or not company_name:
            record_result(log, ResultLevel.WARNING, f"Failed to resolve company details from ticket [{ticket_number}]")
            return

        if company_identifier == "MIT":
            expected_code = get_secret_value(log, http_client, vault_name, "MIT-AuthenticationCode")
            if auth_code != expected_code:
                record_result(log, ResultLevel.WARNING, "Provided authentication code is incorrect")
                return

        if provided_token:
            tenant_token = provided_token
            log.info("Using provided access token")
            if not isinstance(tenant_token, str) or "." not in tenant_token:
                record_result(log, ResultLevel.WARNING, "Provided access token is malformed (missing dots)")
                return
        else:
            client_id = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientID")
            client_secret = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientSecret")
            azure_domain = get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")

            if not all([client_id, client_secret, azure_domain]):
                record_result(log, ResultLevel.WARNING, "Failed to retrieve required secrets")
                return

            tenant_id = get_tenant_id_from_domain(log, http_client, azure_domain)
            if not tenant_id:
                record_result(log, ResultLevel.WARNING, "Unable to resolve tenant ID")
                return

            tenant_token = get_access_token(log, http_client, tenant_id, client_id, client_secret)
            if not isinstance(tenant_token, str) or "." not in tenant_token:
                record_result(log, ResultLevel.WARNING, "Access token is malformed (missing dots)")
                return

        report_path = generate_user_license_report(log, http_client, msgraph_base_url, tenant_token)
        if not report_path:
            record_result(log, ResultLevel.WARNING, "Failed to generate user license report")
            return

        mangano_client_id = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientID")
        mangano_client_secret = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientSecret")
        mangano_domain = "manganoit.com.au"
        mangano_tenant_id = get_tenant_id_from_domain(log, http_client, mangano_domain)
        if not mangano_tenant_id:
            record_result(log, ResultLevel.WARNING, "Unable to resolve Mangano tenant ID")
            return

        mangano_token = get_access_token(log, http_client, mangano_tenant_id, mangano_client_id, mangano_client_secret)
        if not isinstance(mangano_token, str) or "." not in mangano_token:
            record_result(log, ResultLevel.WARNING, "Mangano access token is malformed")
            return

        subject = f"User License Report for {company_name}"
        html_body = generate_license_email_body(log, company_name)

        if send_email(log, http_client, msgraph_base_url, mangano_token, sender_email, recipient_emails, subject, html_body, report_path):
            if isinstance(recipient_emails, str):
                recipient_emails = [email.strip() for email in recipient_emails.split(",") if email.strip()]

            for email in recipient_emails:
                record_result(log, ResultLevel.SUCCESS, f"License report emailed successfully to [{email}]")
        else:
            record_result(log, ResultLevel.WARNING, "License report generated but email failed to send")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()