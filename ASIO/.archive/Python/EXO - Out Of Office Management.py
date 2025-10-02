import sys
import random
import re
import subprocess
import os
import time
import urllib.parse
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

data_to_log = {}
bot_name = "EXO - Out Of Office Management"
log.info("Static variables set")

def record_result(log, level, message):
    log.result_message(level, f"[{bot_name}]: {message}")

    if level == ResultLevel.WARNING:
        data_to_log["Result"] = "Fail"
    elif level == ResultLevel.SUCCESS:
        if "Result" not in data_to_log or data_to_log["Result"] != "Fail":
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

def execute_powershell(log, command, shell="pwsh"):
    try:
        if shell not in ("pwsh", "powershell"):
            raise ValueError("Invalid shell specified. Use 'pwsh' or 'powershell'")

        result = subprocess.run([shell, "-Command", command], capture_output=True, text=True)

        if result.returncode == 0 and not result.stderr.strip():
            log.info(f"{shell} command executed successfully")
            return True, result.stdout.strip()
        else:
            error_message = result.stderr.strip() or result.stdout.strip()
            log.error(f"{shell} execution failed: {error_message}")
            return False, error_message

    except Exception as e:
        log.exception(e, f"Exception occurred during {shell} execution")
        return False, str(e)

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

def get_exo_token(log, http_client, vault_name, company_identifier):
    client_id = get_secret_value(log, http_client, vault_name, f"{company_identifier}-ExchangeApp-ClientID")
    client_secret = get_secret_value(log, http_client, vault_name, f"{company_identifier}-ExchangeApp-ClientSecret")
    azure_domain = get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")

    if not all([client_id, client_secret, azure_domain]):
        log.error("Failed to retrieve required secrets for EXO")
        return "", ""

    tenant_id = get_tenant_id_from_domain(log, http_client, azure_domain)
    if not tenant_id:
        log.error("Failed to resolve tenant ID for EXO")
        return "", ""

    token = get_access_token(
        log, http_client, tenant_id, client_id, client_secret,
        scope="https://outlook.office365.com/.default", log_prefix="EXO"
    )
    if not isinstance(token, str) or "." not in token:
        log.error("EXO access token is malformed (missing dots)")
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
    if response and response.status_code == 200:
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
    else:
        log.error(f"No response received when retrieving ticket [{ticket_number}]")

    return "", "", 0, []

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

def get_aad_user_data(log, http_client, msgraph_base_url, user_identifier, token):
    log.info(f"Resolving user ID and email for [{user_identifier}]")
    headers = {"Authorization": f"Bearer {token}"}

    if re.fullmatch(
        r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}",
        user_identifier,
    ):
        endpoint = f"{msgraph_base_url}/users/{user_identifier}"
        response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
        if response and response.status_code == 200:
            user = response.json()
            return (
                user.get("id", ""),
                user.get("userPrincipalName", ""),
                user.get("onPremisesSamAccountName", ""),
                user.get("onPremisesSyncEnabled", False)
            )
        return "", "", "", False

    filters = [
        f"startswith(displayName,'{user_identifier}')",
        f"startswith(userPrincipalName,'{user_identifier}')",
        f"startswith(mail,'{user_identifier}')",
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
            return (
                user.get("id", ""),
                user.get("userPrincipalName", ""),
                user.get("onPremisesSamAccountName", ""),
                user.get("onPremisesSyncEnabled", False)
            )
    log.error(f"Failed to resolve user ID and email for [{user_identifier}]")
    return "", "", "", False

def execute_out_of_office_command(log, azure_domain, user_email, exo_access_token, ps_command_body, operation_type="Out of Office"):
    ps_command = f"""
      $ErrorActionPreference = 'Stop'
      $ProgressPreference = 'SilentlyContinue'

      if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {{
         Write-Output "Installing ExchangeOnlineManagement module..."
         Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
      }}

      try {{
         Import-Module ExchangeOnlineManagement -ErrorAction Stop
         
         Connect-ExchangeOnline -AccessToken '{exo_access_token}' -Organization '{azure_domain}' -ShowBanner:$false -ErrorAction Stop

         {ps_command_body}
      }} 
      catch {{
         Write-Output "ERROR: Failed to execute {operation_type} operation for [{user_email}] - $_"
      }} 
      finally {{
         try {{ Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }} catch {{ }}
      }}
    """

    success, output = execute_powershell(log, ps_command)
    if not success:
        log.error(f"Failed to execute {operation_type} operation: {output}")
        return False, output
    return True, output

def set_out_of_office_manual(log, azure_domain, user_email, internal_message, external_message, external_audience, exo_access_token):
    log.info(f"Configuring Manual Out Of Office for [{user_email}]")

    if external_audience:
        external_audience = "All"
        log.info(f"Setting to both internal and external audiences for [{user_email}]")
    else:
        external_audience = "None"
        log.info(f"Setting to internal audience only for [{user_email}]")

    if not internal_message.strip().lower().startswith("<html>"):
        internal_message = f"<html><body><p>{internal_message.replace('\n', '<br>')}</p></body></html>"
    if not external_message.strip().lower().startswith("<html>"):
        external_message = f"<html><body><p>{external_message.replace('\n', '<br>')}</p></body></html>"

    ps_command_body = f"""
      $internalMessage = @"
      {internal_message}
      "@
      $externalMessage = @"
      {external_message}
      "@
      Set-MailboxAutoReplyConfiguration -Identity '{user_email}' `
         -AutoReplyState Enabled `
         -InternalMessage $internalMessage `
         -ExternalMessage $externalMessage `
         -ExternalAudience {external_audience}
      $result = Get-MailboxAutoReplyConfiguration -Identity '{user_email}'
      Write-Output "SUCCESS: Set out of office for [{user_email}]"
    """
    return execute_out_of_office_command(log, azure_domain, user_email, exo_access_token, ps_command_body, "manual Out of Office")

def set_out_of_office_with_delegate(log, azure_domain, user_email, delegate_email, exo_access_token, company_name):
    log.info(f"Configuring departure Out Of Office with delegate for [{user_email}], delegate: [{delegate_email}]")
    external_audience = "All"
    log.info(f"Setting to both internal and external audiences for [{user_email}]")
    log.info(f"Using company name: [{company_name}]")
    ps_command_body = f"""
      $message = "<html><body><p>Thank you for your email. Please note that I am no longer with {company_name}. For assistance, please contact {delegate_email}.</p></body></html>"
      Set-MailboxAutoReplyConfiguration -Identity '{user_email}' `
         -AutoReplyState Enabled `
         -InternalMessage $message `
         -ExternalMessage $message `
         -ExternalAudience {external_audience}
      $result = Get-MailboxAutoReplyConfiguration -Identity '{user_email}'
      Write-Output "SUCCESS: Set out of office for [{user_email}] with delegate to [{delegate_email}]"
    """
    return execute_out_of_office_command(log, azure_domain, user_email, exo_access_token, ps_command_body, "Out of Office with delegate")

def set_out_of_office_without_delegate(log, azure_domain, user_email, exo_access_token, company_name):
    log.info(f"Configuring departure Out Of Office without delegate for [{user_email}]")
    external_audience = "All"
    log.info(f"Setting to both internal and external audiences for [{user_email}]")
    log.info(f"Using company name: [{company_name}]")
    ps_command_body = f"""
      $message = "<html><body><p>Thank you for your email. Please note that I am no longer with {company_name}. This email account is no longer monitored. Please contact the appropriate department for assistance.</p></body></html>"
      Set-MailboxAutoReplyConfiguration -Identity '{user_email}' `
         -AutoReplyState Enabled `
         -InternalMessage $message `
         -ExternalMessage $message `
         -ExternalAudience {external_audience}
      $result = Get-MailboxAutoReplyConfiguration -Identity '{user_email}'
      Write-Output "SUCCESS: Set out of office without delegate for [{user_email}] without delegate"
    """
    return execute_out_of_office_command(log, azure_domain, user_email, exo_access_token, ps_command_body, "Out of Office without delegate")

def remove_out_of_office(log, azure_domain, user_email, exo_access_token):
    log.info(f"Removing Out Of Office for [{user_email}]")

    ps_command_body = f"""
      Write-Output "Removing out of office for: {user_email}"

      Set-MailboxAutoReplyConfiguration -Identity '{user_email}' `
         -AutoReplyState Disabled
            
      $result = Get-MailboxAutoReplyConfiguration -Identity '{user_email}'
      Write-Output "Auto-reply state: $($result.AutoReplyState)"
      Write-Output "SUCCESS: Out Of Office removed for [{user_email}]"
    """

    return execute_out_of_office_command(log, azure_domain, user_email, exo_access_token, ps_command_body, "remove Out of Office")

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1751347835114")
            operation = input.get_value("Operation_1752790836720")
            user_identifier = input.get_value("User_1751347754196")
            internal_message = input.get_value("InternalMessage_1751347755129")
            external_message = input.get_value("ExternalMessage_1751347756083")
            external_audience = input.get_value("ReplytoExternalEmails_1751347799107")
            delegate_email = input.get_value("DelegateEmail_1752792205382")
            auth_code = input.get_value("AuthCode_1751347837907")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        operation = operation.strip() if operation else "Set Out of Office - Manual"
        user_identifier = user_identifier.strip() if user_identifier else ""
        internal_message = internal_message.strip() if internal_message else ""
        external_message = external_message.strip() if external_message else internal_message
        external_audience = external_audience.strip() if external_audience else "All"
        delegate_email = delegate_email.strip() if delegate_email else ""
        auth_code = auth_code.strip() if auth_code else ""

        log.info(f"Ticket Number = [{ticket_number}]")
        log.info(f"Requested operation = [{operation}]")
        log.info(f"Received input user = [{user_identifier}]")
        log.info(f"Internal message = [{internal_message}]")
        log.info(f"External message = [{external_message}]")
        log.info(f"External audience = [{external_audience}]")
        log.info(f"Delegate email = [{delegate_email}]")

        if not user_identifier:
            record_result(log, ResultLevel.WARNING, "User identifier is empty or invalid")
            return
            
        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number is required but missing")
            return

        company_identifier, company_name, company_id, company_types = get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
        if not company_identifier:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve company identifier from ticket [{ticket_number}]")
            return

        if company_identifier == "MIT":
            if not validate_mit_authentication(log, http_client, vault_name, auth_code):
                return

        graph_tenant_id, graph_access_token = get_graph_token(log, http_client, vault_name, company_identifier)
        if not graph_access_token:
            record_result(log, ResultLevel.WARNING, "Failed to obtain MS Graph access token")
            return
            
        azure_domain = get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")
        if not azure_domain:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve Azure domain for company [{company_identifier}]")
            return

        exo_tenant_id, exo_access_token = get_exo_token(log, http_client, vault_name, company_identifier)
        if not exo_access_token:
            record_result(log, ResultLevel.WARNING, "Failed to obtain Exchange Online access token")
            return

        user_result = get_aad_user_data(log, http_client, msgraph_base_url, user_identifier, graph_access_token)
        if isinstance(user_result, list):
            details = "\n".join([f"- {u.get('displayName')} | {u.get('userPrincipalName')} | {u.get('id')}" for u in user_result])
            record_result(log, ResultLevel.WARNING, f"Multiple users found for [{user_identifier}]\n{details}")
            return

        user_id, user_email, sam_account_name, is_synced = user_result
        if not user_id:
            record_result(log, ResultLevel.WARNING, f"Failed to resolve user ID for [{user_identifier}]")
            return
        if not user_email:
            record_result(log, ResultLevel.WARNING, f"Unable to resolve user email for [{user_identifier}]")
            return

        success = False
        output = ""
        
        if operation == "Set Out of Office - Manual":
            success, output = set_out_of_office_manual(log, azure_domain, user_email, internal_message, external_message, external_audience, exo_access_token)
            
        elif operation == "Set Out of Office - With Delegate":
            if not delegate_email:
                record_result(log, ResultLevel.WARNING, "Delegate email is required for 'With Delegate' operation")
                return
            if not company_name:
                record_result(log, ResultLevel.WARNING, "Company name is required for 'With Delegate' operation")
                return
            success, output = set_out_of_office_with_delegate(log, azure_domain, user_email, delegate_email, exo_access_token, company_name)
            
        elif operation == "Set Out of Office - Without Delegate":
            if not company_name:
                record_result(log, ResultLevel.WARNING, "Company name is required for 'Without Delegate' operation")
                return
            success, output = set_out_of_office_without_delegate(log, azure_domain, user_email, exo_access_token, company_name)
            
        elif operation == "Remove Out of Office":
            success, output = remove_out_of_office(log, azure_domain, user_email, exo_access_token)

        else:
            record_result(log, ResultLevel.WARNING, f"Unknown operation: [{operation}]")
            return

        if not success:
            record_result(log, ResultLevel.WARNING, f"Failed to execute operation '{operation}': {output}")
            return

        if "SUCCESS:" in output:
            cleaned_output = output.replace("SUCCESS:", "").strip()
            record_result(log, ResultLevel.SUCCESS, cleaned_output)
        else:
            record_result(log, ResultLevel.WARNING, f"Unexpected or failed response: {output}")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()