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
log.info("Static variables set")

def record_result(log, level, message):
    log.result_message(level, message)
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
                    if headers and headers.get("Content-Type") == "application/x-www-form-urlencoded":
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
        company_type = ""
        if company_response and company_response.status_code == 200:
            company_data = company_response.json()
            company_type = company_data.get("type", {}).get("name", "")
            log.info(f"Company type for ID [{company_id}]: [{company_type}]")
        else:
            log.warning(f"Unable to retrieve company type for ID [{company_id}]")
        return company_identifier, company_name, company_id, company_type
    elif ticket_response:
        log.error(f"Failed to retrieve ticket [{ticket_number}] Status: {ticket_response.status_code}, Body: {ticket_response.text}")
    else:
        log.error(f"No response received when retrieving ticket [{ticket_number}]")
    return "", "", 0, ""

def get_ticket_data(log, http_client, cwpsa_base_url, ticket_number):
    try:
        log.info(f"Retrieving full ticket details for ticket number [{ticket_number}]")
        endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}"
        response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
        if not response or response.status_code != 200:
            log.error(f"Failed to retrieve ticket [{ticket_number}] - Status: {response.status_code if response else 'N/A'}")
            return "", "", "", ""
        ticket = response.json()
        ticket_summary = ticket.get("summary", "")
        ticket_type = ticket.get("type", {}).get("name", "")
        priority_name = ticket.get("priority", {}).get("name", "")
        due_date = ticket.get("requiredDate", "")
        log.info(f"Ticket [{ticket_number}] Summary = [{ticket_summary}], Type = [{ticket_type}], Priority = [{priority_name}], Due = [{due_date}]")
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
        log.result_message(ResultLevel.FAILED, "Failed to retrieve expected authentication code for MIT")
        return False
    if auth_code.strip() != expected_code.strip():
        log.result_message(ResultLevel.FAILED, "Provided authentication code is incorrect")
        return False
    return True

def get_aad_user_data(log, http_client, msgraph_base_url, user_identifier, token):
    log.info(f"Resolving user ID and email for [{user_identifier}]")
    headers = {"Authorization": f"Bearer {token}"}
    if re.fullmatch(r"[0-9a-fA-F\-]{36}", user_identifier):
        endpoint = f"{msgraph_base_url}/users/{user_identifier}"
        response = execute_api_call(log, http_client, "get", endpoint, headers=headers)
        if response and response.status_code == 200:
            user = response.json()
            return user.get("id", ""), user.get("userPrincipalName", ""), user.get("onPremisesSamAccountName", "")
        return "", "", ""
    filters = [
        f"startswith(displayName,'{user_identifier}')",
        f"startswith(userPrincipalName,'{user_identifier}')",
        f"startswith(mail,'{user_identifier}')"
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
            return user.get("id", ""), user.get("userPrincipalName", ""), user.get("onPremisesSamAccountName", "")
    log.error(f"Failed to resolve user ID and email for [{user_identifier}]")
    return "", "", ""

def append_proxyaddresses(log, user_identifier, azure_domain, token):
    log.info(f"Appending proxy addresses for user [{user_identifier}]")
    if not user_identifier:
        log.error("User identifier is empty or invalid")
        return "", ""

    ps_command = f"""
    $ErrorActionPreference = 'Stop'

    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -AccessToken '{token}' -Organization '{azure_domain}' -ShowBanner:$false

    try {{
        $domains = Get-AcceptedDomain | Where-Object {{ $_.DomainType -eq 'Authoritative' }} | Select-Object -ExpandProperty DomainName
        $user = "{user_identifier}"
        $userObj = Get-Recipient -Identity $user
        $proxyAddresses = @($userObj.EmailAddresses)
        foreach ($domain in $domains) {{
            $newProxy = "smtp:$($user.Split('@')[0])@$domain"
            if ($proxyAddresses -notcontains $newProxy) {{
                $proxyAddresses += $newProxy
            }}
        }}
        $proxyAddresses = $proxyAddresses | Where-Object {{ $_ -notlike "*codetwo*" }}
        Set-Mailbox -Identity $user -EmailAddresses $proxyAddresses
        Write-Output "success : $proxyAddresses"
    }} catch {{
        Write-Output "$proxyAddresses"
        Write-Error "Error: $_"
    }}

    Disconnect-ExchangeOnline -Confirm:$false
    """

    success, output = execute_powershell(log, ps_command, shell="pwsh")
    if not success or not output or "Error:" in output:
        log.error(f"Failed to resolve user UPN and SAMAccountName for [{user_identifier}]")
        return "", ""

    log.info(f"Raw PowerShell output:\n{output}")
    return success, output.strip()

def main():
    try:
        try:
            user_identifier = input.get_value("User_1750834613938")
            ticket_number = input.get_value("TicketNumber_1750834615225")
            auth_code = input.get_value("AuthenticationCode_1750834616194")
            provided_token = input.get_value("AccessToken_1750834656991")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        user_identifier = user_identifier.strip() if user_identifier else ""
        ticket_number = ticket_number.strip() if ticket_number else ""
        auth_code = auth_code.strip() if auth_code else ""
        provided_token = provided_token.strip() if provided_token else ""

        log.info(f"Received input user = [{user_identifier}]")
        if not user_identifier:
            record_result(log, ResultLevel.WARNING, "User identifier is empty or invalid")
            return

        access_token = ""
        exo_access_token = ""
        company_identifier = ""
        user_id = user_email = user_sam = ""

        if provided_token:
            access_token = provided_token
            log.info("Using provided access token")
            if not isinstance(access_token, str) or "." not in access_token:
                record_result(log, ResultLevel.WARNING, "Provided access token is malformed (missing dots)")
                return
        elif ticket_number:
            log.info(f"Retrieving company data for ticket [{ticket_number}]")
            company_identifier, company_name, company_id, company_type = get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
            if not company_identifier:
                record_result(log, ResultLevel.WARNING, f"Failed to retrieve company identifier from ticket [{ticket_number}]")
                return

            if company_identifier == "MIT":
                if not validate_mit_authentication(log, http_client, vault_name, auth_code):
                    return

            client_id = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientID")
            client_secret = get_secret_value(log, http_client, vault_name, "MIT-PartnerApp-ClientSecret")
            azure_domain = get_secret_value(log, http_client, vault_name, f"{company_identifier}-PrimaryDomain")
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

            exo_client_id = get_secret_value(log, http_client, vault_name, f"{company_identifier}-ExchangeApp-ClientID")
            exo_client_secret = get_secret_value(log, http_client, vault_name, f"{company_identifier}-ExchangeApp-ClientSecret")
            if not all([exo_client_id, exo_client_secret, azure_domain]):
                record_result(log, ResultLevel.WARNING, "Failed to retrieve required internal secrets")
                return

            exo_access_token = get_access_token(log, http_client, tenant_id, exo_client_id, exo_client_secret, scope="https://outlook.office365.com/.default")
            if not isinstance(exo_access_token, str) or "." not in exo_access_token:
                record_result(log, ResultLevel.WARNING, "Exchange Online access token is malformed (missing dots)")
                return

        aad_user_result = get_aad_user_data(log, http_client, msgraph_base_url, user_identifier, access_token)
        if isinstance(aad_user_result, list):
            details = "\n".join([f"- {u.get('displayName')} | {u.get('userPrincipalName')} | {u.get('id')}" for u in aad_user_result])
            record_result(log, ResultLevel.WARNING, f"Multiple users found for [{user_identifier}]\n{details}")
            return

        user_id, user_email, user_sam = aad_user_result
        if not user_id:
            record_result(log, ResultLevel.WARNING, f"Failed to resolve user ID for [{user_identifier}]")
            return
        if not user_email:
            record_result(log, ResultLevel.WARNING, f"Unable to resolve user email for [{user_identifier}]")
            return

        success, output = append_proxyaddresses(log, user_email, azure_domain, exo_access_token)
        if not success:
            record_result(log, ResultLevel.WARNING, f"Failed to append proxy addresses for user [{user_email}]")
            return
        record_result(log, ResultLevel.SUCCESS, f"Successfully appended proxy addresses for user [{user_email}]")
        record_result(log, ResultLevel.INFO, output)

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()