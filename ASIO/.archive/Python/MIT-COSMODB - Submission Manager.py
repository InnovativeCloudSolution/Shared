import sys
import json
import random
import os
import base64
import hmac
import hashlib
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
vault_name = "mit-azu1-prod1-akv1"

data_to_log = {}
bot_name = "MIT-COSMODB - Submission Manager"
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

def execute_cosmosdb_call(log, http_client, method, endpoint, data=None, headers=None):
    try:
        if not all([endpoint, headers]):
            log.error("Missing required Cosmos DB call parameters")
            return None

        response = execute_api_call(log, http_client, method, endpoint, data=data, headers=headers)

        if response and response.status_code in [200, 201]:
            return response

        log.error(
            f"Cosmos DB call failed Status: {response.status_code if response else 'N/A'}, Body: {response.text if response else 'N/A'}"
        )
        return None

    except Exception as e:
        log.exception(e, "Exception occurred during Cosmos DB call")
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
    if response and response.status_code == 200:
        secret_value = response.json().get("value", "")
        if secret_value:
            log.info(f"Successfully retrieved secret [{secret_name}]")
            return secret_value
    log.error(f"Failed to retrieve secret [{secret_name}] Status code: {response.status_code if response else 'N/A'}")
    return ""

def generate_cosmosdb_auth_header(verb, resource_type, resource_link, date_utc, master_key):
    key = base64.b64decode(master_key)
    text = f"{verb.lower()}\n{resource_type.lower()}\n{resource_link}\n{date_utc.lower()}\n\n"
    signature = base64.b64encode(hmac.new(key, text.encode("utf-8"), hashlib.sha256).digest()).decode()
    return f"type=master&ver=1.0&sig={urllib.parse.quote(signature)}"

def get_user_onboarding(log, http_client, ticket_number, endpoint_url, db_name, container_name, cosmos_key):
    try:
        log.info(
            f"Querying Cosmos DB for onboarding submission with ticket [{ticket_number}]")
        verb = "post"
        resource_type = "docs"
        resource_link = f"dbs/{db_name}/colls/{container_name}"
        date_utc = datetime.utcnow().strftime('%a, %d %b %Y %H:%M:%S GMT')

        auth_header = generate_cosmosdb_auth_header(
            verb, resource_type, resource_link, date_utc, cosmos_key)

        query = {
            "query": "SELECT * FROM c WHERE (c[\"cwpsa_ticket\"] = @ticket)",
            "parameters": [{"name": "@ticket", "value": int(ticket_number)}]
        }

        headers = {
            "Authorization": auth_header,
            "x-ms-version": "2017-02-22",
            "x-ms-date": date_utc,
            "x-ms-documentdb-isquery": "true",
            "x-ms-documentdb-query-enablecrosspartition": "true",
            "Content-Type": "application/query+json"
        }

        endpoint = f"{endpoint_url}/dbs/{db_name}/colls/{container_name}/docs"
        response = execute_cosmosdb_call(
            log, http_client, verb, endpoint, data=query, headers=headers)

        if response and response.status_code == 200:
            results = response.json().get("Documents", [])
            if results:
                log.info(
                    f"Found {len(results)} onboarding submissions for ticket [{ticket_number}]")
                return results
            else:
                log.info(
                    f"No onboarding submission found for ticket [{ticket_number}]")
                return None

        log.error(
            f"Failed to query Cosmos DB for ticket [{ticket_number}] Status: {response.status_code if response else 'N/A'}, Body: {response.text if response else 'N/A'}")
        return None

    except Exception as e:
        log.exception(
            e, f"Exception occurred while querying Cosmos DB for ticket [{ticket_number}]")
        return None

def update_user_onboarding(log, http_client, document, endpoint_url, db_name, container_name, cosmos_key):
    try:
        if "cwpsa_ticket" not in document or not document["cwpsa_ticket"]:
            log.error("Missing required partition key field [cwpsa_ticket]")
            return None

        verb = "post"
        resource_type = "docs"
        resource_link = f"dbs/{db_name}/colls/{container_name}"
        date_utc = datetime.utcnow().strftime('%a, %d %b %Y %H:%M:%S GMT')

        auth_header = generate_cosmosdb_auth_header(
            verb, resource_type, resource_link, date_utc, cosmos_key)

        headers = {
            "Authorization": auth_header,
            "x-ms-version": "2017-02-22",
            "x-ms-date": date_utc,
            "x-ms-documentdb-is-upsert": "true",
            "x-ms-documentdb-partitionkey": json.dumps([document["cwpsa_ticket"]]),
            "Content-Type": "application/json"
        }

        endpoint = f"{endpoint_url}/dbs/{db_name}/colls/{container_name}/docs"
        response = execute_cosmosdb_call(
            log, http_client, verb, endpoint, data=document, headers=headers)

        if response and response.status_code in [200, 201]:
            doc_id = response.json().get("id", "")
            log.info(
                f"Successfully updated document in Cosmos DB with ID [{doc_id}]")
            return doc_id

        log.error(
            f"Failed to update document Status: {response.status_code if response else 'N/A'}, Body: {response.text if response else 'N/A'}")
        return None

    except Exception as e:
        log.exception(
            e, "Exception occurred while updating document in Cosmos DB")
        return None

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1746962747623")
            operation = input.get_value("Operation_1746962782594")
            db_name = input.get_value("DatabaseName_1748299609003")
            container_name = input.get_value("ContainerName_1748299612528")

            status_preapproved = input.get_value("PreapprovedStatus_1753156237314")
            status_approved = input.get_value("ApprovedStatus_1753156285005")
            status_sspr = input.get_value("SSPRStatus_1753156325386")

            user_firstname = input.get_value("UserFirstName_1746962880202")
            user_lastname = input.get_value("UserLastName_1746963059372")
            user_fullname = input.get_value("UserFullName_1746963060740")
            user_username = input.get_value("UserUsername_1748301425847")
            user_upn = input.get_value("UserUPN_1748301417854")
            user_mailnickname = input.get_value("UserMailNickname_1748301415040")
            user_externalemailaddress = input.get_value("UserExternalEmailAddress_1750000000002")
            user_displayname = input.get_value("UserDisplayname_1748301420277")
            user_primary_smtp = input.get_value("user_primary_smtp")
            user_mobile_personal = input.get_value("UserMobilePersonal_1746963063489")
            user_employee_id = input.get_value("EmployeeID_1749523432590")
            user_employee_type = input.get_value("EmployeeType_1749523447422")
            password_reset_required = input.get_value("PasswordResetonLogin_1749523613580")
            user_startdate = input.get_value("UserStartDate_1746963062254")
            user_enddate = input.get_value("UserEndDate_1750000000001")
            user_fax = input.get_value("UserFax_1748301436378")

            organisation_company = input.get_value("OrganisationCompany_1749523229365")
            organisation_company_append = input.get_value("AppendCompanytoDisplayName_1749991647287")
            organisation_persona = input.get_value("OrganisationPersona_1746963067152")
            organisation_manager = input.get_value("OrganisationManager_1746963068726")
            organisation_manager_email = input.get_value("OrganisationManagerEmail_1750000000003")
            organisation_department = input.get_value("OrganisationDepartment_1746963072961")
            organisation_title = input.get_value("OrganisationTitle_1746963074520")
            organisation_site = input.get_value("OrganisationSite_1746963075817")
            organisation_site_office = input.get_value("OrganisationSiteOffice_1749525809096")
            organisation_site_streetaddress = input.get_value("OrganisationSiteStreetAddress_1746963078376")
            organisation_site_city = input.get_value("OrganisationSiteCity_1746963077096")
            organisation_site_state = input.get_value("OrganisationSiteState_1746963079827")
            organisation_site_zip = input.get_value("OrganisationSiteZip_1746963081496")
            organisation_site_country = input.get_value("OrganisationSiteCountry_1746963082819")

            clone_tag = input.get_value("CloneTag_1746963064762")
            clone_user = input.get_value("CloneUser_1746963065908")

            microsoft_domain = input.get_value("MicrosoftDomain_1746963084306")
            exchange_email_domain = input.get_value("EmailDomain_1749523578934")
            exchange_add_aliases = input.get_value("AllDomainsasAliases_1749523592265")
            exchange_sharedmailbox = input.get_value("ExchangeSharedMailbox_1746963093999")
            exchange_distributionlist = input.get_value("ExchangeDistributionList_1746963095611")
            exchange_mailboxdelegate = input.get_value("ExchangeMailboxDelegate_1750000000005")
            exchange_ooodelegate = input.get_value("ExchangeOOODelegate_1750000000006")
            exchange_forwardto = input.get_value("ExchangeForwardTo_1750000000007")
            group_license_sku = input.get_value("GroupLicenseSKU_1750000000004")
            group_license = input.get_value("GroupLicense_1746963086024")
            group_security = input.get_value("GroupSecurity_1746963091083")
            group_teams = input.get_value("GroupTeams_1746963087600")
            group_software = input.get_value("GroupSoftware_1746963092571")
            group_sharepoint = input.get_value("GroupSharePoint_1746963089514")

            hybrid_ou = input.get_value("HybridOU_1748301423288")
            hybrid_home_drive = input.get_value("HybridHomeDrive_1748301431243")
            hybrid_home_driveletter = input.get_value("HybridHomeDriveLetter_1748301433873")

            teams_delegateowner = input.get_value("TeamsDelegateOwner_1750000000008")

            mobile_required = input.get_value("MobileRequired_1746963665549")
            mobile_source = input.get_value("MobileSource_1746963667663")
            mobile_vendor = input.get_value("MobileVendor_1746963694692")
            mobile_tag = input.get_value("MobileTag_1746963697647")

            mobile_number_required = input.get_value("MobileNumberRequired_1746963700533")
            mobile_number_source = input.get_value("MobileNumberSource_1746963703735")
            mobile_number_tag = input.get_value("MobileNumberTag_1746963708306")

            endpoint_required = input.get_value("EndpointRequired_1746963710776")
            endpoint_source = input.get_value("EndpointSource_1746963714461")
            endpoint_vendor = input.get_value("EndpointVendor_1746963717382")
            endpoint_tag = input.get_value("EndpointTag_1746964068878")

            tablet_required = input.get_value("TabletRequired_1749523738508")
            tablet_source = input.get_value("TabletSource_1749523750228")
            tablet_vendor = input.get_value("TabletVendor_1749523762944")
            tablet_tag = input.get_value("TabletTag_1749523771667")

            asset_delegate = input.get_value("AssetDelegate_1750000000009")

            manual_task = input.get_value("ManualTask_1748301428460")

            notification_submitter = input.get_value("NotificationSubmitter_1750279268125")

        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = int(ticket_number.strip()) if ticket_number else ""
        operation = operation.strip() if operation else ""
        db_name = db_name.strip() if db_name else ""
        container_name = container_name.strip() if container_name else ""
        log.info(f"Operation: {operation}, Ticket Number: {ticket_number}, Database: {db_name}, Container: {container_name}")
        log.info(f"status_approved: {status_approved}, status_preapproved: {status_preapproved}, status_sspr: {status_sspr}")

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number is required")
            return

        endpoint_url = get_secret_value(log, http_client, vault_name, "MIT-AZU1-CosmosDB-Endpoint")
        cosmos_key = get_secret_value(log, http_client, vault_name, "MIT-AZU1-CosmosDB-Key")

        if not all([endpoint_url, db_name, container_name, cosmos_key]):
            record_result(log, ResultLevel.WARNING, "Missing Cosmos DB configuration values")
            return

        if operation == "Get Submission":
            log.info(f"Retrieving onboarding submission for ticket [{ticket_number}]")
            results = get_user_onboarding(log, http_client, ticket_number, endpoint_url, db_name, container_name, cosmos_key)
            results = results if results is not None else []
            output = results[0] if results else {}
            data_to_log.update(output)

            if results:
                log.info(output)
                record_result(log, ResultLevel.SUCCESS, f"Retrieved {len(results)} onboarding submission(s) for ticket [{ticket_number}]")
                log.result_data(output)
            else:
                record_result(log, ResultLevel.WARNING, f"No onboarding submission found for ticket [{ticket_number}]")

        elif operation == "Update Submission":
            log.info(f"Updating onboarding submission for ticket [{ticket_number}]")
            results = get_user_onboarding(log, http_client, ticket_number, endpoint_url, db_name, container_name, cosmos_key)
            results = results if results else []

            if not results:
                record_result(log, ResultLevel.WARNING, f"No existing onboarding submission found for ticket [{ticket_number}]")
                return

            existing = results[0]
            if not existing.get("id"):
                record_result(log, ResultLevel.WARNING, f"Existing submission does not contain required [id] field")
                return

            updated_document = existing.copy()

            if status_preapproved:
                updated_document["status_preapproved"] = status_preapproved
            if status_approved:
                updated_document["status_approved"] = status_approved
            if status_sspr:
                updated_document["status_sspr"] = status_sspr
                
            if user_firstname:
                updated_document["user_firstname"] = user_firstname
            if user_lastname:
                updated_document["user_lastname"] = user_lastname
            if user_fullname:
                updated_document["user_fullname"] = user_fullname
            if user_username:
                updated_document["user_username"] = user_username
            if user_upn:
                updated_document["user_upn"] = user_upn
            if user_mailnickname:
                updated_document["user_mailnickname"] = user_mailnickname
            if user_externalemailaddress:
                updated_document["user_externalemailaddress"] = user_externalemailaddress
            if user_displayname:
                updated_document["user_displayname"] = user_displayname
            if user_primary_smtp:
                updated_document["user_primary_smtp"] = user_primary_smtp
            if user_mobile_personal:
                updated_document["user_mobile_personal"] = user_mobile_personal
            if user_employee_id:
                updated_document["user_employee_id"] = user_employee_id
            if user_employee_type:
                updated_document["user_employee_type"] = user_employee_type
            if password_reset_required:
                updated_document["user_password_reset_required"] = password_reset_required
            if user_startdate:
                updated_document["user_startdate"] = user_startdate
            if user_enddate:
                updated_document["user_enddate"] = user_enddate
            if user_fax:
                updated_document["user_fax"] = user_fax
                
            if organisation_company:
                updated_document["organisation_company"] = organisation_company
            if organisation_company_append:
                updated_document["organisation_company_append"] = organisation_company_append
            if organisation_persona:
                updated_document["organisation_persona"] = organisation_persona
            if organisation_manager:
                updated_document["organisation_manager_name"] = organisation_manager
            if organisation_manager_email:
                updated_document["organisation_manager_email"] = organisation_manager_email
            if organisation_department:
                updated_document["organisation_department"] = organisation_department
            if organisation_title:
                updated_document["organisation_title"] = organisation_title
            if organisation_site:
                updated_document["organisation_site"] = organisation_site
            if organisation_site_office:
                updated_document["organisation_site_office"] = organisation_site_office
            if organisation_site_streetaddress:
                updated_document["organisation_site_streetaddress"] = organisation_site_streetaddress
            if organisation_site_city:
                updated_document["organisation_site_city"] = organisation_site_city
            if organisation_site_state:
                updated_document["organisation_site_state"] = organisation_site_state
            if organisation_site_zip:
                updated_document["organisation_site_zip"] = organisation_site_zip
            if organisation_site_country:
                updated_document["organisation_site_country"] = organisation_site_country
            
            if clone_tag:
                updated_document["clone_tag"] = clone_tag
            if clone_user:
                updated_document["clone_user"] = clone_user
                
            if microsoft_domain:
                updated_document["microsoft_domain"] = microsoft_domain
            if exchange_email_domain:
                updated_document["exchange_email_domain"] = exchange_email_domain
            if exchange_add_aliases:
                updated_document["exchange_add_aliases"] = exchange_add_aliases
            if exchange_sharedmailbox:
                updated_document["exchange_sharedmailbox"] = exchange_sharedmailbox
            if exchange_distributionlist:
                updated_document["exchange_distributionlist"] = exchange_distributionlist
            if exchange_mailboxdelegate:
                updated_document["exchange_mailboxdelegate"] = exchange_mailboxdelegate
            if exchange_ooodelegate:
                updated_document["exchange_ooodelegate"] = exchange_ooodelegate
            if exchange_forwardto:
                updated_document["exchange_forwardto"] = exchange_forwardto
                
            if group_license_sku:
                updated_document["group_license_sku"] = group_license_sku
            if group_license:
                updated_document["group_license"] = group_license
            if group_security:
                updated_document["group_security"] = group_security
            if group_teams:
                updated_document["group_teams"] = group_teams
            if group_software:
                updated_document["group_software"] = group_software
            if group_sharepoint:
                updated_document["group_sharepoint"] = group_sharepoint
                
            if hybrid_ou:
                updated_document["hybrid_ou"] = hybrid_ou
            if hybrid_home_drive:
                updated_document["hybrid_home_drive"] = hybrid_home_drive
            if hybrid_home_driveletter:
                updated_document["hybrid_home_driveletter"] = hybrid_home_driveletter
                
            if teams_delegateowner:
                updated_document["teams_delegateowner"] = teams_delegateowner
                
            if mobile_required:
                updated_document["mobile_required"] = mobile_required
            if mobile_source:
                updated_document["mobile_source"] = mobile_source
            if mobile_vendor:
                updated_document["mobile_vendor"] = mobile_vendor
            if mobile_tag:
                updated_document["mobile_tag"] = mobile_tag
                
            if mobile_number_required:
                updated_document["mobile_number_required"] = mobile_number_required
            if mobile_number_source:
                updated_document["mobile_number_source"] = mobile_number_source
            if mobile_number_tag:
                updated_document["mobile_number_tag"] = mobile_number_tag
                
            if endpoint_required:
                updated_document["endpoint_required"] = endpoint_required
            if endpoint_source:
                updated_document["endpoint_source"] = endpoint_source
            if endpoint_vendor:
                updated_document["endpoint_vendor"] = endpoint_vendor
            if endpoint_tag:
                updated_document["endpoint_tag"] = endpoint_tag
                
            if tablet_required:
                updated_document["tablet_required"] = tablet_required
            if tablet_source:
                updated_document["tablet_source"] = tablet_source
            if tablet_vendor:
                updated_document["tablet_vendor"] = tablet_vendor
            if tablet_tag:
                updated_document["tablet_tag"] = tablet_tag
                
            if asset_delegate:
                updated_document["asset_delegate"] = asset_delegate
                
            if manual_task:
                updated_document["manual_task"] = manual_task
                
            if notification_submitter:
                updated_document["notification_submitter"] = notification_submitter

            log.info(f"Updating onboarding submission: {json.dumps(updated_document, indent=2)}")
            doc_id = update_user_onboarding(log, http_client, updated_document, endpoint_url, db_name, container_name, cosmos_key)
            data_to_log.update(updated_document)

            if doc_id:
                record_result(log, ResultLevel.SUCCESS, f"Updated onboarding submission with ID [{doc_id}]")
            else:
                record_result(log, ResultLevel.WARNING, "Failed to update onboarding submission")

        else:
            record_result(log, ResultLevel.WARNING, f"Unknown operation [{operation}]. Supported operations: get submission, update submission")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()