import sys
import random
import os
import time
import requests
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

cwpsa_base_url = "https://au.myconnectwise.net/v4_6_release/apis/3.0"
vault_name = "mit-azu1-prod1-akv1"
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
        log.error(
            f"Failed to retrieve ticket [{ticket_number}] "
            f"Status: {ticket_response.status_code}, Body: {ticket_response.text}"
        )
    else:
        log.error(f"No response received when retrieving ticket [{ticket_number}]")

    return "", "", 0, ""
   
def get_site_id(log, http_client, cwpsa_base_url, company_id, site_name):
    log.info(f"Fetching site ID for company ID: {company_id} and site name: {site_name}")
    endpoint = f"{cwpsa_base_url}/company/companies/{company_id}/sites?conditions=name='{site_name}'"
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    if response and response.status_code == 200:
        data = response.json()
        if data:
            site_id = data[0].get("id")
            log.info(f"Site ID for company ID {company_id} and site name {site_name} is {site_id}")
            return site_id
        else:
            log.warning(f"No site found for company ID {company_id} with name: {site_name}")
    else:
        log.error(f"Failed to fetch site ID for company ID {company_id}, Status Code: {response.status_code if response else 'No Response'}")
    return None

def get_contact(log, http_client, cwpsa_base_url, contact_id):
    log.info(f"Retrieving contact ID: {contact_id}")
    endpoint = f"{cwpsa_base_url}/company/contacts/{contact_id}"
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")

    if response and response.status_code == 200:
        log.info(f"Successfully retrieved contact ID: {contact_id}")
        return response.json()
    else:
        log.error(f"Failed to retrieve contact ID: {contact_id}, Status Code: {response.status_code if response else 'No Response'}")
        return None

def create_contact(log, http_client, cwpsa_base_url, input_data, company_id, site_id):
    log.info(f"Creating contact for company ID: {company_id}")
    endpoint = f"{cwpsa_base_url}/company/companies/{company_id}/contacts"
    data = {
        "firstName": input_data.get("first_name", ""),  
        "lastName": input_data.get("last_name", ""), 
        "company": {"id": company_id},
        "site": {"id": site_id},
        "inactiveFlag": False,
        "title": input_data.get("job_title", ""),
        "disablePortalLoginFlag": True,
        "communicationItems": [],
        "types": input_data.get("types", "").split(",") if input_data.get("types") else [],
        "customFields": [
            {
                "id": 1,
                "value": input_data.get("vip", "")
            }
        ],
        "ignoreDuplicates": False
    }

    if input_data.get("direct_phone"):
        data["communicationItems"].append({"type": {"id": 2}, "value": input_data["direct_phone"]})
    if input_data.get("email"):
        data["communicationItems"].append({"type": {"id": 1}, "value": input_data["email"]})
    if input_data.get("mobile_phone"):
        data["communicationItems"].append({"type": {"id": 4}, "value": input_data["mobile_phone"]})

    response = execute_api_call(log, http_client, "post", endpoint, data=data, integration_name="cw_psa")
    if response and response.status_code == 201:
        log.info(f"Successfully created contact for company ID: {company_id}")
        return response.json()
    else:
        log.error(f"Failed to create contact for company ID: {company_id}, Status Code: {response.status_code if response else 'No Response'}")
        return None

def update_contact(log, http_client, cwpsa_base_url, contact_id, input_data):
    log.info(f"Updating contact ID: {contact_id}")
    endpoint = f"{cwpsa_base_url}/company/contacts/{contact_id}"
    data = {"communicationItems": []}

    if input_data.get("first_name"):
        data["firstName"] = input_data["first_name"]
    if input_data.get("last_name"):
        data["lastName"] = input_data["last_name"]
    if input_data.get("email"):
        data["communicationItems"].append({"type": {"id": 1}, "value": input_data["email"]})
    if input_data.get("direct_phone"):
        data["communicationItems"].append({"type": {"id": 2}, "value": input_data["direct_phone"]})
    if input_data.get("mobile_phone"):
        data["communicationItems"].append({"type": {"id": 4}, "value": input_data["mobile_phone"]})
    if input_data.get("types"):
        data["types"] = input_data["types"].split(",") if isinstance(input_data["types"], str) else input_data["types"]
    if input_data.get("vip"):
        data["customFields"] = [{"id": 1, "value": input_data["vip"]}]

    response = execute_api_call(log, http_client, "put", endpoint, data=data, integration_name="cw_psa")
    if response and response.status_code == 200:
        log.info(f"Successfully updated contact ID: {contact_id}")
        return response.json()
    else:
        log.error(f"Failed to update contact ID: {contact_id}, Status Code: {response.status_code if response else 'No Response'}")
        return None
 
def disable_contact(log, http_client, cwpsa_base_url, contact_id):
    log.info(f"Disabling contact ID: {contact_id}")

    contact_data = get_contact(log, http_client, cwpsa_base_url, contact_id)
    if not contact_data:
        log.error(f"Cannot disable contact [{contact_id}] because contact retrieval failed")
        return False

    first_name = contact_data.get("firstName", "").strip()
    last_name = contact_data.get("lastName", "").strip()

    if not first_name:
        log.error(f"Contact [{contact_id}] missing required field [firstName] to disable")
        return False

    endpoint = f"{cwpsa_base_url}/company/contacts/{contact_id}"
    data = {
        "firstName": first_name,
        "lastName": last_name,
        "inactiveFlag": True
    }

    response = execute_api_call(log, http_client, "put", endpoint, data=data, integration_name="cw_psa")
    if response and response.status_code == 200:
        log.info(f"Successfully disabled contact ID: {contact_id}")
        return True
    else:
        log.error(f"Failed to disable contact ID: {contact_id}, Status Code: {response.status_code if response else 'No Response'}")
        return False

def main():
    try:
        try:
            operation = input.get_value("Operation_1750132764500")
            contact_id = input.get_value("ContactID_1750133950866")
            first_name = input.get_value("FirstName_1749702828995")
            last_name = input.get_value("LastName_1749702831090")
            email = input.get_value("Email_1749702852898")
            site_name = input.get_value("SiteName_1750132886946")
            direct_phone = input.get_value("DirectPhone_1749702888401")
            mobile_phone = input.get_value("MobilePhone_1749702903834")
            types = input.get_value("Types_1749702948907")
            vip = input.get_value("VIP_1750132852158")
            job_title = input.get_value("JobTitle_1750132852159")
            ticket_number = input.get_value("TicketNumber_1749703129853")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        operation = operation.strip() if operation else ""
        contact_id = contact_id.strip() if contact_id else ""
        first_name = first_name.strip() if first_name else ""
        last_name = last_name.strip() if last_name else ""
        email = email.strip() if email else ""
        site_name = site_name.strip() if site_name else ""
        direct_phone = direct_phone.strip() if direct_phone else ""
        mobile_phone = mobile_phone.strip() if mobile_phone else ""
        types = types.strip() if types else ""
        vip = vip.strip() if vip else ""
        job_title = job_title.strip() if job_title else ""
        ticket_number = ticket_number.strip() if ticket_number else ""

        log.info(f"Operation: {operation}, Ticket: {ticket_number}, Contact ID: {contact_id}, First Name: {first_name}, Last Name: {last_name}, Email: {email}, Site: {site_name}, Direct Phone: {direct_phone}, Mobile Phone: {mobile_phone}, Types: {types}, VIP: {vip}")

        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation name is required")
            return

        if operation == "Create new contact":
            if not ticket_number:
                record_result(log, ResultLevel.WARNING, "Ticket number is required to create a new contact")
                return

            company_identifier, company_name, company_id, company_type = get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
            if not company_identifier or not company_id:
                record_result(log, ResultLevel.WARNING, f"Failed to retrieve company information from ticket [{ticket_number}]")
                return

            site_id = get_site_id(log, http_client, cwpsa_base_url, company_id, site_name)
            if not site_id:
                record_result(log, ResultLevel.WARNING, f"Failed to retrieve site ID for company ID [{company_id}] and site name [{site_name}]")
                return

            input_data = {
                "first_name": first_name,
                "last_name": last_name,
                "email": email,
                "site_name": site_name,
                "direct_phone": direct_phone,
                "mobile_phone": mobile_phone,
                "types": types,
                "vip": vip,
                "job_title": job_title
            }

            contact = create_contact(log, http_client, cwpsa_base_url, input_data, company_id, site_id)
            if contact:
                record_result(log, ResultLevel.SUCCESS, f"Contact created successfully for company ID [{company_id}] and site ID [{site_id}]")
                data_to_log["ContactID"] = contact.get("id")
            else:
                record_result(log, ResultLevel.WARNING, "Failed to create contact")
                data_to_log["CompanyID"] = company_id
                data_to_log["SiteID"] = site_id

        elif operation == "Update existing contact":
            if not contact_id:
                record_result(log, ResultLevel.WARNING, "Contact ID is required to update an existing contact")
                return
            if not ticket_number:
                record_result(log, ResultLevel.WARNING, "Ticket number is required to update an existing contact")
                return

            company_identifier, company_name, company_id, company_type = get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
            if not company_identifier or not company_id:
                record_result(log, ResultLevel.WARNING, f"Failed to retrieve company information from ticket [{ticket_number}]")
                return

            site_id = get_site_id(log, http_client, cwpsa_base_url, company_id, site_name)
            if not site_id:
                record_result(log, ResultLevel.WARNING, f"Failed to retrieve site ID for company ID [{company_id}] and site name [{site_name}]")
                return

            input_data = {
                "first_name": first_name,
                "last_name": last_name,
                "email": email,
                "site_name": site_name,
                "direct_phone": direct_phone,
                "mobile_phone": mobile_phone,
                "types": types,
                "vip": vip,
                "job_title": job_title
            }

            if update_contact(log, http_client, cwpsa_base_url, contact_id, input_data):
                record_result(log, ResultLevel.SUCCESS, f"Contact with ID [{contact_id}] updated successfully")
                data_to_log["ContactID"] = contact_id
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to update contact ID [{contact_id}]")

        elif operation == "Disable existing contact":
            if not contact_id:
                record_result(log, ResultLevel.WARNING, "Contact ID is required to disable an existing contact")
                return

            if disable_contact(log, http_client, cwpsa_base_url, contact_id):
                record_result(log, ResultLevel.SUCCESS, f"Contact with ID [{contact_id}] disabled successfully")
                data_to_log["ContactID"] = contact_id
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to disable contact with ID [{contact_id}]")
                data_to_log["ContactID"] = contact_id

        else:
            record_result(log, ResultLevel.WARNING, f"Unknown operation [{operation}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()