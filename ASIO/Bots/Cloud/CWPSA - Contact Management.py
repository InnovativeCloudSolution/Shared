import sys
import random
import os
import time
import requests
import urllib.parse
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
bot_name = "MIT-CWPSA - Contact management"
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
   
def get_site_id(log, http_client, cwpsa_base_url, company_id, site_name):
    log.info(f"Fetching site ID for company ID: {company_id} and site name: {site_name}")
    
    if not site_name or not site_name.strip():
        log.error(f"Site name is empty or invalid for company ID: {company_id}")
        return None
    
    conditions = f"name='{site_name.strip()}'"
    encoded_conditions = urllib.parse.quote(conditions)
    endpoint = f"{cwpsa_base_url}/company/companies/{company_id}/sites?conditions={encoded_conditions}"
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    
    if response and response.status_code == 200:
        try:
            data = response.json()
            if data and len(data) > 0:
                site_id = data[0].get("id")
                log.info(f"Site ID for company ID {company_id} and site name {site_name} is {site_id}")
                return site_id
            else:
                log.warning(f"No sites found matching name '{site_name}' for company ID {company_id}")
        except Exception as e:
            log.error(f"Failed to parse JSON response from sites API: {str(e)}")
            return None
    
    log.warning(f"Direct site search failed. Trying to fetch all sites for company {company_id}")
    fallback_endpoint = f"{cwpsa_base_url}/company/companies/{company_id}/sites"
    fallback_response = execute_api_call(log, http_client, "get", fallback_endpoint, integration_name="cw_psa")
    
    if fallback_response and fallback_response.status_code == 200:
        try:
            sites = fallback_response.json()
            log.info(f"Found {len(sites)} sites for company {company_id}")
            
            if not sites:
                log.error(f"No sites found for company ID {company_id}")
                return None
            
            for site in sites:
                if site.get("name", "").strip() == site_name.strip():
                    site_id = site.get("id")
                    log.info(f"Found matching site ID {site_id} for site name: {site_name}")
                    return site_id
            
            log.warning(f"No exact match found for site name: {site_name}")
            log.info("Available sites:")
            for site in sites:
                log.info(f"  - {site.get('name', 'Unknown')} (ID: {site.get('id', 'Unknown')})")
            
            first_site = sites[0]
            first_site_id = first_site.get("id")
            first_site_name = first_site.get("name", "Unknown")
            log.warning(f"Using first available site as fallback: {first_site_name} (ID: {first_site_id})")
            return first_site_id
        except Exception as e:
            log.error(f"Failed to parse JSON response from fallback sites API: {str(e)}")
            return None
    else:
        log.error(f"Failed to fetch sites for company ID {company_id}")
    
    return None

def get_communication_type_id(log, http_client, cwpsa_base_url, type_name):
    log.info(f"Getting communication type ID for: {type_name}")
    endpoint = f"{cwpsa_base_url}/company/communicationTypes"
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    
    if response and response.status_code == 200:
        try:
            types = response.json()
            for comm_type in types:
                if comm_type.get("description", "").strip().lower() == type_name.strip().lower():
                    type_id = comm_type.get("id")
                    log.info(f"Found communication type ID {type_id} for '{type_name}'")
                    return type_id
            log.warning(f"Communication type '{type_name}' not found")
            return None
        except Exception as e:
            log.error(f"Failed to parse communication types response: {str(e)}")
            return None
    else:
        log.error(f"Failed to fetch communication types")
        return None

def get_contact_type_id(log, http_client, cwpsa_base_url, type_name):
    log.info(f"Getting contact type ID for: {type_name}")
    endpoint = f"{cwpsa_base_url}/company/contacts/types"
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    
    if response and response.status_code == 200:
        try:
            types = response.json()
            for contact_type in types:
                if contact_type.get("description", "").strip() == type_name.strip():
                    type_id = contact_type.get("id")
                    log.info(f"Found contact type ID {type_id} for '{type_name}'")
                    return type_id
            log.warning(f"Contact type '{type_name}' not found")
            return None
        except Exception as e:
            log.error(f"Failed to parse contact types response: {str(e)}")
            return None
    else:
        log.error(f"Failed to fetch contact types")
        return None

def get_contact(log, http_client, cwpsa_base_url, email_address, company_identifier):
    log.info(f"Searching for contact by email: {email_address} in company: {company_identifier}")
    conditions = f"childConditions=communicationItems/value='{email_address}' AND communicationItems/communicationType = 'Email'&conditions=company/identifier='{company_identifier}'"
    encoded_conditions = urllib.parse.quote(conditions)
    endpoint = f"{cwpsa_base_url}/company/contacts?{encoded_conditions}"
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    if not response:
        return None
    contacts = response.json()
    if contacts:
        contact = contacts[0]
        log.info(f"Found contact ID: {contact.get('id')} for email: {email_address} in company: {company_identifier}")
        return contact
    else:
        log.warning(f"No contact found with email address: {email_address} in company: {company_identifier}")
        return None

def get_contact_by_contact_details(log, http_client, cwpsa_base_url, contact_details, company_identifier):
    log.info(f"Searching for contact by provided contact details: {contact_details} in company: {company_identifier}")
    contact_id = contact_details.get("contact_id")
    email = contact_details.get("email")
    first_name = contact_details.get("first_name")
    last_name = contact_details.get("last_name")
    site_name = contact_details.get("site_name")
    direct_phone = contact_details.get("direct_phone")
    mobile_phone = contact_details.get("mobile_phone")
    types = contact_details.get("types")
    if contact_details.get("vip") == "Yes":
        vip = "true"
    else:
        vip = "false"
    # Build conditions string with all contact details
    # Check for conditions
    if contact_id:
        conditions = f"conditions=id={contact_id}"
    # Check for childConditions
    else:
        conditions = "childConditions="
        if email:
            conditions += f"communicationItems/value='{email}' AND communicationItems/communicationType = 'Email' AND "
        if first_name:
            conditions += f"firstName='{first_name}' AND "
        if last_name:
            conditions += f"lastName='{last_name}' AND "
        if site_name:
            conditions += f"site/name='{site_name}' AND "
        if direct_phone:
            conditions += f"communicationItems/value='{direct_phone}' AND communicationItems/communicationType = 'Direct' AND "
        if mobile_phone:
            conditions += f"communicationItems/value='{mobile_phone}' AND communicationItems/communicationType = 'Mobile' AND "
        if types:
            conditions += f"types/name='{types}' AND "
        if conditions.endswith(' AND '):
            conditions = conditions[:-5]
        # Check for customFieldsConditions
        if vip:
            conditions += f"&customFieldConditions=caption='VIP?' AND value='{vip}'"
        # Add company identifier suffix
        conditions += f"&conditions=company/identifier='{company_identifier}'"
    log.info(f"Conditions: {conditions}")
    encoded_conditions = urllib.parse.quote(conditions)
    log.info(f"Encoded conditions: {encoded_conditions}")
    endpoint = f"{cwpsa_base_url}/company/contacts?{encoded_conditions}"
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    if not response:
        return None
    contacts = response.json()
    if contacts:
        log.info(f"Found contacts with contact details: {contact_details} in company: {company_identifier}")
        return contacts
    else:
        log.warning(f"No contact found with contact details: {contact_details} in company: {company_identifier}")
        return None

def create_contact(log, http_client, cwpsa_base_url, first_name, last_name, email, direct_phone, mobile_phone, types, vip, company_id, site_id):
    log.info(f"Creating contact for company ID: {company_id}")
    endpoint = f"{cwpsa_base_url}/company/contacts"
    data = {
        "firstName": first_name,
        "lastName": last_name,
        "company": {"id": company_id},
        "site": {"id": site_id},
        "inactiveFlag": False,
        "disablePortalLoginFlag": True,
        "ignoreDuplicates": False
    }

    if direct_phone or email or mobile_phone:
        data["communicationItems"] = []
        if email:
            email_type_id = get_communication_type_id(log, http_client, cwpsa_base_url, "Email")
            if email_type_id:
                data["communicationItems"].append({"type": {"id": email_type_id}, "value": email})
        if direct_phone:
            phone_type_id = get_communication_type_id(log, http_client, cwpsa_base_url, "Direct")
            if phone_type_id:
                data["communicationItems"].append({"type": {"id": phone_type_id}, "value": direct_phone})
        if mobile_phone:
            mobile_type_id = get_communication_type_id(log, http_client, cwpsa_base_url, "Mobile")
            if mobile_type_id:
                data["communicationItems"].append({"type": {"id": mobile_type_id}, "value": mobile_phone})

    if types and types in ["End User", "Non-Support User"]:
        type_id = get_contact_type_id(log, http_client, cwpsa_base_url, types)
        if type_id:
            data["types"] = [{"id": type_id}]

    if vip == "Yes":
        data["customFields"] = [{"id": 1, "value": True}]

    log.info(f"Contact creation payload: {data}")
    response = execute_api_call(log, http_client, "post", endpoint, data=data, integration_name="cw_psa")
    if response and response.status_code == 201:
        log.info(f"Successfully created contact for company ID: {company_id}")
        return response.json()
    elif response:
        log.error(f"Failed to create contact for company ID: {company_id}, Status Code: {response.status_code}, Response: {response.text}")
        return None
    else:
        log.error(f"Failed to create contact for company ID: {company_id}, Status Code: No Response")
        return None

def update_contact(log, http_client, cwpsa_base_url, first_name, last_name, email, direct_phone, mobile_phone, types, vip, contact_id):
    log.info(f"Updating contact ID: {contact_id}")
    endpoint = f"{cwpsa_base_url}/company/contacts/{contact_id}"
    data = {}

    if first_name:
        data["firstName"] = first_name
    if last_name:
        data["lastName"] = last_name
    
    if email or direct_phone or mobile_phone:
        data["communicationItems"] = []
        if email:
            email_type_id = get_communication_type_id(log, http_client, cwpsa_base_url, "Email")
            if email_type_id:
                data["communicationItems"].append({"type": {"id": email_type_id}, "value": email})
        if direct_phone:
            phone_type_id = get_communication_type_id(log, http_client, cwpsa_base_url, "Direct")
            if phone_type_id:
                data["communicationItems"].append({"type": {"id": phone_type_id}, "value": direct_phone})
        if mobile_phone:
            mobile_type_id = get_communication_type_id(log, http_client, cwpsa_base_url, "Mobile")
            if mobile_type_id:
                data["communicationItems"].append({"type": {"id": mobile_type_id}, "value": mobile_phone})
    
    if types and types in ["End User", "Non-Support User"]:
        type_id = get_contact_type_id(log, http_client, cwpsa_base_url, types)
        if type_id:
            data["types"] = [{"id": type_id}]
    
    if vip == "Yes":
        data["customFields"] = [{"id": 1, "value": True}]

    response = execute_api_call(log, http_client, "put", endpoint, data=data, integration_name="cw_psa")
    if response and response.status_code == 200:
        log.info(f"Successfully updated contact ID: {contact_id}")
        return response.json()
    elif response:
        log.error(f"Failed to update contact ID: {contact_id}, Status Code: {response.status_code}, Response: {response.text}")
        return None
    else:
        log.error(f"Failed to update contact ID: {contact_id}, Status Code: No Response")
        return None
 
def disable_contact(log, http_client, cwpsa_base_url, contact_id):
    log.info(f"Disabling contact ID: {contact_id}")
    endpoint = f"{cwpsa_base_url}/company/contacts/{contact_id}"
    get_response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    
    if not get_response or get_response.status_code != 200:
        log.error(f"Cannot disable contact [{contact_id}] because contact retrieval failed")
        return False
    
    contact_data = get_response.json()
    first_name = contact_data.get("firstName", "").strip()
    last_name = contact_data.get("lastName", "").strip()

    if not first_name:
        log.error(f"Contact [{contact_id}] missing required field [firstName] to disable")
        return False

    data = {
        "firstName": first_name,
        "lastName": last_name,
        "inactiveFlag": True
    }

    response = execute_api_call(log, http_client, "put", endpoint, data=data, integration_name="cw_psa")
    if response:
        log.info(f"Successfully disabled contact ID: {contact_id}")
        return True
    return False

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1749703129853")
            operation = input.get_value("Operation_1750132764500")
            contact_id = input.get_value("ContactID_1750133950866")
            email = input.get_value("Email_1749702852898")
            first_name = input.get_value("FirstName_1749702828995")
            last_name = input.get_value("LastName_1749702831090")
            site_name = input.get_value("SiteName_1750132886946")
            direct_phone = input.get_value("DirectPhone_1749702888401")
            mobile_phone = input.get_value("MobilePhone_1749702903834")
            types = input.get_value("Types_1757372976943")
            vip = input.get_value("VIP_1750132852158")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        operation = operation.strip() if operation else ""
        contact_id = contact_id.strip() if contact_id else ""
        email = email.strip() if email else ""
        first_name = first_name.strip() if first_name else ""
        last_name = last_name.strip() if last_name else ""
        site_name = site_name.strip() if site_name else ""
        direct_phone = direct_phone.strip() if direct_phone else ""
        mobile_phone = mobile_phone.strip() if mobile_phone else ""
        types = types.strip() if types else ""
        vip = vip.strip() if vip else ""

        contact_details = {
            "contact_id": contact_id,
            "email": email,
            "first_name": first_name,
            "last_name": last_name,
            "site_name": site_name,
            "direct_phone": direct_phone,
            "mobile_phone": mobile_phone,
            "types": types,
            "vip": vip
        }

        if not first_name and not last_name and email:
            email_name = email.split("@")[0]
            if "." in email_name:
                name_parts = email_name.split(".")
                first_name = name_parts[0].capitalize()
                last_name = name_parts[-1].capitalize()
                log.info(f"Extracted names from email: firstName='{first_name}', lastName='{last_name}'")
            else:
                first_name = email_name.capitalize()
                last_name = "Contact"
                log.info(f"Used email prefix as first name: firstName='{first_name}', lastName='{last_name}'")
        elif not first_name and last_name:
            first_name = "Unknown"
            log.info(f"Using default first name: firstName='{first_name}'")
        elif first_name and not last_name:
            last_name = "Contact"
            log.info(f"Using default last name: lastName='{last_name}'")
        elif not first_name and not last_name:
            first_name = "New"
            last_name = "Contact"
            log.info(f"Using default names: firstName='{first_name}', lastName='{last_name}'")

        log.info(f"Ticket Number = [{ticket_number}]")
        log.info(f"Requested operation = [{operation}]")
        log.info(f"Received input email = [{email}]")

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number is required but missing")
            return
        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation value is missing or invalid")
            return

        log.info(f"Retrieving company data for ticket [{ticket_number}]")
        company_identifier, company_name, company_id, company_types = get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
        if not company_identifier:
            record_result(log, ResultLevel.WARNING, f"Failed to retrieve company identifier from ticket [{ticket_number}]")
            return

        if operation == "Get contact":
            contacts = get_contact_by_contact_details(log, http_client, cwpsa_base_url, contact_details, company_identifier)
            if contacts:
                log.info(contacts)
                if len(contacts) > 1:
                    record_result(log, ResultLevel.SUCCESS, f"Multiple contacts found with provided details in company: {company_identifier}")
                    for contact in contacts:
                        email = next((item['value'] for item in contact['communicationItems'] if item.get('type', {}).get('name') == 'Email'), None)
                        direct_phone = next((item['value'] for item in contact['communicationItems'] if item.get('type', {}).get('name') == 'Direct'), None)
                        mobile_phone = next((item['value'] for item in contact['communicationItems'] if item.get('type', {}).get('name') == 'Mobile'), None)
                        types = [type["name"] for type in contact["types"]]
                        vip_status = next((item['value'] for item in contact['customFields'] if item.get('caption') == "VIP"), None)
                        if vip_status == "false":
                            vip = "No"
                        else:
                            vip = "Yes"
                        record_result(log, ResultLevel.SUCCESS,f"Contact ID: [{contact.get('id')}], First Name: [{contact.get('firstName')}], Last Name: [{contact.get('lastName')}], Email: [{email}],Direct Phone: [{direct_phone}], Mobile Phone: [{mobile_phone}], Types: [{types}], VIP: [{vip}]")
                        data_to_log["contact_id"] = contacts[0].get("id")
                    return

                elif len(contacts) == 0:
                    record_result(log, ResultLevel.WARNING, f"No contact found with provided contact details: {contact_details} in company: {company_identifier}")
                    return

                contact = contacts[0]
                record_result(log, ResultLevel.SUCCESS, f"Contact found with provided details in company: {company_identifier}")
                email = next((item['value'] for item in contact['communicationItems'] if item.get('type', {}).get('name') == 'Email'), None)
                direct_phone = next((item['value'] for item in contact['communicationItems'] if item.get('type', {}).get('name') == 'Direct'), None)
                mobile_phone = next((item['value'] for item in contact['communicationItems'] if item.get('type', {}).get('name') == 'Mobile'), None)
                types = [type["name"] for type in contact["types"]]
                vip_status = next((item['value'] for item in contact['customFields'] if item.get('caption') == "VIP"), None)
                if vip_status == "false":
                    vip = "No"
                else:
                    vip = "Yes"
                record_result(log, ResultLevel.SUCCESS,f"Contact ID: [{contact.get('id')}], First Name: [{contact.get('firstName')}], Last Name: [{contact.get('lastName')}], Email: [{email}],Direct Phone: [{direct_phone}], Mobile Phone: [{mobile_phone}], Types: [{types}], VIP: [{vip}]")
                data_to_log["contact_id"] = contact.get("id")
            else:
                record_result(log, ResultLevel.WARNING, f"No contact found with contact details: {contact_details} in company: {company_identifier}")

        if operation == "Create new contact":
            if not site_name or not site_name.strip():
                record_result(log, ResultLevel.WARNING, f"Site name is required for creating a new contact but was not provided")
                return

            site_id = get_site_id(log, http_client, cwpsa_base_url, company_id, site_name)
            if not site_id:
                record_result(log, ResultLevel.WARNING, f"Failed to retrieve site ID for company ID [{company_id}] and site name [{site_name}]")
                return

            contact = create_contact(log, http_client, cwpsa_base_url, first_name, last_name, email, direct_phone, mobile_phone, types, vip, company_id, site_id)
            if contact:
                contact_id = contact.get("id")
                record_result(log, ResultLevel.SUCCESS, f"Contact created successfully for company ID [{company_id}], site ID [{site_id}], contact ID [{contact_id}]")
                data_to_log["contact_id"] = contact_id
            else:
                record_result(log, ResultLevel.WARNING, "Failed to create contact")
                data_to_log["company_id"] = company_id
                data_to_log["site_id"] = site_id

        elif operation == "Update existing contact":
            if not email and not contact_id:
                record_result(log, ResultLevel.WARNING, "Either email or contact ID is required to update an existing contact")
                return

            if contact_id:
                log.info(f"Using provided Contact ID: {contact_id}")
                existing_contact_id = contact_id
                contact_company_id = company_id
            else:
                contact = get_contact(log, http_client, cwpsa_base_url, email, company_identifier)
                if not contact:
                    record_result(log, ResultLevel.WARNING, f"No contact found with email [{email}]")
                    return
                existing_contact_id = contact.get("id")
                contact_company_id = contact.get("company", {}).get("id")
                if not contact_company_id:
                    record_result(log, ResultLevel.WARNING, f"Unable to determine company for contact with email [{email}]")
                    return

            site_id = get_site_id(log, http_client, cwpsa_base_url, contact_company_id, site_name)
            if not site_id:
                record_result(log, ResultLevel.WARNING, f"Failed to retrieve site ID for company ID [{contact_company_id}] and site name [{site_name}]")
                return

            if update_contact(log, http_client, cwpsa_base_url, first_name, last_name, email, direct_phone, mobile_phone, types, vip, existing_contact_id):
                record_result(log, ResultLevel.SUCCESS, f"Contact with ID [{existing_contact_id}] updated successfully")
                data_to_log["contact_id"] = existing_contact_id
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to update contact ID [{existing_contact_id}]")

        elif operation == "Disable existing contact":
            if not email and not contact_id:
                record_result(log, ResultLevel.WARNING, "Either email or contact ID is required to disable an existing contact")
                return

            if contact_id:
                log.info(f"Using provided Contact ID: {contact_id}")
                existing_contact_id = contact_id
            else:
                contact = get_contact(log, http_client, cwpsa_base_url, email, company_identifier)
                if not contact:
                    record_result(log, ResultLevel.WARNING, f"No contact found with email [{email}]")
                    return
                existing_contact_id = contact.get("id")
            if disable_contact(log, http_client, cwpsa_base_url, existing_contact_id):
                record_result(log, ResultLevel.SUCCESS, f"Contact with ID [{existing_contact_id}] disabled successfully")
                data_to_log["contact_id"] = existing_contact_id
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to disable contact with ID [{existing_contact_id}]")
                data_to_log["contact_id"] = existing_contact_id

        else:
            record_result(log, ResultLevel.WARNING, f"Unknown operation [{operation}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()