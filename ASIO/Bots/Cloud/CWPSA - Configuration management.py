from sqlite3 import SQLITE_LIMIT_VARIABLE_NUMBER
import sys
import random
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
vault_name = "mit-azu1-prod1-akv1"

data_to_log = {}
bot_name = "MIT-CWPSA - Configuration management"
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
    return "", "", 0, ""

def get_ticket_data(log, http_client, cwpsa_base_url, ticket_number):
    try:
        log.info(f"Retrieving full ticket details for ticket number [{ticket_number}]")
        endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}"
        response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
        if not response:
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

def get_contact(log, http_client, cwpsa_base_url, user_identifier, company_identifier):
    log.info(f"Searching for contact by user identifier: {user_identifier} in company: {company_identifier}")
    conditions = f'company/identifier="{company_identifier}"'
    child_conditions = f'communicationItems/value="{user_identifier}"'
    encoded_conditions = urllib.parse.quote(conditions)
    encoded_child_conditions = urllib.parse.quote(child_conditions)
    endpoint = f"{cwpsa_base_url}/company/contacts?conditions={encoded_conditions}&childConditions={encoded_child_conditions}"
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    if not response:
        return None
    contacts = response.json()
    if contacts:
        contact = contacts[0]
        log.info(f"Found contact ID: {contact.get('id')} for user identifier: {user_identifier} in company: {company_identifier}")
        return contact
    else:
        log.warning(f"No contact found with user identifier: {user_identifier} in company: {company_identifier}")
        return None

def build_query_string(configuration_data, company_identifier):
    """Builds URL-encoded conditions and childConditions query parameters."""
    conditions_matrix = [
        {"id": "condition_id", "condition": "id", "operator": "=", "type": "condition"},
        {"id": "configuration_name", "condition": "name", "operator": "like", "type": "condition"},
        {"id": "type", "condition": "type/name", "operator": "like", "type": "condition"},
        {"id": "contact_id", "condition": "contact/id", "operator": "=", "type": "condition"},
        {"id": "site_id", "condition": "site/name", "operator": "like", "type": "condition"},
        {"id": "serial_number", "condition": "serialNumber", "operator": "like", "type": "condition"},
        {"id": "tag_number", "condition": "tagNumber", "operator": "like", "type": "condition"},
        {"id": "model_number", "condition": "modelNumber", "operator": "like", "type": "condition"},
        {"id": "last_login_name", "condition": "lastLoginName", "operator": "like", "type": "condition"},
        {"id": "os_type", "condition": "osType", "operator": "=", "type": "condition"},
        {"id": "active", "condition": "status/name", "operator": "=", "type": "condition"},
    ]
    conditions = []
    child_conditions = []
    for item in conditions_matrix:
        value = configuration_data.get(item["id"])
        if value is None:
            continue
        operator = " LIKE " if item["operator"] == "like" else item["operator"]
        condition_str = f"{item['condition']}{operator}'{value}'"
        if item["type"] == "condition" and configuration_data.get(item["id"]) not in [None, ""]:
            conditions.append(condition_str)
        elif item["type"] == "childCondition" and configuration_data.get(item["id"]) not in [None, ""]:
            child_conditions.append(condition_str)
    conditions.append(f"company/identifier='{company_identifier}'")
    query_parts = []
    if conditions:
        conditions_str = " AND ".join(conditions)
        query_parts.append(f"conditions={urllib.parse.quote(conditions_str)}")
    if child_conditions:
        child_conditions_str = " AND ".join(child_conditions)
        query_parts.append(f"childConditions={urllib.parse.quote(child_conditions_str)}")
    return "&".join(query_parts)

def get_configurations_by_criteria(log, http_client, cwpsa_base_url, company_identifier, configuration_data):
    log.info(f"Searching for configurations by provided details: {configuration_data} in company: {company_identifier}")
    query_string = build_query_string(configuration_data, company_identifier)
    log.info(f"Query string: {query_string}")
    endpoint = f"{cwpsa_base_url}/company/configurations?{query_string}"
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    if not response:
        return None
    configs = response.json()
    if configs:
        log.info(f"Found configurations with details: {configuration_data} in company: {company_identifier}")
        return configs
    else:
        log.warning(f"No configurations found with details: {configuration_data} in company: {company_identifier}")
        return None

def get_configuration_data(log, http_client, cwpsa_base_url, company_id, flag_caption):
    conditions = f"company/id={company_id}"
    custom_conditions = f'caption="{flag_caption}" AND value=true'
    query_string = f"conditions={urllib.parse.quote(conditions)}&customFieldConditions={urllib.parse.quote(custom_conditions)}"
    url = f"{cwpsa_base_url}/company/configurations?{query_string}"
    response = execute_api_call(log, http_client, "get", url, integration_name="cw_psa")
    if response:
        try:
            configs = response.json()
        except Exception:
            log.warning(f"Failed to decode configuration response JSON, or no configurations found for flag caption: [{flag_caption}]")
            return None
        for config in configs:
            custom_fields = config.get("customFields", [])
            has_flag = any(
                f.get("caption") == flag_caption and f.get("value") is True
                for f in custom_fields
            )
            if has_flag:
                endpoint_id = next(
                    (f.get("value") for f in custom_fields if f.get("caption") == "Endpoint ID"),
                    ""
                )
                return {
                    "name": config.get("name", ""),
                    "id": config.get("id"),
                    "endpoint_id": endpoint_id
                }
    log.warning(f"No configuration found for flag caption: [{flag_caption}]")
    return None

def get_user_configuration_items(log, http_client, cwpsa_base_url, company_id, contact_id):
    log.info(f"Fetching configurations for company ID [{company_id}] and contact ID [{contact_id}]")
    conditions = f"company/id={company_id} and contact/id={contact_id}"
    query_string = f"conditions={urllib.parse.quote(conditions)}"
    url = f"{cwpsa_base_url}/company/configurations?{query_string}"

    response = execute_api_call(log, http_client, "get", url, integration_name="cw_psa")
    if response:
        try:
            items = response.json()
        except Exception:
            log.error("Failed to decode configuration response JSON")
            return "", ""
        formatted_line = []
        formatted_note = []
        for item in items:
            fields = [
                item.get("name", ""),
                item.get("tagNumber", ""),
                item.get("serialNumber", ""),
                item.get("manufacturer", {}).get("name", ""),
                item.get("modelNumber", "")
            ]
            non_blank_fields = [f"[{f}]" for f in fields if f]
            entry = " - ".join(non_blank_fields)
            formatted_line.append(entry)
            formatted_note.append(entry)
        return ", ".join(formatted_line), "\n".join(formatted_note)
    log.warning(f"No configuration items found for contact ID [{contact_id}] in company [{company_id}]")
    return "", ""

def get_configuration_question(log, http_client, cwpsa_base_url, ticket_number, company_identifier):
    log.info(f"Searching for configuration with name [{ticket_number}] and type 'Automation - Submission' in company [{company_identifier}]")
    
    conditions = f'name="{ticket_number}" AND type/name="Automation - Submission" AND company/identifier="{company_identifier}"'
    query_string = f"conditions={urllib.parse.quote(conditions)}"
    url = f"{cwpsa_base_url}/company/configurations?{query_string}"
    
    response = execute_api_call(log, http_client, "get", url, integration_name="cw_psa")
    if not response:
        return {}
    
    configs = response.json()
    if not configs:
        log.warning(f"No configuration found with name [{ticket_number}] and type 'Automation - Submission'")
        return {}
    
    config = configs[0]
    config_id = config.get("id")
    log.info(f"Found configuration: {config.get('name', '')} (ID: {config_id})")
    
    questions = config.get("questions", [])
    if not questions:
        log.info("No questions found in configuration")
        return {"id": config_id, "questions": {}}
    
    submission_data = {}
    for question in questions:
        question_text = question.get("question", "")
        answer_value = question.get("answer")
        question_id = question.get("questionId")
        
        if answer_value is not None and str(answer_value).strip():
            submission_data[question_text] = {"answer": answer_value, "questionId": question_id}
            log.info(f"Question: [{question_text}] = Answer: [{answer_value}] (ID: {question_id})")
    
    return {"id": config_id, "questions": submission_data}

def update_configuration_question(log, http_client, cwpsa_base_url, config_id, question_id, answer):
    log.info(f"Updating configuration [{config_id}] question ID [{question_id}] with answer [{answer}]")
    
    update_data = [
        {
            "op": "replace",
            "path": "/questions",
            "value": [
                {
                    "questionId": question_id,
                    "answer": answer
                }
            ]
        }
    ]
    
    update_url = f"{cwpsa_base_url}/company/configurations/{config_id}"
    update_response = execute_api_call(log, http_client, "patch", update_url, data=update_data, integration_name="cw_psa")
    
    if update_response and update_response.status_code == 200:
        log.info(f"Successfully updated configuration [{config_id}] question [{question_id}] with new answer")
        return True
    return False

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1755168905267")
            operation = input.get_value("Operation_1755168907892")
            user_identifier = input.get_value("User_1755168921921")
            question = input.get_value("Question_1755202849136")
            answer = input.get_value("Answer_1755202853492")
            contact_id = input.get_value("ContactID_1759383051920")
            active = input.get_value("Active_1759383209338")
            last_login_name = input.get_value("LastLoginName_1759383169373")
            os_type = input.get_value("OSType_1759383189081")
            type = input.get_value("Type_1759383071008")
            configuration_name = input.get_value("Name_1759383061920")
            serial_number = input.get_value("SerialNumber_1759383111403")
            tag_number = input.get_value("TagNumber_1759383129787")
            model_number = input.get_value("ModelNumber_1759383146134")
            site = input.get_value("Site_1759383091327")

        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        operation = operation.strip() if operation else ""
        user_identifier = user_identifier.strip() if user_identifier else ""
        question = question.strip() if question else ""
        answer = answer.strip() if answer else ""

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number input is required")
            return
        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation input is required")
            return
        
        company_identifier, company_name, company_id, company_type = get_company_data_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
        if not company_id:
            record_result(log, ResultLevel.WARNING, f"Unable to retrieve company details from ticket [{ticket_number}]")
            return

        if operation == "Get Configurations":
            configuration_data = {
                "contact_id": contact_id,
                "configuration_name": configuration_name,
                "type": type,
                "site": site,
                "serial_number": serial_number,
                "tag_number": tag_number,
                "model_number": model_number,
                "last_login_name": last_login_name,
                "os_type": os_type,
                "active": active
            }
            configurations = get_configurations_by_criteria(log, http_client, cwpsa_base_url, company_identifier, configuration_data)
            if configurations:
                if len(configurations) > 1:
                    record_result(log, ResultLevel.WARNING, f"Multiple configurations found:")
                    config_ids = []
                    for configuration in configurations:
                        record_result(log, ResultLevel.SUCCESS, f"Configuration ID: {configuration.get('id', '')}")
                        record_result(log, ResultLevel.SUCCESS, f"Configuration Type: {configuration.get('type', {}).get('name', '')}")
                        record_result(log, ResultLevel.SUCCESS, f"Configuration Site: {configuration.get('site', {}).get('name', '')}")
                        record_result(log, ResultLevel.SUCCESS, f"Configuration Serial Number: {configuration.get('serialNumber', '')}")
                        record_result(log, ResultLevel.SUCCESS, f"Configuration Tag Number: {configuration.get('tagNumber', '')}")
                        record_result(log, ResultLevel.SUCCESS, f"Configuration Model Number: {configuration.get('modelNumber', '')}")
                        record_result(log, ResultLevel.SUCCESS, f"Configuration Last Login Name: {configuration.get('lastLoginName', '')}")
                        record_result(log, ResultLevel.SUCCESS, f"Configuration OS Type: {configuration.get('osType', '')}")
                        record_result(log, ResultLevel.SUCCESS, f"Configuration Active: {configuration.get('status', {}).get('name', '')}")
                        config_ids.append(configuration.get('id'))
                    data_to_log["configuration_id"] = config_ids
                    return
                elif len(configurations) == 1:
                    config = configurations[0]
                    record_result(log, ResultLevel.SUCCESS, f"Configuration found:")
                    record_result(log, ResultLevel.SUCCESS, f"Configuration: {config.get('name', '')}")
                    record_result(log, ResultLevel.SUCCESS, f"Configuration ID: {config.get('id', '')}")
                    record_result(log, ResultLevel.SUCCESS, f"Configuration Type: {config.get('type', {}).get('name', '')}")
                    record_result(log, ResultLevel.SUCCESS, f"Configuration Site: {config.get('site', {}).get('name', '')}")
                    record_result(log, ResultLevel.SUCCESS, f"Configuration Serial Number: {config.get('serialNumber', '')}")
                    record_result(log, ResultLevel.SUCCESS, f"Configuration Tag Number: {config.get('tagNumber', '')}")
                    record_result(log, ResultLevel.SUCCESS, f"Configuration Model Number: {config.get('modelNumber', '')}")
                    record_result(log, ResultLevel.SUCCESS, f"Configuration Last Login Name: {config.get('lastLoginName', '')}")
                    record_result(log, ResultLevel.SUCCESS, f"Configuration OS Type: {config.get('osType', '')}")
                    record_result(log, ResultLevel.SUCCESS, f"Configuration Active: {config.get('status', {}).get('name', '')}")
                    data_to_log["configuration_id"] = config.get('id')
            else:
                record_result(log, ResultLevel.INFO, f"No configurations found with details: {configuration_data}")

        if operation == "Get PDC and ADS":
            pdc = get_configuration_data(log, http_client, cwpsa_base_url, company_id, "PDC")
            if pdc != None:
                data_to_log["pdc_name"] = pdc.get("name", "")
                data_to_log["pdc_id"] = pdc.get("id", "")
                data_to_log["pdc_endpoint_id"] = pdc.get("endpoint_id", "")
                record_result(log, ResultLevel.SUCCESS, f"PDC Name: {pdc.get('name', '')}")
                record_result(log, ResultLevel.SUCCESS, f"PDC ID: {pdc.get('id', '')}")
                record_result(log, ResultLevel.SUCCESS, f"PDC Endpoint ID: {pdc.get('endpoint_id', '')}")
            else:
                record_result(log, ResultLevel.INFO, "No configuration with PDC=True found")

            ads = get_configuration_data(log, http_client, cwpsa_base_url, company_id, "ADS")
            if ads != None:
                data_to_log["ads_name"] = ads.get("name", "")
                data_to_log["ads_id"] = ads.get("id", "")
                data_to_log["ads_endpoint_id"] = ads.get("endpoint_id", "")
                record_result(log, ResultLevel.SUCCESS, f"ADS Name: {ads.get('name', '')}")
                record_result(log, ResultLevel.SUCCESS, f"ADS ID: {ads.get('id', '')}")
                record_result(log, ResultLevel.SUCCESS, f"ADS Endpoint ID: {ads.get('endpoint_id', '')}")
            else:
                record_result(log, ResultLevel.INFO, "No configuration with ADS=True found")

        elif operation == "Get Configurations of User":
            if not user_identifier or not company_id:
                record_result(log, ResultLevel.WARNING, "Missing User Identifier or CompanyID for configuration lookup")
                return
            
            contact = get_contact(log, http_client, cwpsa_base_url, user_identifier, company_identifier)
            if not contact:
                record_result(log, ResultLevel.WARNING, f"No contact found with user identifier [{user_identifier}] in company [{company_identifier}]")
                return
            
            contact_id = contact.get("id")
            log.info(f"Found contact ID [{contact_id}] for user identifier [{user_identifier}]")
            
            device_line, device_note = get_user_configuration_items(log, http_client, cwpsa_base_url, company_id, contact_id)
            if device_line:
                data_to_log["Device"] = device_line
                record_result(log, ResultLevel.SUCCESS, "Configuration found assigned to user:")
                for line in device_note.split("\n"):
                    record_result(log, ResultLevel.SUCCESS, line)
            else:
                record_result(log, ResultLevel.SUCCESS, f"No configuration items found for contact ID [{contact_id}]")

        elif operation == "Get Submission":
            submission_data = get_configuration_question(log, http_client, cwpsa_base_url, ticket_number, company_identifier)
            if submission_data:
                data_to_log["Submission_ID"] = submission_data.get("id")
                for question, question_info in submission_data["questions"].items():
                    data_to_log[question] = question_info["answer"]
                record_result(log, ResultLevel.SUCCESS, f"Retrieved submission data for ticket [{ticket_number}] with {len(submission_data['questions'])} identifiers")
            else:
                record_result(log, ResultLevel.WARNING, f"No submission configuration found for ticket [{ticket_number}] or no identifiers provided")

        elif operation == "Update Submission":
            if not question:
                record_result(log, ResultLevel.WARNING, "Question input is required for Update Submission operation")
                return
            if not answer:
                record_result(log, ResultLevel.WARNING, "Answer input is required for Update Submission operation")
                return
            
            submission_data = get_configuration_question(log, http_client, cwpsa_base_url, ticket_number, company_identifier)
            config_id = submission_data.get("id")
            if not config_id:
                record_result(log, ResultLevel.WARNING, f"No configuration found with name [{ticket_number}] and type 'Automation - Submission' to update")
                return

            question_id = None
            for question_text, question_info in submission_data["questions"].items():
                if question_text == question:
                    question_id = question_info.get("questionId")
                    break

            if not question_id:
                record_result(log, ResultLevel.WARNING, f"Question '{question}' not found in configuration [{ticket_number}]")
                return

            success = update_configuration_question(log, http_client, cwpsa_base_url, config_id, question_id, answer)
            if success:
                updated_submission_data = get_configuration_question(log, http_client, cwpsa_base_url, ticket_number, company_identifier)
                if updated_submission_data:
                    data_to_log["Submission_ID"] = updated_submission_data.get("id")
                    for question_text, question_info in updated_submission_data["questions"].items():
                        data_to_log[question_text] = question_info["answer"]
                record_result(log, ResultLevel.SUCCESS, f"Updated question [{question}] with value [{answer}] for configuration [{ticket_number}]")
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to update question [{question}] with value [{answer}] for configuration [{ticket_number}]")

        else:
            record_result(log, ResultLevel.WARNING, f"Unknown operation [{operation}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
