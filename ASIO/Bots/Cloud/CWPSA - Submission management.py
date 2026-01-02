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

cwpsa_base_url = "https://aus.myconnectwise.net"
cwpsa_base_url_path = "/v4_6_release/apis/3.0"
vault_name = "PLACEHOLDER-akv1"

data_to_log = {}
bot_name = "CWPSA - Submission management"
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

def get_company_data_from_ticket(log, http_client, cwpsa_base_url, cwpsa_base_url_path, ticket_number):
    log.info(f"Retrieving company details for ticket [{ticket_number}]")
    ticket_endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets/{ticket_number}"
    ticket_response = execute_api_call(log, http_client, "get", ticket_endpoint, integration_name="cw_psa")

    if ticket_response and ticket_response.status_code == 200:
        ticket_data = ticket_response.json()
        company = ticket_data.get("company", {})

        company_id = company["id"]
        company_identifier = company["identifier"]
        company_name = company["name"]

        log.info(f"Company ID: [{company_id}], Identifier: [{company_identifier}], Name: [{company_name}]")

        company_endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/company/companies/{company_id}"
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

def get_configuration_question(log, http_client, cwpsa_base_url, ticket_number, company_identifier):
    log.info(f"Searching for configuration with name [{ticket_number}] and type 'Automation - Submission' in company [{company_identifier}]")
    
    conditions = f'name="{ticket_number}" AND type/name="Automation - Submission" AND company/identifier="{company_identifier}"'
    query_string = f"conditions={urllib.parse.quote(conditions)}"
    url = f"{cwpsa_base_url}{cwpsa_base_url_path}/company/configurations?{query_string}"
    
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
    
    update_url = f"{cwpsa_base_url}{cwpsa_base_url_path}/company/configurations/{config_id}"
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
            question = input.get_value("Question_1755202849136")
            answer = input.get_value("Answer_1755202853492")

        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        operation = operation.strip() if operation else ""
        question = question.strip() if question else ""
        answer = answer.strip() if answer else ""

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number input is required")
            return
        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation input is required")
            return
        
        company_identifier, company_name, company_id, company_type = get_company_data_from_ticket(log, http_client, cwpsa_base_url, cwpsa_base_url_path, ticket_number)
        if not company_id:
            record_result(log, ResultLevel.WARNING, f"Unable to retrieve company details from ticket [{ticket_number}]")
            return

        if operation == "Get Submission":
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
