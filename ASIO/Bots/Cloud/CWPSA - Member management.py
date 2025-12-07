import random
import time
import urllib.parse
import requests
from cw_rpa import Logger, Input, HttpClient, ResultLevel

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

cwpsa_base_url = "https://au.myconnectwise.net/v4_6_release/apis/3.0"

data_to_log = {}
bot_name = "CWPSA - Member management"
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

def get_members(log, http_client, cwpsa_base_url, status_filter):
    log.info(f"Retrieving members with status filter: [{status_filter}]")
    conditions = []
    if status_filter:
        conditions.append(f"inactiveFlag={status_filter}")
    query_string = f"conditions={urllib.parse.quote(' AND '.join(conditions))}" if conditions else ""
    endpoint = f"{cwpsa_base_url}/system/members"
    if query_string:
        endpoint = f"{endpoint}?{query_string}"
    
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    if response and response.status_code == 200:
        members = response.json()
        log.info(f"Retrieved {len(members)} members")
        return members
    else:
        log.error(f"Failed to retrieve members")
        return []

def get_member_skills(log, http_client, cwpsa_base_url, member_id):
    log.info(f"Retrieving skills for member ID [{member_id}]")
    endpoint = f"{cwpsa_base_url}/system/members/{member_id}/skills"
    
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    if response and response.status_code == 200:
        skills = response.json()
        log.info(f"Retrieved {len(skills)} skills for member ID [{member_id}]")
        return skills
    else:
        log.error(f"Failed to retrieve skills for member ID [{member_id}]")
        return []

def handle_get_members(log, http_client, cwpsa_base_url, status_filter, data_to_log):
    members = get_members(log, http_client, cwpsa_base_url, status_filter)
    if members:
        member_list = []
        for member in members:
            member_info = {
                "id": member.get("id", ""),
                "identifier": member.get("identifier", ""),
                "firstName": member.get("firstName", ""),
                "lastName": member.get("lastName", ""),
                "emailAddress": member.get("emailAddress", ""),
                "title": member.get("title", ""),
                "inactive": member.get("inactiveFlag", False)
            }
            member_list.append(member_info)
            record_result(log, ResultLevel.SUCCESS, f"Member ID: {member_info['id']} | {member_info['firstName']} {member_info['lastName']} | {member_info['identifier']} | Email: {member_info['emailAddress']}")
        
        data_to_log["members"] = member_list
        data_to_log["member_count"] = len(member_list)
        record_result(log, ResultLevel.SUCCESS, f"Total members retrieved: {len(member_list)}")
    else:
        record_result(log, ResultLevel.INFO, "No members found")

def handle_get_member_skills(log, http_client, cwpsa_base_url, member_id, data_to_log):
    if not member_id:
        record_result(log, ResultLevel.WARNING, "Member ID is required for retrieving skills")
        return
    
    skills = get_member_skills(log, http_client, cwpsa_base_url, member_id)
    if skills:
        skill_list = []
        for skill in skills:
            skill_info = {
                "id": skill.get("id", ""),
                "skill": skill.get("skill", {}).get("name", ""),
                "skillCategory": skill.get("skillCategory", {}).get("name", ""),
                "item": skill.get("item", {}).get("name", ""),
                "type": skill.get("type", {}).get("name", ""),
                "subType": skill.get("subType", {}).get("name", "")
            }
            skill_list.append(skill_info)
            record_result(log, ResultLevel.SUCCESS, f"Skill: {skill_info['skill']} | Category: {skill_info['skillCategory']} | Type: {skill_info['type']} | SubType: {skill_info['subType']}")
        
        data_to_log["member_id"] = member_id
        data_to_log["skills"] = skill_list
        data_to_log["skill_count"] = len(skill_list)
        record_result(log, ResultLevel.SUCCESS, f"Total skills retrieved for member ID [{member_id}]: {len(skill_list)}")
    else:
        record_result(log, ResultLevel.INFO, f"No skills found for member ID [{member_id}]")

def main():
    try:
        try:
            operation = input.get_value("Operation_xxxxxxxxxxxxx")
            status_filter = input.get_value("StatusFilter_xxxxxxxxxxxxx")
            member_id = input.get_value("MemberID_xxxxxxxxxxxxx")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        operation = operation.strip() if operation else ""
        status_filter = status_filter.strip() if status_filter else ""
        member_id = member_id.strip() if member_id else ""

        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation input is required")
            return

        if operation == "Get All Members":
            handle_get_members(log, http_client, cwpsa_base_url, status_filter, data_to_log)

        elif operation == "Get Member Skills":
            handle_get_member_skills(log, http_client, cwpsa_base_url, member_id, data_to_log)

        else:
            record_result(log, ResultLevel.WARNING, f"Unknown operation [{operation}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()