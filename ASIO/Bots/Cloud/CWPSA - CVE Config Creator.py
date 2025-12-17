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

cwpsa_base_url = "https://aus.myconnectwise.net/v4_6_release/apis/3.0"
nist_nvd2_base_url = "https://api.vulncheck.com"

data_to_log = {}
bot_name = "CWPSA - CVE Config Creator"
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

def find_configuration_by_name(log, http_client, cwpsa_base_url, config_name, company_id, config_type="CVE Vulnerability"):
    log.info(f"Searching for configuration with name [{config_name}] and type [{config_type}] in company [{company_id}]")
    
    conditions = f'name="{config_name}" AND type/name="{config_type}" AND company/id={company_id}'
    encoded_conditions = urllib.parse.quote(conditions)
    endpoint = f"{cwpsa_base_url}/company/configurations?conditions={encoded_conditions}"
    
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    if response:
        try:
            configs = response.json()
            if configs and len(configs) > 0:
                config = configs[0]
                log.info(f"Found existing configuration ID [{config.get('id')}] with name: [{config_name}]")
                return config
        except Exception as e:
            log.exception(e, f"Failed to parse configuration search response for name: [{config_name}]")
            return None
    log.info(f"No configuration found with name: [{config_name}]")
    return None

def create_configuration(log, http_client, cwpsa_base_url, config_name, company_id, config_type="CVE Vulnerability", status="Active", questions=None):
    log.info(f"Creating new configuration with name: [{config_name}], type: [{config_type}]")
    
    config_data = {
        "name": config_name,
        "type": {"name": config_type},
        "company": {"id": company_id},
        "status": {"name": status}
    }
    
    if questions:
        config_data["questions"] = questions
        log.info(f"Including {len(questions)} questions in initial POST")
    
    endpoint = f"{cwpsa_base_url}/company/configurations"
    response = execute_api_call(log, http_client, "post", endpoint, data=config_data, integration_name="cw_psa")
    
    if response and response.status_code == 201:
        try:
            config = response.json()
            config_id = config.get("id")
            log.info(f"Successfully created configuration ID [{config_id}] with name: [{config_name}]")
            return config
        except Exception as e:
            log.exception(e, f"Failed to parse create configuration response for name: [{config_name}]")
            return None
    else:
        log.error(f"Failed to create configuration with name: [{config_name}]")
        return None

def update_configuration_questions(log, http_client, cwpsa_base_url, config_id, questions):
    log.info(f"Updating configuration [{config_id}] with {len(questions)} questions")
    
    update_data = [
        {
            "op": "replace",
            "path": "/questions",
            "value": questions
        }
    ]
    
    endpoint = f"{cwpsa_base_url}/company/configurations/{config_id}"
    response = execute_api_call(log, http_client, "patch", endpoint, data=update_data, integration_name="cw_psa")
    
    if response and response.status_code == 200:
        log.info(f"Successfully updated configuration [{config_id}] questions")
        return True
    else:
        log.error(f"Failed to update configuration [{config_id}] questions")
        return False

def get_cve_data_from_nist(log, cve_number):
    log.info(f"Retrieving CVE data for [{cve_number}] from NIST NVD2")
    
    endpoint = f"{nist_nvd2_base_url}/v3/index/nist-nvd2?cve={cve_number}"
    
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="custom_wf_apikey")
    
    if response:
        try:
            cve_data = response.json()
            log.info(f"Successfully retrieved CVE data for [{cve_number}]")
            return cve_data
        except Exception as e:
            log.exception(e, f"Failed to parse CVE data for [{cve_number}]")
    else:
        log.warning(f"Failed to retrieve CVE data for [{cve_number}]")
    
    return None

def main():
    try:
        try:
            cve_number = input.get_value("CVENumber_1765965781687")
            company_id = input.get_value("CompanyID_1765965784144")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch required input values")
            return

        cve_number = cve_number.strip() if cve_number else ""
        
        if not cve_number:
            record_result(log, ResultLevel.WARNING, "CVE number input is required")
            return
        
        if not company_id:
            record_result(log, ResultLevel.WARNING, "Company ID input is required")
            return
        
        try:
            company_id = int(company_id)
        except:
            record_result(log, ResultLevel.WARNING, f"Invalid company ID: [{company_id}]")
            return
        
        log.info(f"Processing CVE: [{cve_number}] for company ID [{company_id}]")
        data_to_log["cve_number"] = cve_number
        data_to_log["company_id"] = company_id
        
        cve_config = find_configuration_by_name(log, http_client, cwpsa_base_url, cve_number, company_id, "CVE Vulnerability")
        
        if cve_config:
            log.info(f"Configuration already exists for CVE: [{cve_number}] (ID: {cve_config.get('id')})")
            data_to_log["config_id"] = cve_config.get("id")
            data_to_log["config_existed"] = True
            record_result(log, ResultLevel.SUCCESS, f"Found existing configuration for CVE: [{cve_number}]")
        else:
            log.info(f"Retrieving CVE data for [{cve_number}] before creating configuration")
            cve_data = get_cve_data_from_nist(log, cve_number)
            
            questions_data = None
            if cve_data:
                log.info(f"CVE data structure keys: {list(cve_data.keys())}")
                
                q1_cve_id = cve_number
                q2_publish_date = ""
                q3_last_modified = ""
                q4_cvss_version = ""
                q5_cvss_score = ""
                q6_advisory_url = ""
                q7_advisory_sources = ""
                q8_advisory_tags = ""
                q9_description = "No description available"
                q10_nvd_status = ""
                
                if "data" in cve_data and isinstance(cve_data["data"], list) and len(cve_data["data"]) > 0:
                    cve_details = cve_data["data"][0]
                    log.info(f"CVE details keys: {list(cve_details.keys())}")
                    
                    q1_cve_id = cve_details.get("id", cve_number)
                    
                    published_raw = cve_details.get("published", "")
                    q2_publish_date = published_raw.split("T")[0] if published_raw else ""
                    
                    modified_raw = cve_details.get("lastModified", "")
                    q3_last_modified = modified_raw.split("T")[0] if modified_raw else ""
                    
                    q10_nvd_status = cve_details.get("vulnStatus", "")
                    
                    if "descriptions" in cve_details and isinstance(cve_details["descriptions"], list) and len(cve_details["descriptions"]) > 0:
                        q9_description = cve_details["descriptions"][0].get("value", "No description available")
                        log.info(f"Extracted description: [{q9_description[:100]}...]")
                    
                    if "metrics" in cve_details and "cvssMetricV31" in cve_details["metrics"]:
                        metrics = cve_details["metrics"]["cvssMetricV31"]
                        if isinstance(metrics, list) and len(metrics) > 0:
                            cvss_data = metrics[0].get("cvssData", {})
                            base_score = cvss_data.get("baseScore", "")
                            cvss_version = cvss_data.get("version", "")
                            if base_score:
                                q5_cvss_score = str(base_score)
                            if cvss_version:
                                q4_cvss_version = f"CVSS Version {cvss_version}"
                            log.info(f"Extracted CVSS: Version [{q4_cvss_version}], Score [{q5_cvss_score}]")
                    
                    if "references" in cve_details and isinstance(cve_details["references"], list) and len(cve_details["references"]) > 0:
                        first_ref = cve_details["references"][0]
                        q6_advisory_url = first_ref.get("url", "")
                        q7_advisory_sources = first_ref.get("source", "")
                        tags = first_ref.get("tags", [])
                        if tags and isinstance(tags, list):
                            q8_advisory_tags = ", ".join(tags)
                        log.info(f"Extracted advisory: URL [{q6_advisory_url}], Source [{q7_advisory_sources}], Tags [{q8_advisory_tags}]")
                else:
                    log.warning(f"Unable to extract CVE details from data structure")
                
                questions_data = [
                    {"questionId": 484, "answer": q1_cve_id},
                    {"questionId": 485, "answer": q2_publish_date},
                    {"questionId": 486, "answer": q3_last_modified},
                    {"questionId": 487, "answer": q4_cvss_version},
                    {"questionId": 488, "answer": q5_cvss_score},
                    {"questionId": 489, "answer": q6_advisory_url},
                    {"questionId": 490, "answer": q7_advisory_sources},
                    {"questionId": 491, "answer": q8_advisory_tags},
                    {"questionId": 492, "answer": q9_description},
                    {"questionId": 493, "answer": q10_nvd_status}
                ]
                log.info(f"Mapped all CVE fields to 10 questions")
            else:
                log.warning(f"No CVE data retrieved for [{cve_number}] - using default values for required questions")
                questions_data = [
                    {"questionId": 484, "answer": cve_number},
                    {"questionId": 485, "answer": ""},
                    {"questionId": 486, "answer": ""},
                    {"questionId": 487, "answer": ""},
                    {"questionId": 488, "answer": ""},
                    {"questionId": 489, "answer": ""},
                    {"questionId": 490, "answer": ""},
                    {"questionId": 491, "answer": ""},
                    {"questionId": 492, "answer": "No description available"},
                    {"questionId": 493, "answer": ""}
                ]
            
            log.info(f"Creating configuration item for CVE: [{cve_number}]")
            cve_config = create_configuration(log, http_client, cwpsa_base_url, cve_number, company_id, "CVE Vulnerability", "Active", questions_data)
            
            if cve_config:
                data_to_log["config_id"] = cve_config.get("id")
                data_to_log["config_existed"] = False
                record_result(log, ResultLevel.SUCCESS, f"Created configuration for CVE: [{cve_number}]")
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to create configuration for CVE: [{cve_number}]")
                data_to_log["config_id"] = None
                data_to_log["config_existed"] = False
                return
        
        log.info(f"Completed CVE config creation/retrieval for [{cve_number}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
