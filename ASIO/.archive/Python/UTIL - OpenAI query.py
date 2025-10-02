import sys
import os
import json
import base64
import hashlib
import hmac
import urllib.parse
import random
import time
from datetime import datetime
from openai import OpenAI
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

vault_name = "mit-azu1-prod1-akv1"
cwpsa_base_url = "https://au.myconnectwise.net/v4_6_release/apis/3.0"
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
                response = getattr(http_client.third_party_integration(integration_name), method)(url=endpoint, json=data) if data else getattr(http_client.third_party_integration(integration_name), method)(url=endpoint)
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

            if response.status_code in [200, 204]:
                return response
            if response.status_code in [429, 503]:
                retry_after = response.headers.get("Retry-After")
                wait_time = int(retry_after) if retry_after else base_delay * (2 ** attempt) + random.uniform(0, 3)
                log.warning(f"Rate limit exceeded. Retrying in {wait_time:.2f} seconds")
                time.sleep(wait_time)
            elif response.status_code == 404:
                log.warning(f"Skipping non-existent resource [{endpoint}]")
                return None
            else:
                log.error(f"API request failed Status: {response.status_code}, Response: {response.text}")
                return response
        except Exception as e:
            log.exception(e, f"Exception during API call to {endpoint}")
            return None
    return None

def execute_cosmosdb_call(log, http_client, method, endpoint, data=None, headers=None):
    if not all([endpoint, headers]):
        log.error("Missing required Cosmos DB call parameters")
        return None
    response = execute_api_call(log, http_client, method, endpoint, data=data, headers=headers)
    if response and response.status_code in [200, 201]:
        return response
    log.error(f"Cosmos DB call failed Status: {response.status_code if response else 'N/A'}, Body: {response.text if response else 'N/A'}")
    return None

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

def generate_cosmosdb_auth_header(verb, resource_type, resource_link, date_utc, master_key):
    key = base64.b64decode(master_key)
    text = f"{verb.lower()}\n{resource_type.lower()}\n{resource_link}\n{date_utc.lower()}\n\n"
    signature = base64.b64encode(hmac.new(key, text.encode("utf-8"), hashlib.sha256).digest()).decode()
    return f"type=master&ver=1.0&sig={urllib.parse.quote(signature)}"

def get_company_identifier_from_ticket(log, http_client, cwpsa_base_url, ticket_number):
    log.info(f"Retrieving company identifier for ticket [{ticket_number}]")
    endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}"
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    if response:
        if response.status_code == 200:
            data = response.json()
            company_identifier = data.get("company", {}).get("identifier", "")
            if company_identifier:
                log.info(f"Company identifier for ticket [{ticket_number}] is [{company_identifier}]")
                return company_identifier
            else:
                log.error(f"Company identifier not found in response for ticket [{ticket_number}]")
        else:
            log.error(f"Failed to retrieve company identifier for ticket [{ticket_number}] Status: {response.status_code}, Body: {response.text}")
    else:
        log.error(f"Failed to retrieve company identifier for ticket [{ticket_number}]: No response received")
    return ""

def get_persona_template(log, http_client, persona_name, company_identifier, endpoint_url, cosmos_key):
    try:
        db_name = company_identifier.lower()
        container_name = "persona"

        verb = "post"
        resource_type = "docs"
        resource_link = f"dbs/{db_name}/colls/{container_name}"
        date_utc = datetime.utcnow().strftime('%a, %d %b %Y %H:%M:%S GMT')
        auth_header = generate_cosmosdb_auth_header(verb, resource_type, resource_link, date_utc, cosmos_key)

        query = {
            "query": "SELECT * FROM c WHERE c[\"persona\"] = @persona",
            "parameters": [{"name": "@persona", "value": persona_name}]
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
        response = execute_cosmosdb_call(log, http_client, verb, endpoint, data=query, headers=headers)

        if response and response.status_code == 200:
            documents = response.json().get("Documents", [])
            if documents:
                log.info(f"Found {len(documents)} persona template(s) for [{persona_name}] in [{db_name}]")
                return documents[0]
            else:
                log.info(f"No persona template found for [{persona_name}] in [{db_name}]")
                return {}

        log.error(f"Failed to query persona template Status: {response.status_code if response else 'N/A'}, Body: {response.text if response else 'N/A'}")
        return {}

    except Exception as e:
        log.exception(e, f"Exception occurred while querying persona template for [{persona_name}]")
        return {}

def get_user_onboarding_default(log, http_client, company_identifier, endpoint_url, cosmos_key):
    try:
        db_name = company_identifier.lower()
        container_name = "user_onboarding_default"

        verb = "post"
        resource_type = "docs"
        resource_link = f"dbs/{db_name}/colls/{container_name}"
        date_utc = datetime.utcnow().strftime('%a, %d %b %Y %H:%M:%S GMT')
        auth_header = generate_cosmosdb_auth_header(verb, resource_type, resource_link, date_utc, cosmos_key)

        query = {
            "query": "SELECT * FROM c"
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
        response = execute_cosmosdb_call(log, http_client, verb, endpoint, data=query, headers=headers)

        if response and response.status_code == 200:
            documents = response.json().get("Documents", [])
            if documents:
                log.info(f"Found {len(documents)} onboarding default document(s) in [{db_name}]")
                return documents[0]
            else:
                log.info(f"No onboarding defaults found in [{db_name}]")
                return {}

        log.error(f"Failed to query onboarding defaults Status: {response.status_code if response else 'N/A'}, Body: {response.text if response else 'N/A'}")
        return {}

    except Exception as e:
        log.exception(e, "Exception occurred while querying all onboarding defaults")
        return {}

def query_openai(log, api_key, prompt):
    try:
        client = OpenAI(api_key=api_key)
        response = client.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a helpful assistant"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.5
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        log.error(f"GPT query failed: {str(e)}")
        return ""

def main():
    try:
        try:
            ticket_number = input.get_value("TicketNumber_1742953896873")
            persona = input.get_value("Persona_1746937219795")
            prompt_template = input.get_value("PromptTemplate_1747968620394")
            db_name = input.get_value("DatabaseName_1748299609003")
            container_name = input.get_value("ContainerName_1748299612528")
        except:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""
        persona = persona.strip() if persona else ""
        prompt_template = prompt_template.strip() if prompt_template else ""
        db_name = db_name.strip().lower() if db_name else ""
        container_name = container_name.strip() if container_name else ""

        if not ticket_number or not prompt_template or not db_name or not container_name:
            record_result(log, ResultLevel.WARNING, "Missing required input values")
            return

        company = get_company_identifier_from_ticket(log, http_client, cwpsa_base_url, ticket_number)
        if not company:
            record_result(log, ResultLevel.WARNING, "Unable to retrieve company from CW")
            return

        cosmos_endpoint = get_secret_value(log, http_client, vault_name, "MIT-AZU1-CosmosDB-Endpoint")
        cosmos_key = get_secret_value(log, http_client, vault_name, "MIT-AZU1-CosmosDB-Key")
        if not cosmos_endpoint or not cosmos_key:
            record_result(log, ResultLevel.WARNING, "Failed to get Cosmos DB credentials")
            return

        user_defaults = get_user_onboarding_default(log, http_client, db_name, cosmos_endpoint, cosmos_key)
        persona_defaults = get_persona_template(log, http_client, persona, db_name, cosmos_endpoint, cosmos_key)

        combined_payload = {
            "ticket": ticket_number,
            "company": company,
            "persona": persona,
            "defaults": user_defaults,
            "persona_template": persona_defaults
        }

        final_prompt = prompt_template.replace("$input", json.dumps(combined_payload, indent=2))

        api_key = get_secret_value(log, http_client, vault_name, "MIT-OpenAI-APIKey")
        if not api_key:
            record_result(log, ResultLevel.WARNING, "Missing OpenAI API key")
            return

        gpt_response = query_openai(log, api_key, final_prompt)
        if not gpt_response:
            record_result(log, ResultLevel.WARNING, "No response from GPT")
            return

        data_to_log["response"] = gpt_response
        data_to_log["payload"] = combined_payload
        data_to_log["prompt"] = final_prompt
        record_result(log, ResultLevel.SUCCESS, "Completed successfully")

    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()