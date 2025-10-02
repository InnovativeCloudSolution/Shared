import sys
import random
import os
import time
import requests
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

vault_name = "mit-azu1-prod1-akv1"
dictionary_url = "https://mitazu1pubfilestore.blob.core.windows.net/automation/Password_safe_wordlist.txt"

data_to_log = {}
bot_name = "MIT-UTIL - Generate strong password"
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

def get_random_word_list(log, http_client) -> list:
    log.info("Fetching word list from safe wordlist")
    response = execute_api_call(log, http_client, "get", dictionary_url)
    if not response:
        log.error("API call to fetch word list failed. No response received.")
        return []
    if response.status_code != 200:
        log.error(f"Failed to retrieve word list, HTTP Status: {response.status_code}, Response: {response.text}")
        return []
    try:
        lines = response.text.splitlines()
        words = [line.strip().lower() for line in lines if line.strip()]
    except Exception as e:
        log.exception("Failed to parse safe word list")
        return []
    log.info(f"Total words retrieved: {len(words)}")
    if not words:
        log.error("No words found, cannot generate password.")
        return []
    return words

def generate_secure_password(log, http_client, word_count: int) -> str:
    log.info(f"Generating password with [{word_count}] words")
    word_list = get_random_word_list(log, http_client)
    if len(word_list) < word_count:
        log.error("Not enough words to generate password")
        return ""
    words = random.sample(word_list, word_count)
    cap_index = random.randint(0, word_count - 1)
    words[cap_index] = words[cap_index].capitalize()
    password_base = "-".join(words)
    digits = f"{random.randint(0, 99):02d}"
    password = f"{password_base}-{digits}"
    return password

def main():
    try:
        try:
            word_count_str = input.get_value("WordCount_1756874896234")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        word_count_str = word_count_str.strip() if word_count_str else ""
        log.info(f"Received input Word Count=[{word_count_str}]")

        if not word_count_str or not word_count_str.isdigit():
            record_result(log, ResultLevel.WARNING, "Invalid or missing word count input")
            return

        word_count = int(word_count_str)
        password = generate_secure_password(log, http_client, word_count)
        if not password:
            record_result(log, ResultLevel.WARNING, "Failed to generate password")
            return
        else:
            data_to_log["password"] = password
            record_result(log, ResultLevel.SUCCESS, "Password generated successfully")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()