# Synopsis
# --------
# Bot Name       : FetchAsioEndpointsBot
# Description    : This bot fetches endpoint data from the Asio API and logs the number of endpoints.
# Developed by   : Prakash Chalgeri
# Created Date   : 21/01/2025 (dd/mm/yyyy)
# Updated On     : 
# Version        : 1.0
# Jira           : RPA

# Import System libraries required for Pre-Check

import subprocess
import os
import json
import sys
import hashlib
import urllib
import random
import traceback
import random
from datetime import datetime
from cw_rpa import Logger, Input, HttpClient

# Declare global variables & initialize logger and HTTP client
log = Logger()
http_client = HttpClient()
input = Input()

# Function block started
# Log errors with a unique reference
def generate_error_reference():
    error_ref = random.randint(10000, 99999)
    return f'#{error_ref}'

def log_stdout(msg):
    error_code = generate_error_reference()
    sys.stdout.write(f"\nError reference: {error_code}\nError: {msg}\nTraceback : {traceback.format_exc()}".replace('\n', '\\n'))
    log.error(f"An internal error occured, error reference: {error_code}")
    log.result_failed_message(f"An internal error occured, error reference: {error_code}")
# Function block end

# Main function to execute the bot logic
def main():
    log.info("Bot execution has started.")

    # Step 1: Fetch input values
    try:
        base_url = input.get_value("cwOpenAPIURL")
        endpoint = f"{base_url}/api/platform/v1/device/endpoints"
        log.info(f"Base URL fetched successfully: {base_url}")
    except Exception as e:
        log_stdout(f"Error while fetching input values: {e}")
        return

    # Step 2: Make the API call
    try:
        log.info(f"Executing API call to endpoint: {endpoint}")
        response = http_client.get(endpoint)
        response.raise_for_status()
        log.result_success_message("API call executed successfully.")
    except Exception as e:
        log_stdout(f"API call failed: {e}")
        return

    # Step 3: Process the response
    try:
        res_json = response.json()
        endpoints = res_json.get("endpoints", [])
        log.info(f"Number of endpoints retrieved: {len(endpoints)}")
        log.result_success_message(f"Number of endpoints: {len(endpoints)}")
    except Exception as e:
        log_stdout(f"Error while processing API response: {e}")

    log.info("Bot execution has ended.")

# Entry point
if __name__ == "__main__":
    main()











