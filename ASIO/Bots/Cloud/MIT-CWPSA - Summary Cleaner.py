import sys
import re
import os
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

cwpsa_base_url = "https://au.myconnectwise.net/v4_6_release/apis/3.0"
msgraph_base_url = "https://graph.microsoft.com/v1.0"
msgraph_base_url_beta = "https://graph.microsoft.com/beta"
vault_name = "mit-azu1-prod1-akv1"

data_to_log = {}
bot_name = "MIT-CWPSA - Summary Cleaner"
log.info("Static variables set")

def record_result(log, level, message):
    log.result_message(level, f"[{bot_name}]: {message}")
    if level == ResultLevel.WARNING:
        data_to_log["status_result"] = "Fail"
    elif level == ResultLevel.SUCCESS:
        if "status_result" not in data_to_log or data_to_log["status_result"] != "Fail":
            data_to_log["status_result"] = "Success"

def clean_summary_description(text: str) -> str:
    cleaned = re.sub(r'^\s*\[Extended Summary\]\s*', '', text, flags=re.IGNORECASE)
    cleaned = re.sub(r'\s*\[Description\]\s*$', '', cleaned, flags=re.IGNORECASE)
    return cleaned.strip()

def remove_email_brackets(text: str) -> str:
    return re.sub(r'\s*<[^<>]+>', '', text).strip()

def main():
    try:
        input_note = input.get_value("InputNote_1756876308990").strip()

        log.info(f"Received input note = [{input_note}]")

        if not input_note or not input_note.strip():
            log.error("Input note is empty or invalid")
            record_result(log, ResultLevel.WARNING, "Input note is empty or invalid")
            return

        cleaned_note = clean_summary_description(input_note)
        cleaned_note = remove_email_brackets(cleaned_note)
        cleaned_note = cleaned_note.strip().splitlines()[0]

        log.info(f"Final cleaned summary = [{cleaned_note}]")
        data_to_log["cleanedsummary"] = cleaned_note
        log.result_data(data_to_log)
        record_result(log, ResultLevel.SUCCESS, f"Summary Cleaned: [{cleaned_note}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()