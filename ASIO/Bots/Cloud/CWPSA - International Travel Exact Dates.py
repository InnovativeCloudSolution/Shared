import sys
import re
import os
from datetime import datetime, timedelta, timezone
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
bot_name = "MIT-CWPSA - International Travel Extract Dates"
log.info("Static variables set")

def record_result(log, level, message):
    log.result_message(level, f"[{bot_name}]: {message}")
    if level == ResultLevel.WARNING:
        data_to_log["status_result"] = "Fail"
    elif level == ResultLevel.SUCCESS:
        if "status_result" not in data_to_log or data_to_log["status_result"] != "Fail":
            data_to_log["status_result"] = "Success"

def extract_start_end_datetime(input_str):
    pattern = r"\b(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s\d{1,2}\s(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec),\s\d{4}\s\d{2}:\d{2}"
    matches = re.findall(pattern, input_str)

    if len(matches) == 2:
        start_dt = datetime.strptime(matches[0].strip(), "%a %d %b, %Y %H:%M")
        end_dt = datetime.strptime(matches[1].strip(), "%a %d %b, %Y %H:%M")

        start_formatted = start_dt.strftime("%m/%d/%Y %I:%M %p")
        end_formatted = end_dt.strftime("%m/%d/%Y %I:%M %p")

        start_midnight = start_dt.replace(hour=0, minute=0, second=0, microsecond=0)
        end_midnight = end_dt.replace(hour=0, minute=0, second=0, microsecond=0) + timedelta(days=1)

        start_string = start_midnight.strftime("%Y-%m-%dT%H:%M")
        end_string = end_midnight.strftime("%Y-%m-%dT%H:%M")

        return start_formatted, end_formatted, start_string, end_string
    else:
        raise ValueError("Could not extract two valid datetime strings")

def main():
    try:
        input_summary = input.get_value("TicketSummary_1757461359295").strip()

        log.info(f"Received input summary = [{input_summary}]")

        if not input_summary:
            log.error("Input summary is empty or invalid")
            record_result(log, ResultLevel.WARNING, "Input summary is empty or invalid")
            return

        try:
            start_formatted, end_formatted, start_string, end_string = extract_start_end_datetime(input_summary)

            log.info(f"Extracted start = [{start_formatted}], end = [{end_formatted}]")

            data_to_log["start"] = start_string
            data_to_log["end"] = end_string
            log.result_data(data_to_log)

            record_result(log, ResultLevel.SUCCESS, f"Start: {start_formatted}, End: {end_formatted}")
        
        except ValueError as ve:
            log.warning(f"Datetime extraction failed: {ve}")
            record_result(log, ResultLevel.WARNING, "Unable to extract start and end datetime")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()