import sys
import os
import re
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

cwpsa_base_url = "https://aus.myconnectwise.net/v4_6_release/apis/3.0"

data_to_log = {}
bot_name = "CWPSA - CVE Extractor"
log.info("Static variables set")

def record_result(log, level, message):
    log.result_message(level, f"[{bot_name}]: {message}")
    if level == ResultLevel.WARNING:
        data_to_log["status_result"] = "Fail"
    elif level == ResultLevel.SUCCESS:
        if "status_result" not in data_to_log or data_to_log["status_result"] != "Fail":
            data_to_log["status_result"] = "Success"

def execute_api_call(log, http_client, method, endpoint, integration_name="cw_psa"):
    log.info(f"Executing API call: {method.upper()} {endpoint}")
    try:
        response = getattr(http_client.third_party_integration(integration_name), method)(url=endpoint)
        if 200 <= response.status_code < 300:
            return response
        elif response.status_code == 404:
            log.warning(f"Resource not found: [{endpoint}]")
            return None
        else:
            log.error(f"API error Status: {response.status_code}, Response: {response.text}")
            return None
    except Exception as e:
        log.exception(e, f"Exception during API call to {endpoint}")
        return None

def get_ticket_notes(log, http_client, cwpsa_base_url, ticket_number):
    log.info(f"Retrieving notes for ticket [{ticket_number}]")
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/service/tickets/{ticket_number}/notes"
    response = execute_api_call(log, http_client, "get", endpoint)
    if response:
        try:
            notes = response.json()
            log.info(f"Retrieved {len(notes)} notes for ticket [{ticket_number}]")
            return notes
        except Exception as e:
            log.exception(e, f"Failed to parse notes response for ticket [{ticket_number}]")
            return []
    log.warning(f"No notes found for ticket [{ticket_number}]")
    return []

def extract_cve_numbers(log, text):
    cve_pattern = r'CVE-\d+-\d+'
    matches = re.findall(cve_pattern, text, re.IGNORECASE)
    if matches:
        unique_cves = list(set([match.upper() for match in matches]))
        log.info(f"Found CVE numbers: {unique_cves}")
        return unique_cves
    return []

def main():
    try:
        try:
            ticket_number = input.get_value("TicketID_1765965563857")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch ticket number input")
            return

        ticket_number = ticket_number.strip() if ticket_number else ""

        if not ticket_number:
            record_result(log, ResultLevel.WARNING, "Ticket number input is required")
            return
        
        log.info(f"Extracting CVE numbers from ticket [{ticket_number}]")
        data_to_log["ticket_id"] = ticket_number
        
        notes = get_ticket_notes(log, http_client, cwpsa_base_url, ticket_number)
        if not notes or len(notes) == 0:
            record_result(log, ResultLevel.WARNING, f"No notes found for ticket [{ticket_number}]")
            data_to_log["cve_list"] = []
            data_to_log["cve_count"] = 0
            return
        
        all_cves = []
        for note in notes:
            note_text = note.get("text", "")
            cves_in_note = extract_cve_numbers(log, note_text)
            all_cves.extend(cves_in_note)
        
        all_cves = list(set(all_cves))
        data_to_log["cve_list"] = all_cves
        data_to_log["cve_count"] = len(all_cves)
        
        if len(all_cves) == 0:
            record_result(log, ResultLevel.INFO, "No CVE numbers found in ticket notes")
        else:
            record_result(log, ResultLevel.SUCCESS, f"Found {len(all_cves)} unique CVE numbers: {', '.join(all_cves)}")
        
        log.info(f"Completed CVE extraction for ticket [{ticket_number}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
