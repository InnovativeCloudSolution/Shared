import re
import json
from cw_rpa import Logger, Input

log = Logger()
log.info("Imports completed successfully.")

def main():
    try:
        input = Input()
        string = input.get_value("String_1736292093000")
        pattern = input.get_value("Regex_1736292115905")
    except Exception as e:
        log.exception(e, "An error occurred while getting input values")
        log.result_failed_message("An error occurred while getting input values")
        return

    log.info(f"Fetching details for string: {string} and pattern: {pattern}")

    try:
        matches = re.findall(pattern, string)
        log.info(f"Match: {matches}")

        if matches:
            if len(matches) > 1:
                data_to_log = {"matches": matches, "match_count": len(matches)}
                log.info(f"Multiple match: {json.dumps(data_to_log)}")
                log.result_success_message(f"Multiple match: {matches_string}")
            else:
                matches_string = matches[0]
                data_to_log = {"matches": matches_string, "match_count": 1}
                log.info(f"Single match: {json.dumps(data_to_log)}")
                log.result_success_message(f"Single match: {matches_string}")
        else:
            data_to_log = {"matches": "", "match_count": 0}
            log.info(f"No Match: {json.dumps(data_to_log)}")
            log.result_success_message("No matches found.")

        log.result_data(data_to_log)

    except Exception as e:
        log.exception(e, "Unexpected error during match retrieval.")
        log.result_failed_message("Unexpected error during pattern matching.")
        return

if __name__ == "__main__":
    main()
