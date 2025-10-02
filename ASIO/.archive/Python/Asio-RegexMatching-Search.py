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
        match = re.search(pattern, string)
        log.info(f"Match: {match}")

        if match:
            matched_value = match.group()
            data_to_log = {"output": matched_value}
            log.info(f"Matched: {json.dumps(data_to_log)}")
            log.result_success_message(f"Matched: {matched_value}")
        else:
            data_to_log = {"output": None}
            log.error("Pattern matching failed. No matches found.")
            log.info(f"Output: {json.dumps(data_to_log)}")
            log.result_failed_message("Pattern matching failed. No matches found.")

    except Exception as e:
        log.exception(e, "Unexpected error during match retrieval.")
        log.result_failed_message("Unexpected error during pattern matching.")
        return


if __name__ == "__main__":
    main()
