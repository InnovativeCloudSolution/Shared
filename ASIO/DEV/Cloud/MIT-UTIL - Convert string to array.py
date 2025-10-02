import os
import sys
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

vault_name = "mit-azu1-prod1-akv1"

data_to_log = {}
bot_name = "MIT-UTIL - Convert string to array"
log.info("Static variables set")

def record_result(log, level, message):
    log.result_message(level, f"[{bot_name}]: {message}")
    if level == ResultLevel.WARNING:
        data_to_log["status_result"] = "Fail"
    elif level == ResultLevel.SUCCESS:
        if "status_result" not in data_to_log or data_to_log["status_result"] != "Fail":
            data_to_log["status_result"] = "Success"

def normalize_entry(entry, d1, d2, d3, d4):
    parts = []

    if d4 == "|" and "|" in entry:
        entry, string5 = entry.split("|", 1)
        string5 = string5.strip()
    else:
        string5 = ""

    delimiters = [d1, d2, d3]
    remaining = entry

    for delimiter in delimiters:
        if delimiter and delimiter in remaining:
            split_part, remaining = remaining.split(delimiter, 1)
            parts.append(split_part.strip())
        else:
            parts.append(remaining.strip())
            remaining = ""

    parts.append(remaining.strip())
    parts = parts[:4] + [""] * (4 - len(parts))
    parts.append(string5)

    return {
        "String1": parts[0],
        "String2": parts[1],
        "String3": parts[2],
        "String4": parts[3],
        "String5": parts[4]
    }

def convert_string_to_list(log, input_string, d1, d2, d3, d4):
    output = []
    if not input_string:
        return output

    entries = [e.strip() for e in input_string.split(",") if e.strip()]
    for entry in entries:
        output.append(normalize_entry(entry, d1, d2, d3, d4))

    log.info(f"Parsed {len(output)} entries into structured list")
    return output

def main():
    try:
        try:
            raw_string = input.get_value("StringToConvert_1756863426826")
            delimiter1 = input.get_value("Delimiter1_1756863464025")
            delimiter2 = input.get_value("Delimiter2_1756863465105")
            delimiter3 = input.get_value("Delimiter3_1756863466106")
            delimiter4 = input.get_value("Delimiter4_1756863467489")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        raw_string = raw_string.strip() if raw_string else ""
        delimiter1 = delimiter1.strip() if delimiter1 else ":"
        delimiter2 = delimiter2.strip() if delimiter2 else ":"
        delimiter3 = delimiter3.strip() if delimiter3 else ":"
        delimiter4 = delimiter4.strip() if delimiter4 else "|"

        log.info(f"Received input string = [{raw_string}]")
        structured_list = convert_string_to_list(
            log, raw_string, delimiter1, delimiter2, delimiter3, delimiter4
        )
        data_to_log["Parsed"] = structured_list

        log_lines = [f"Processed {len(structured_list)} entries from Parsed input"]
        for row in structured_list:
            line = []
            for i in range(1, 6):
                value = row.get(f"String{i}", "").strip()
                if value:
                    line.append(f"{i} [{value}]")
            log_lines.append("- " + ", ".join(line))

        for msg in log_lines:
            record_result(log, ResultLevel.SUCCESS, msg)

    except Exception:
        record_result(log, ResultLevel.WARNING, "String conversion process failed")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()