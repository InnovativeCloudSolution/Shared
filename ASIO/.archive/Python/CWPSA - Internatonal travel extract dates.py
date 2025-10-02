import sys
import traceback
import json
import random
import re
import subprocess
import os
import io
import base64
import hashlib
import time
import urllib.parse
import string
from datetime import datetime, timedelta, timezone

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from cw_rpa import Logger, Input, HttpClient, ResultLevel

log = Logger()
http_client = HttpClient()
log.info("Imports completed successfully")

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
        input = Input()
        input_summary = input.get_value("Summary_1743556477136").strip()

        log.info(f"Received input summary = [{input_summary}]")

        if not input_summary:
            log.error("Input summary is empty or invalid")
            log.result_message(ResultLevel.FAILED, "Input summary is empty or invalid")
            return

        try:
            start_formatted, end_formatted, start_string, end_string = extract_start_end_datetime(input_summary)

            log.info(f"Extracted start = [{start_formatted}], end = [{end_formatted}]")

            data_to_log = {
                "start": start_string,
                "end": end_string
            }
            log.result_data(data_to_log)

            log.result_message(ResultLevel.SUCCESS, f"Start: {start_formatted}, End: {end_formatted}")
        
        except ValueError as ve:
            log.warning(f"Datetime extraction failed: {ve}")
            log.result_message(ResultLevel.FAILED, "Unable to extract start and end datetime")

    except Exception:
        log.exception("An error occurred while extracting datetime")
        log.result_message(ResultLevel.FAILED, "Process failed")

if __name__ == "__main__":
    main()