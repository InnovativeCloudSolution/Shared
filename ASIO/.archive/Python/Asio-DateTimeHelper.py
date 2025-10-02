import re
import subprocess
import os
import io
import json
import sys
import base64
import hashlib
import time
import urllib
import traceback
import random
import string
from datetime import datetime, timedelta, timezone
from cw_rpa import Logger, Input, HttpClient

log = Logger()
log.info("Imports completed successfully.")

def main():
    try:
        input = Input()
        datetime_input = input.get_value("DateTime_1737616719880")
        days_delta = int(input.get_value("Days_1737615762185") or 0)
        hours_delta = int(input.get_value("Hours_1737615763537") or 0)
        minutes_delta = int(input.get_value("Minutes_1737615764912") or 0)
        operation = input.get_value("Operation_1737620902560")
    except Exception as e:
        log.exception(e, "An error occurred while getting input values")
        log.result_failed_message("An error occurred while getting input values")
        return

    log.info(f"Processing datetime: {datetime_input}")

    try:
        if 'T' not in datetime_input and ' ' not in datetime_input:
            datetime_input += "T00:00"

        aest_timezone = timezone(timedelta(hours=10))
        parsed_datetime = datetime.strptime(datetime_input, "%Y-%m-%dT%H:%M") if 'T' in datetime_input else datetime.strptime(datetime_input, "%m/%d/%Y %I:%M %p")
        parsed_datetime = parsed_datetime.replace(tzinfo=aest_timezone)

        if operation and operation.strip().lower() in ["add", "subtract"]:
            delta = timedelta(days=days_delta, hours=hours_delta, minutes=minutes_delta)
            modified_datetime = parsed_datetime + delta if operation.lower() == "add" else parsed_datetime - delta
        else:
            modified_datetime = parsed_datetime

        formatted_datetime = modified_datetime.strftime("%Y-%m-%dT%H:%M")

        data_to_log = {"output": formatted_datetime, "datetime_object": modified_datetime}
        log.info(f"Modified Datetime: {data_to_log}")

        log.result_success_message(formatted_datetime)

    except ValueError as ve:
        log.exception(ve, "Invalid datetime format or operation provided.")
        log.result_failed_message("Invalid datetime format or operation provided.")
    except Exception as e:
        log.exception(e, "Unexpected error during datetime processing.")
        log.result_failed_message("Unexpected error during datetime processing.")

if __name__ == "__main__":
    main()
