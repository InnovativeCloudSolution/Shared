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

def clean_summary_description(text: str) -> str:
    cleaned = re.sub(r'^\s*\[Extended Summary\]\s*', '', text, flags=re.IGNORECASE)
    cleaned = re.sub(r'\s*\[Description\]\s*$', '', cleaned, flags=re.IGNORECASE)
    return cleaned.strip()

def remove_email_brackets(text: str) -> str:
    return re.sub(r'\s*<[^<>]+>', '', text).strip()

def main():
    try:
        input = Input()
        input_note = input.get_value("InputNote_1743484623974").strip()

        log.info(f"Received input note = [{input_note}]")

        if not input_note or not input_note.strip():
            log.error("Input note is empty or invalid")
            log.result_message(ResultLevel.FAILED, "Input note is empty or invalid")
            return

        cleaned_note = clean_summary_description(input_note)
        cleaned_note = remove_email_brackets(cleaned_note)
        cleaned_note = cleaned_note.strip().splitlines()[0]

        log.info(f"Final cleaned summary = [{cleaned_note}]")
        data_to_log = {"cleanedsummary": cleaned_note}
        log.result_data(data_to_log)
        log.result_message(ResultLevel.SUCCESS, f"Summary Cleaned: [{cleaned_note}]")

    except Exception:
        log.exception("An error occurred while processing")
        log.result_message(ResultLevel.FAILED, "Process failed")

if __name__ == "__main__":
    main()