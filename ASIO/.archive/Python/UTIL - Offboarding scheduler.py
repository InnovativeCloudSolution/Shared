import sys
import os
from datetime import datetime, timedelta, timezone
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

cwpsa_base_url = "https://au.myconnectwise.net/v4_6_release/apis/3.0"
vault_name = "mit-azu1-prod1-akv1"

data_to_log = {}
log.info("Static variables set")


def record_result(log, level, message):
    log.result_message(level, message)
    if level == ResultLevel.WARNING:
        data_to_log["Result"] = "Fail"
    elif level == ResultLevel.SUCCESS:
        if "Result" not in data_to_log:
            data_to_log["Result"] = "Success"


def convert_to_epoch_and_formats(log, raw_date, time_delta):
    try:
        log.info(f"Parsing raw date: {raw_date} with input timezone: {time_delta}")
        formats = ["%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%dT%H:%M", "%Y-%m-%d"]
        for fmt in formats:
            try:
                dt_naive = datetime.strptime(raw_date, fmt)
                break
            except ValueError:
                continue
        else:
            log.warning(f"Unrecognized date format: {raw_date}")
            return None, None, None, None

        input_offset = timedelta(hours=time_delta)
        input_tz = timezone(input_offset)

        dt_utc = dt_local.astimezone(timezone.utc)
        dt_local = dt_utc.astimezone(input_tz)

        epoch = int(dt_utc.timestamp())
        iso_utc = dt_utc.isoformat()
        iso_local = dt_local.isoformat()

        if fmt == "%Y-%m-%d":
            friendly = dt_local.strftime("%d/%m/%Y")
        else:
            friendly = dt_local.strftime("%d/%m/%Y %I:%M %p")

        return iso_utc, iso_local, friendly, epoch
    except Exception as e:
        log.exception(e, "Error in datetime conversion")
        return None, None, None, None

def check_input_date_is_today(log, str_local, time_delta=10):
    dt_now_local = datetime.now(timezone(timedelta(hours=time_delta)))
    log.info(f"Current system datetime with timedelta {time_delta}: {dt_now_local.isoformat()}")

    dt_local = datetime.fromisoformat(str_local)
    log.info(f"Input datetime with timedelta {time_delta}: {dt_local.isoformat()}")

    schedule_window_start = dt_now_local.replace(hour=0, minute=0, second=0, microsecond=0)
    schedule_window_end = dt_now_local.replace(hour=23, minute=59, second=59, microsecond=999999)

    if schedule_window_start <= dt_local <= schedule_window_end:
        log.info(f"Input date {dt_local.isoformat()} is within today's window: {schedule_window_start.isoformat()} to {schedule_window_end.isoformat()}")
        return "Yes"
    else:
        log.info(f"Input date {dt_local.isoformat()} is outside today's window: {schedule_window_start.isoformat()} to {schedule_window_end.isoformat()}")
        return "No"

def main():
    try:
        try:
            operation = input.get_value("Operation_1751001071319")
            date_time_raw = input.get_value("DateTime_1751001076863")
            time_delta = input.get_value("Timedelta_1750987527408", default=10)
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        operation = operation.strip() if operation else ""
        date_time_raw = date_time_raw.strip() if date_time_raw else ""
        time_delta = int(time_delta) if time_delta else 10
        log.info(f"Received inputs - operation: [{operation}]datetime: [{date_time_raw}], timedelta: [{time_delta}]")

        if not date_time_raw:
            record_result(log, ResultLevel.WARNING, "Datetime input is required")
            return
        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation input is required")
            return

        
        if operation == "Check if input date is today":
            sched_iso_utc, sched_iso_local, sched_friendly, sched_epoch = convert_to_epoch_and_formats(log, date_time_raw, time_delta)
            offboard_today = check_input_date_is_today(log, sched_iso_local, time_delta)
            if offboard_today == "Yes":
                record_result(log, ResultLevel.SUCCESS, "Input date is today")
                data_to_log["offboard_today"] = "Yes"
            else:
                record_result(log, ResultLevel.WARNING, "Input date is not today")
                data_to_log["offboard_today"] = "No"
            

        if operation == "Convert datetime to epoch and formats":
            sched_iso_utc, sched_iso_aest, sched_friendly, sched_epoch = convert_to_epoch_and_formats(log, date_time_raw, time_delta)
            if sched_epoch:
                data_to_log["sched_iso_utc"] = sched_iso_utc
                data_to_log["sched_iso_aest"] = sched_iso_aest
                data_to_log["sched_friendly_time_aest"] = sched_friendly
                data_to_log["sched_epoch"] = sched_epoch
                record_result(log, ResultLevel.SUCCESS, "Datetime converted successfully")
            else:
                record_result(log, ResultLevel.WARNING, "Failed to convert input datetime")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)


if __name__ == "__main__":
    main()
