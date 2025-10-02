import os
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo
from cw_rpa import Logger, Input, HttpClient, ResultLevel

import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

vault_name = "mit-azu1-prod1-akv1"

data_to_log = {}
bot_name = "MIT-UTIL - DateTime helper"
log.info("Static variables set")

def record_result(log, level, message):
    log.result_message(level, f"[{bot_name}]: {message}")
    if level == ResultLevel.WARNING:
        data_to_log["status_result"] = "Fail"
    elif level == ResultLevel.SUCCESS:
        if "status_result" not in data_to_log or data_to_log["status_result"] != "Fail":
            data_to_log["status_result"] = "Success"

def convert_to_epoch_and_formats(log, raw_date, country, city, add_minutes):
    try:
        tz = f"{country}/{city}" if city else country
        tzinfo = ZoneInfo(tz)
        log.info(f"Using timezone: {tz}")

        dt_local = None
        dt_friendly = None

        if raw_date.isdigit():
            dt_local = datetime.fromtimestamp(int(raw_date) / 1000, tz=tzinfo)
            dt_friendly = dt_local
        else:
            formats = [
                ("%Y-%m-%dT%H:%M:%SZ", ZoneInfo("UTC")),
                ("%Y-%m-%dT%H:%M:%S", tzinfo),
                ("%Y-%m-%dT%H:%M", tzinfo),
                ("%Y-%m-%d", tzinfo)
            ]
            for fmt, source_tz in formats:
                try:
                    dt_naive = datetime.strptime(raw_date, fmt)
                    dt_local = dt_naive.replace(tzinfo=source_tz).astimezone(tzinfo)
                    dt_friendly = dt_local
                    break
                except ValueError:
                    continue
            else:
                log.warning(f"Unrecognized date format: {raw_date}")
                return {}

        dt_utc = dt_local.astimezone(ZoneInfo("UTC"))
        original_epoch = int(dt_utc.timestamp() * 1000)
        friendly = dt_friendly.strftime("%d/%m/%Y %I:%M %p")
        cwpsa_friendly = dt_friendly.astimezone(ZoneInfo("UTC")).strftime("%Y-%m-%dT%H:%M:%SZ")

        dt_offset_local = dt_local + timedelta(minutes=add_minutes)
        dt_offset_utc = dt_offset_local.astimezone(ZoneInfo("UTC"))
        friendly_offset = dt_offset_local.strftime("%d/%m/%Y %I:%M %p")
        cwpsa_offset = dt_offset_local.astimezone(ZoneInfo("UTC")).strftime("%Y-%m-%dT%H:%M:%SZ")

        dt_now_local = datetime.now(tzinfo)
        dt_now_utc = dt_now_local.astimezone(ZoneInfo("UTC"))
        friendly_now = dt_now_local.strftime("%d/%m/%Y %I:%M %p")
        cwpsa_now = dt_now_local.astimezone(ZoneInfo("UTC")).strftime("%Y-%m-%dT%H:%M:%SZ")

        return {
            "original_local": dt_local.strftime("%Y-%m-%dT%H:%M:%S"),
            "original_utc": dt_utc.strftime("%Y-%m-%dT%H:%M:%S"),
            "original_friendly_time": friendly,
            "original_cwpsa_friendly": cwpsa_friendly,
            "original_epoch": original_epoch,
            "offset_local": dt_offset_local.strftime("%Y-%m-%dT%H:%M:%S"),
            "offset_utc": dt_offset_utc.strftime("%Y-%m-%dT%H:%M:%S"),
            "offset_friendly_time": friendly_offset,
            "offset_cwpsa_friendly": cwpsa_offset,
            "now_local": dt_now_local.strftime("%Y-%m-%dT%H:%M:%S"),
            "now_utc": dt_now_utc.strftime("%Y-%m-%dT%H:%M:%S"),
            "now_friendly_time": friendly_now,
            "now_cwpsa_friendly": cwpsa_now,
            "datetime_format": "yyyy-MM-dd'T'HH:mm:ss",
            "datetime_tz": tz
        }
    except Exception as e:
        log.exception(e, "Error in datetime conversion")
        return {}

def check_input_date_is_today(log, raw_date_local, country, city):
    tz_local = ZoneInfo(f"{country}/{city}" if city else country)
    log.info(f"Using local timezone: {tz_local}")

    dt_now_local = datetime.now(tz_local)
    log.info(f"Current local datetime: {dt_now_local.isoformat()}")

    try:
        dt_sched_naive = datetime.strptime(raw_date_local, "%Y-%m-%dT%H:%M:%S")
        dt_sched_local = dt_sched_naive.replace(tzinfo=tz_local)
    except ValueError:
        log.warning(f"Could not parse input datetime as local: {raw_date_local}")
        return False

    log.info(f"Input local datetime: {dt_sched_local.isoformat()}")

    start = dt_now_local.replace(hour=0, minute=0, second=0, microsecond=0)
    end = dt_now_local.replace(hour=23, minute=59, second=59, microsecond=999999)

    return start <= dt_sched_local <= end

def check_input_date_in_past(log, raw_date_local, country, city):
    tz_local = ZoneInfo(f"{country}/{city}" if city else country)
    log.info(f"Using local timezone: {tz_local}")

    dt_now_local = datetime.now(tz_local)
    log.info(f"Current local datetime: {dt_now_local.isoformat()}")

    try:
        dt_sched_naive = datetime.strptime(raw_date_local, "%Y-%m-%dT%H:%M:%S")
        dt_sched_local = dt_sched_naive.replace(tzinfo=tz_local)
    except ValueError:
        log.warning(f"Could not parse input datetime as local: {raw_date_local}")
        return False

    log.info(f"Input local datetime: {dt_sched_local.isoformat()}")
    return dt_sched_local < dt_now_local

def main():
    try:
        try:
            operation = input.get_value("Operation_1756870118168")
            date_time_raw = input.get_value("RawDateTime_1756870185044")
            country = input.get_value("TimezoneCountry_1756870197200")
            city = input.get_value("TimezoneCity_1756870198401")
            add_minutes_raw = input.get_value("AddMinutes_1756870195889")
        except Exception:
            record_result(log, ResultLevel.WARNING, "WARNING: Failed to fetch input values.")
            return

        operation = operation.strip() if operation else ""
        date_time_raw = date_time_raw.strip() if date_time_raw else ""
        add_minutes = int(add_minutes_raw.strip()) if add_minutes_raw and add_minutes_raw.strip().isdigit() else 0
        country = country.strip() if country else "Australia"
        city = city.strip() if city else "Brisbane"
        log.info(f"Inputs - operation: [{operation}], datetime: [{date_time_raw}], country: [{country}], city: [{city}], add_minutes: [{add_minutes}]")

        if not date_time_raw or not operation:
            record_result(log, ResultLevel.WARNING, "Datetime and operation inputs are required")
            return

        datetime_data = convert_to_epoch_and_formats(log, date_time_raw, country, city, add_minutes)

        if not datetime_data or not all(datetime_data.get(k) for k in ["original_local", "original_utc"]):
            record_result(log, ResultLevel.WARNING, "Datetime conversion failed. Invalid or unrecognized input format.")
            return

        tz = ZoneInfo(f"{country}/{city}" if city else country)
        data_to_log.update(datetime_data)

        if operation == "Check if input date is today or in the past":
            if check_input_date_in_past(log, datetime_data["original_local"], country, city):
                data_to_log["datetime_in_past"] = "Yes"
                data_to_log["datetime_today"] = "No"
                now_offset_utc = datetime.now(tz) + timedelta(minutes=add_minutes)
                epoch = int(now_offset_utc.astimezone(ZoneInfo("UTC")).timestamp() * 1000)
                data_to_log["epoch"] = epoch
                if (add_minutes > 0):
                    record_result(log, ResultLevel.SUCCESS, f"Input datetime is in the past. Added {add_minutes} minutes to the input time.")
                else:
                    record_result(log, ResultLevel.SUCCESS, f"Input datetime is in the past. No minutes added to the input time.")
                log.info(data_to_log)
                return
            if check_input_date_is_today(log, datetime_data["original_local"], country, city):
                data_to_log["datetime_today"] = "Yes"
                data_to_log["datetime_in_past"] = "No"
                epoch = int(datetime.strptime(datetime_data["original_utc"], "%Y-%m-%dT%H:%M:%S").timestamp() * 1000)
                data_to_log["epoch"] = epoch
                if (add_minutes > 0):
                    record_result(log, ResultLevel.SUCCESS, f"Input datetime is today. Added {add_minutes} minutes to the input time.")
                else:
                    record_result(log, ResultLevel.SUCCESS, f"Input datetime is today. No minutes added to the input time.")
                log.info(data_to_log)
            else:
                data_to_log["datetime_today"] = "No"
                data_to_log["datetime_in_past"] = "No"
                now_offset_utc = datetime.now(tz) + timedelta(minutes=add_minutes)
                epoch = int(datetime.strptime(datetime_data["original_utc"], "%Y-%m-%dT%H:%M:%S").timestamp() * 1000)
                data_to_log["epoch"] = epoch
                if (add_minutes > 0):
                    record_result(log, ResultLevel.INFO, f"Input date is not today. Added {add_minutes} minutes to the input time.")
                else:
                    record_result(log, ResultLevel.INFO, f"Input date is not today. No minutes added to the input time.")
                log.info(data_to_log)

        elif operation == "Convert datetime to epoch and formats":
            epoch = int(datetime.strptime(datetime_data["original_utc"], "%Y-%m-%dT%H:%M:%S").timestamp() * 1000)
            data_to_log["epoch"] = epoch
            log.info(data_to_log)
            record_result(log, ResultLevel.SUCCESS, "Datetime converted successfully")

        else:
            record_result(log, ResultLevel.WARNING, f"Unknown operation: {operation}")
            return

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()