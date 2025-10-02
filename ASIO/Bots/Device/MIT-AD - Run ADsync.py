import sys
import subprocess
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from cw_rpa import Logger, Input, HttpClient, ResultLevel

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

data_to_log = {}
bot_name = "MIT-AD - Run ADsync"

def record_result(log, level, message):
    log.result_message(level, f"[{bot_name}]: {message}")
    if level == ResultLevel.WARNING:
        data_to_log["status_result"] = "Fail"
    elif level == ResultLevel.SUCCESS:
        if "status_result" not in data_to_log or data_to_log["status_result"] != "Fail":
            data_to_log["status_result"] = "Success"

def execute_powershell(log, command, shell="pwsh"):
    try:
        if shell not in ("pwsh", "powershell"):
            raise ValueError("Invalid shell specified. Use 'pwsh' or 'powershell'")

        result = subprocess.run([shell, "-Command", command], capture_output=True, text=True)

        if result.returncode == 0 and not result.stderr.strip():
            log.info(f"{shell} command executed successfully")
            return True, result.stdout.strip()
        else:
            error_message = result.stderr.strip() or result.stdout.strip()
            log.error(f"{shell} execution failed: {error_message}")
            return False, error_message

    except Exception as e:
        log.exception(e, f"Exception occurred during {shell} execution")
        return False, str(e)

def check_adsync(log):
    log.info("Checking for ADSync module availability")
    check_command = "Get-Module -ListAvailable -Name ADSync"
    success, output = execute_powershell(log, check_command, shell="powershell")
    log.info(f"ADSync module check output: {output}")
    if not success or "ADSync" not in output:
        log.error("ADSync module is not available. This server must have Azure AD Connect installed")
        return False
    return True

def trigger_adsync(log, policy_type="Delta"):
    log.info(f"Triggering Azure AD Connect {policy_type} sync")
    ps_command = f"""
    $ErrorActionPreference = 'Stop'
    function Write-Log {{
        param([string]$Message, [switch]$IsError)
        if ($IsError) {{
            Write-Error $Message
        }} else {{
            Write-Output $Message
        }}
    }}
    try {{
        Import-Module ADSync -ErrorAction Stop
    }} catch {{
        Write-Log "Error: Failed to import the ADSync module" -IsError
        return
    }}
    try {{
        Start-ADSyncSyncCycle -PolicyType {policy_type}
        Write-Output "{policy_type} sync triggered successfully"
    }} catch {{
        Write-Error "Error: Failed to trigger {policy_type} sync - $_"
    }}
    """
    success, output = execute_powershell(log, ps_command, shell="powershell")
    log.info(f"PowerShell execution output: {output}")

    if success and f"{policy_type} sync triggered successfully" in output:
        log.info(f"Azure AD Connect {policy_type} sync completed successfully")
        return True
    if "Failed to import the ADSync module" in output:
        log.error("Azure AD Connect sync failed: ADSync module not found")
        return False
    if "Sync is already running" in output:
        log.info("Azure AD Connect sync is already running; skipping new trigger")
        return True  # or return False if you want to mark as non-success

    log.error(f"Azure AD Connect {policy_type} sync failed")
    return False

def main():
    try:
        try:
            policy_type = input.get_value("PolicyType_1743472020267")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        policy_type = policy_type.strip() if policy_type else ""
        log.info(f"Received policy type input = [{policy_type}]")

        if not policy_type:
            record_result(log, ResultLevel.WARNING, "Policy type is empty or invalid")
            return

        if not check_adsync(log):
            record_result(log, ResultLevel.WARNING, "Azure AD Connect is not installed (ADSync module missing)")
            return

        log.info(f"Starting Azure AD Connect {policy_type} sync process")
        success = trigger_adsync(log, policy_type)
        log.info(f"Completed Azure AD Connect {policy_type} sync process")

        if success:
            record_result(log, ResultLevel.SUCCESS, f"Azure AD Connect {policy_type} sync executed successfully")
        else:
            record_result(log, ResultLevel.WARNING, f"Azure AD Connect {policy_type} sync failed")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
