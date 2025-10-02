import sys
import subprocess
import os
import re
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

data_to_log = {}
log.info("Static variables set")

def record_result(log, level, message):
    log.result_message(level, message)

    if level == ResultLevel.WARNING:
        data_to_log["Result"] = "Fail"
    elif level == ResultLevel.SUCCESS:
        if "Result" not in data_to_log:
            data_to_log["Result"] = "Success"

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

def get_user_data(log, user_identifier):
    log.info(f"Resolving user UPN and SAMAccountName for [{user_identifier}] via PowerShell")

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

    function Get-UserAccount {{
        param ([string]$UserIdentifier)

        try {{
            $UserIdentifier = $UserIdentifier.Trim()

            Import-Module ActiveDirectory -ErrorAction Stop

            $Filter = "sAMAccountName -eq '$UserIdentifier' -or UserPrincipalName -eq '$UserIdentifier' -or mail -eq '$UserIdentifier' -or DisplayName -like '*$UserIdentifier*' -or Name -like '*$UserIdentifier*'"

            $ADUsers = @(Get-ADUser -Filter $Filter -Properties sAMAccountName, UserPrincipalName -ErrorAction SilentlyContinue)

            if ($ADUsers.Count -eq 1) {{
                $MatchedUser = $ADUsers[0]
                Write-Log "User [$UserIdentifier] matched with UPN [$($MatchedUser.UserPrincipalName)] and SAM [$($MatchedUser.sAMAccountName)]"
                return "$($MatchedUser.UserPrincipalName)`n$($MatchedUser.sAMAccountName)"
            }} elseif ($ADUsers.Count -gt 1) {{
                Write-Log "Error: Multiple users found for identifier [$UserIdentifier] unable to determine the correct account" -IsError
                return $null
            }} else {{
                Write-Log "Error: No users found for [$UserIdentifier]" -IsError
                return $null
            }}
        }} catch {{
            Write-Log "Error: Exception while searching for [$UserIdentifier] $_" -IsError
            return $null
        }}
    }}

    Get-UserAccount -UserIdentifier '{user_identifier}'
    """

    success, output = execute_powershell(log, ps_command, shell="powershell")
    if not success or not output or "Error:" in output:
        log.error(f"Failed to resolve user UPN and SAMAccountName for [{user_identifier}]")
        return "", ""

    log.info(f"Raw PowerShell output:\n{output}")

    try:
        matches = [line.strip() for line in output.strip().splitlines() if "@" in line or re.match(r"^[a-zA-Z0-9_.-]+$", line.strip())]
        if len(matches) >= 2:
            return matches[-2], matches[-1]
        log.error(f"Unexpected PowerShell output while resolving user for [{user_identifier}]")
        return "", ""
    except Exception:
        log.error(f"Exception while parsing PowerShell output for [{user_identifier}]")
        return "", ""

def reset_user_password(log, user_sam, password):
    log.info(f"Resetting password for AD user [{user_sam}] via PowerShell")

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
        Import-Module ActiveDirectory -ErrorAction Stop
    }} catch {{
        Write-Log "Error: Failed to import the ActiveDirectory module" -IsError
        return
    }}

    $ADUser = Get-ADUser -Filter "sAMAccountName -eq '{user_sam}'" -ErrorAction SilentlyContinue
    if (-not $ADUser) {{
        Write-Log "Error: The user [{user_sam}] does not exist" -IsError
        return
    }}

    try {{
        $SecurePassword = ConvertTo-SecureString -String '{password}' -AsPlainText -Force
        Set-ADAccountPassword -Identity $ADUser.DistinguishedName -NewPassword $SecurePassword -Reset -ErrorAction Stop
        Write-Log "Success: Password reset for user [{user_sam}]"
    }} catch {{
        Write-Log "Error: Failed to reset password for user [{user_sam}] $_" -IsError
    }}
    """

    success, output = execute_powershell(log, ps_command, shell="powershell")
    if not success:
        log.error(f"Password reset failed for [{user_sam}]: {output}")
        return False, output

    log.info(f"Password reset successful for [{user_sam}]")
    return True, output

def main():
    try:
        try:
            user_identifier = input.get_value("User_1751405954534")
            password = input.get_value("Password_1751405958105")
        except:
            log.error("Failed to fetch input values")
            log.result_message(ResultLevel.WARNING, "Failed to fetch input values")
            return

        user_identifier = user_identifier.strip() if user_identifier else ""
        password = password.strip() if password else ""

        log.info(f"Received input user = [{user_identifier}]")

        if not user_identifier:
            log.error("User identifier is empty or invalid")
            log.result_message(ResultLevel.WARNING, "User identifier is empty or invalid")
            return

        if not password:
            log.error("Password input is empty or invalid")
            log.result_message(ResultLevel.WARNING, "Password input is empty or invalid")
            return

        user_email, user_sam = get_user_data(log, user_identifier)

        if not user_email:
            log.result_message(ResultLevel.WARNING, f"Unable to resolve user principal name for [{user_identifier}]")
            return
        if not user_sam:
            log.result_message(ResultLevel.WARNING, f"No SAM account name found for [{user_identifier}] (cloud-only user)")
            return

        log.info(f"User [{user_identifier}] matched with UPN [{user_email}] and SAM [{user_sam}]")

        success, _ = reset_user_password(log, user_sam, password)

        data_to_log["password"] = password
        if success:
            log.result_message(ResultLevel.SUCCESS, f"Password reset successfully for user [{user_sam}]")
            data_to_log["Result"] = "Success"
        else:
            log.result_message(ResultLevel.WARNING, f"Failed to reset password for user [{user_sam}]")
            data_to_log["Result"] = "Fail"
            
    except Exception:
        record_result(log, ResultLevel.WARNING, "Process failed")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()