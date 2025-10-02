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
bot_name = "AD - Move user to another OU"
data_to_log = {}
log.info("Static variables set")

def record_result(log, level, message):
    log.result_message(level, f"{bot_name}: {message}")

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

def move_user_ou(log, user_sam, target_ou):
    log.info(f"Moving AD user [{user_sam}] to OU [{target_ou}] via PowerShell")

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

    $ADOU = Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '{target_ou}'" -ErrorAction SilentlyContinue
    if (-not $ADOU) {{
        Write-Log "Error: The target OU [{target_ou}] does not exist" -IsError
        return
    }}

    $ADUser = Get-ADUser -Filter "sAMAccountName -eq '{user_sam}'" -ErrorAction SilentlyContinue
    if (-not $ADUser) {{
        Write-Log "Error: The user [{user_sam}] does not exist" -IsError
        return
    }}

    try {{
        Move-ADObject -Identity $ADUser.DistinguishedName -TargetPath $ADOU.DistinguishedName -ErrorAction Stop
        Write-Log "Success: Moved [{user_sam}] to [{target_ou}]"
    }} catch {{
        Write-Log "Error: Failed to move [{user_sam}] to [{target_ou}] $_" -IsError
    }}
    """

    success, output = execute_powershell(log, ps_command, shell="powershell")
    if success and "Success:" in output:
        log.info(f"User [{user_sam}] successfully moved to [{target_ou}]")
        return True, output
    else:
        log.error(f"Failed to move [{user_sam}] to [{target_ou}]: {output}")
        return False, output

def main():
    try:
        try:
            user_identifier = input.get_value("User_1743462131049")
            target_ou = input.get_value("TargetOU_1743462134209")
        except Exception:
            log.result_message(ResultLevel.WARNING, "Failed to fetch input values")
            return

        user_identifier = user_identifier.strip() if user_identifier else ""
        target_ou = target_ou.strip() if target_ou else ""

        log.info(f"Received input user = [{user_identifier}]")

        if not user_identifier:
            log.result_message(ResultLevel.WARNING, "User identifier is empty or invalid")
            return

        if not target_ou:
            log.result_message(ResultLevel.WARNING, "Target OU is empty or invalid")
            return

        user_result = get_user_data(log, user_identifier)

        if isinstance(user_result, list):
            details = "\n".join([f"- {u.get('displayName')} | {u.get('userPrincipalName')} | {u.get('id')}" for u in user_result])
            log.result_message(ResultLevel.WARNING, f"Multiple users found for [{user_identifier}]\n{details}")
            return

        user_email, user_sam = user_result
        if not user_email:
            log.result_message(ResultLevel.WARNING, f"Unable to resolve user principal name for [{user_identifier}]")
            return
        if not user_sam:
            log.result_message(ResultLevel.WARNING, f"No SAM account name found for [{user_identifier}] (cloud-only user)")
            return

        log.info(f"User [{user_identifier}] matched with UPN [{user_email}] and SAM [{user_sam}]")

        success, _ = move_user_ou(log, user_sam, target_ou)
        if success:
            log.result_message(ResultLevel.SUCCESS, f"User [{user_sam}] successfully moved to [{target_ou}]")
        else:
            log.result_message(ResultLevel.WARNING, f"Failed to move user [{user_sam}] to [{target_ou}]")

    except Exception:
        log.result_message(ResultLevel.WARNING, "Process failed")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()