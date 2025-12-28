import sys
import re
import subprocess
import os
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

data_to_log = {}
bot_name = "AD - User group(s) management"
log.info("Static variables set")

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

def get_ad_user_data(log, user_identifier):
    log.info(f"Resolving user UPN and SAMAccountName for [{user_identifier}] via PowerShell")

    ps_command = f"""
    $ErrorActionPreference = 'Stop'

    function Write-Log {{
        param([string]$Message, [switch]$IsError)
        if ($IsError) {{ Write-Error $Message }} else {{ Write-Output $Message }}
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
                Write-Log "Error: Multiple users found for identifier [$UserIdentifier]" -IsError
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

def add_user_to_groups(log, user_sam, user_groups):
    log.info(f"Adding [{user_sam}] to AD groups via PowerShell")

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

    $Groups = "{user_groups}" -split "," | ForEach-Object {{ $_.Trim() }}

    foreach ($Group in $Groups) {{
        $ADGroup = Get-ADGroup -Identity $Group -ErrorAction SilentlyContinue
        if (-not $ADGroup) {{
            Write-Log "Error: The group [$Group] does not exist" -IsError
            continue
        }}

        if ($Group -eq "Domain Users") {{
            Write-Log "Info: Skipping [$Group] as it's a default group"
            continue
        }}

        try {{
            Add-ADGroupMember -Identity $Group -Members "{user_sam}" -ErrorAction Stop
            Write-Log "Success: Added [{user_sam}] to [$Group]"
        }}
        catch {{
            Write-Log "Error: Failed to add [{user_sam}] to [$Group] $_" -IsError
        }}
    }}
    """

    success, output = execute_powershell(log, ps_command, shell="powershell")
    if not success:
        log.error(f"Failed to add user [{user_sam}] to groups")
        return False, output

    for line in output.strip().splitlines():
        if "Success: Added" in line:
            log.info(line)
        elif "Skipping" in line or "does not exist" in line:
            log.info(line)
        elif "Error:" in line:
            log.error(line)

    return True, output

def remove_user_from_groups(log, user_sam, user_groups):
    log.info(f"Removing [{user_sam}] from AD groups via PowerShell")

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

    $Groups = "{user_groups}" -split "," | ForEach-Object {{ $_.Trim() }}

    foreach ($Group in $Groups) {{
        $ADGroup = Get-ADGroup -Identity $Group -ErrorAction SilentlyContinue
        if (-not $ADGroup) {{
            Write-Log "Error: The group [$Group] does not exist" -IsError
            continue
        }}

        if ($Group -eq "Domain Users") {{
            Write-Log "Info: Skipping [$Group] as it's a default group"
            continue
        }}

        try {{
            Remove-ADGroupMember -Identity $Group -Members "{user_sam}" -Confirm:$false -ErrorAction Stop
            Write-Log "Success: Removed [{user_sam}] from [$Group]"
        }}
        catch {{
            Write-Log "Error: Failed to remove [{user_sam}] from [$Group] $_" -IsError
        }}
    }}
    """

    success, output = execute_powershell(log, ps_command, shell="powershell")
    if not success:
        log.error(f"Failed to remove user [{user_sam}] from groups")
        return False, output

    for line in output.strip().splitlines():
        if "Success: Removed" in line:
            log.info(line)
        elif "Skipping" in line or "does not exist" in line:
            log.info(line)
        elif "Error:" in line:
            log.error(line)

    return True, output

def remove_user_from_all_groups(log, user_sam):
    log.info(f"Removing AD user [{user_sam}] from groups via PowerShell")

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

    try {{
        $ADUser = Get-ADUser -Identity '{user_sam}' -Properties MemberOf -ErrorAction Stop
        if (-not $ADUser) {{
            Write-Log "Error: The user [{user_sam}] does not exist." -IsError
            return
        }}

        $UserGroups = @(Get-ADPrincipalGroupMembership -Identity '{user_sam}' | Select-Object -ExpandProperty Name)
        if (-not $UserGroups) {{
            Write-Log "Info: The user [{user_sam}] is not a member of any groups"
            return
        }}

        foreach ($Group in $UserGroups) {{
            $ADGroup = Get-ADGroup -Identity $Group -ErrorAction SilentlyContinue
            if (-not $ADGroup) {{
                Write-Log "Error: Group [$Group] does not exist or could not be retrieved" -IsError
                continue
            }}

            $GroupName = $ADGroup.Name

            if ($GroupName -eq "Domain Users") {{
                Write-Log "Info: Skipping default group [$GroupName]"
                continue
            }}

            try {{
                Remove-ADGroupMember -Identity $Group -Members '{user_sam}' -Confirm:$false -ErrorAction Stop
                Write-Log "Success: Removed [{user_sam}] from [$GroupName]"
            }} catch {{
                Write-Log "Error: Failed to remove [{user_sam}] from [$GroupName] $_" -IsError
            }}
        }}
    }} catch {{
        Write-Log "Error: Failed to retrieve or process group memberships for [{user_sam}] $_" -IsError
    }}
    """

    success, output = execute_powershell(log, ps_command, shell="powershell")
    if not success:
        log.error(f"Failed to remove AD user [{user_sam}] from groups: {output}")
        return False, output

    log.info(f"Successfully removed AD user [{user_sam}] from groups")
    return True, output

def main():
    try:
        try:
            operation = input.get_value("Operation_1744681896698")
            user_identifier = input.get_value("User_1743477061596")
            user_groups = input.get_value("Groups_1743477064940")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        operation = operation.strip() if operation else ""
        user_identifier = user_identifier.strip() if user_identifier else ""
        user_groups = user_groups.strip() if user_groups else ""

        log.info(f"Received input user = [{user_identifier}], operation = [{operation}]")

        if not user_identifier:
            record_result(log, ResultLevel.WARNING, "User identifier is empty or invalid")
            return
        if operation in ("Add", "Remove") and not user_groups:
            record_result(log, ResultLevel.WARNING, "User groups value is empty or invalid")
            return
        if operation not in ("Add", "Remove", "Remove from all groups"):
            record_result(log, ResultLevel.WARNING, "Operation input must be 'Add', 'Remove', or 'Remove from all groups'")
            return

        user_email, user_sam = get_ad_user_data(log, user_identifier)

        if not user_email:
            record_result(log, ResultLevel.WARNING, f"Unable to resolve user principal name for [{user_identifier}]")
            return
        if not user_sam:
            record_result(log, ResultLevel.WARNING, f"No SAM account name found for [{user_identifier}]")
            return

        log.info(f"User [{user_identifier}] matched with UPN [{user_email}] and SAM [{user_sam}]")

        if operation == "Add":
            success, output = add_user_to_groups(log, user_sam, user_groups)
            if success:
                lines = output.splitlines()
                added_groups = [line for line in lines if line.startswith("Success: Added")]
                failed_groups = [line for line in lines if "Error: Failed to add" in line]

                if added_groups:
                    log.info("Added to groups:\n" + "\n".join(added_groups))
                    data_to_log["AddedGroups"] = added_groups
                    for group in added_groups:
                        record_result(log, ResultLevel.SUCCESS, group) 
                elif failed_groups:
                    log.error("Some groups could not be added:\n" + "\n".join(failed_groups))
                    record_result(log, ResultLevel.WARNING, f"Failed to add AD user [{user_sam}] to one or more groups")
                    return
                else:
                    record_result(log, ResultLevel.SUCCESS, f"AD user [{user_sam}] was not added to any groups")
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to add AD user [{user_sam}] to groups")
                return

        if operation == "Remove":
            success, output = remove_user_from_groups(log, user_sam, user_groups)
            if success:
                lines = output.splitlines()
                removed_groups = [line for line in lines if line.startswith("Success: Removed")]
                failed_groups = [line for line in lines if "Error: Failed to remove" in line]

                if removed_groups:
                    log.info("Removed from groups:\n" + "\n".join(removed_groups))
                    data_to_log["RemovedGroups"] = removed_groups
                    for group in removed_groups:
                        record_result(log, ResultLevel.SUCCESS, group) 
                elif failed_groups:
                    log.error("Some groups could not be removed:\n" + "\n".join(failed_groups))
                    record_result(log, ResultLevel.WARNING, f"Failed to remove AD user [{user_sam}] from one or more groups")
                    return
                else:
                    record_result(log, ResultLevel.SUCCESS, f"AD user [{user_sam}] had no removable group memberships")
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to remove AD user [{user_sam}] from groups")
                return

        if operation == "Remove from all groups":
            success, output = remove_user_from_all_groups(log, user_sam)
            if success:
                lines = output.splitlines()
                removed_groups = [line for line in lines if line.startswith("Success: Removed")]
                failed_groups = [line for line in lines if "Error: Failed to remove" in line]

                if removed_groups:
                    log.info("Removed from groups:\n" + "\n".join(removed_groups))
                    data_to_log["RemovedGroups"] = removed_groups
                    for group in removed_groups:
                        record_result(log, ResultLevel.SUCCESS, group) 
                elif failed_groups:
                    log.error("Some groups could not be removed:\n" + "\n".join(failed_groups))
                    record_result(log, ResultLevel.WARNING, f"Failed to remove AD user [{user_sam}] from one or more groups")
                    return
                else:
                    record_result(log, ResultLevel.SUCCESS, f"AD user [{user_sam}] had no removable group memberships")
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to remove AD user [{user_sam}] from all groups")
                return

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()