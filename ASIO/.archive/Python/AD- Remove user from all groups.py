import re
import sys
import subprocess
import os
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

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
            if ($UserIdentifier -match '\\s') {{
                $UserIdentifier = $UserIdentifier -replace '\\s+', '*'
            }}
            $ADUsers = @(Get-ADUser -Filter "sAMAccountName -eq '$UserIdentifier' -or UserPrincipalName -eq '$UserIdentifier' -or mail -eq '$UserIdentifier' -or DisplayName -like '*$UserIdentifier*'" -Properties sAMAccountName, UserPrincipalName -ErrorAction SilentlyContinue)

            if ($ADUsers.Count -eq 1) {{
                $MatchedUser = $ADUsers[0]
                Write-Log "User [$UserIdentifier] matched with UPN [$($MatchedUser.UserPrincipalName)] and SAM [$($MatchedUser.sAMAccountName)]"
                return @($MatchedUser.UserPrincipalName, $MatchedUser.sAMAccountName)
            }} elseif ($ADUsers.Count -gt 1) {{
                Write-Log "Error: Multiple users found for identifier [$UserIdentifier] unable to determine the correct account" -IsError
                return $null
            }}
        }} catch {{
            Write-Log "Error: Failed to search for user with identifier [$UserIdentifier] $_" -IsError
        }}
        return $null
    }}

    try {{
        Import-Module ActiveDirectory -ErrorAction Stop
    }} catch {{
        Write-Log "Error: Failed to import the ActiveDirectory module" -IsError
        return
    }}

    Get-UserAccount -UserIdentifier '{user_identifier}'
    """

    success, output = execute_powershell(log, ps_command, shell="powershell")
    if not success or not output or "Error:" in output:
        log.error(f"Failed to resolve user UPN and SAMAccountName for [{user_identifier}]")
        return "", ""

    try:
        user_upn, user_sam = [line.strip() for line in output.strip().splitlines() if line.strip()][-2:]
        return user_upn, user_sam
    except Exception:
        log.error(f"Unexpected output format while resolving user for [{user_identifier}]")
        return "", ""

def remove_user_from_groups(log, user_sam):
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
            user_identifier = input.get_value("User_1743425589488")
        except Exception as e:
            log.exception(e, "Failed to fetch input values")
            log.result_message(ResultLevel.FAILED, "Failed to fetch input values")
            return
        
        user_identifier = user_identifier.strip() if user_identifier else ""

        log.info(f"Received input user = [{user_identifier}]")

        if not user_identifier:
            log.error("User identifier is empty or invalid")
            log.result_message(ResultLevel.FAILED, "User identifier is empty or invalid")
            return

        user_result = get_user_data(log, user_identifier)

        if isinstance(user_result, list):
            details = "\n".join([f"- {u.get('displayName')} | {u.get('userPrincipalName')} | {u.get('id')}" for u in user_result])
            log.result_message(ResultLevel.FAILED, f"Multiple users found for [{user_identifier}]\n{details}")
            return

        user_email, user_sam = user_result
        if not user_email:
            log.result_message(ResultLevel.FAILED, f"Unable to resolve user principal name for [{user_identifier}]")
            return
        if not user_sam:
            log.result_message(ResultLevel.FAILED, f"No SAM account name found for [{user_identifier}] (cloud-only user)")
            return

        log.info(f"User [{user_identifier}] matched with UPN [{user_email}] and SAM [{user_sam}]")

        success, output = remove_user_from_groups(log, user_sam)
        if success:
            removed_groups = re.findall(r"Success: Removed \[.*?\] from \[.*?\]", output)
            if removed_groups:
                removed_groups_text = "\n".join(removed_groups)
                log.info(f"Removed from groups:\n{removed_groups_text}")
                log.result_message(ResultLevel.SUCCESS, f"AD user [{user_sam}] removed from {len(removed_groups)} groups successfully:\n{removed_groups_text}")
            else:
                log.result_message(ResultLevel.SUCCESS, f"AD user [{user_sam}] had no removable group memberships")
        else:
            log.result_message(ResultLevel.FAILED, f"Failed to remove AD user [{user_sam}] from groups")

    except Exception:
        log.exception("An error occurred while processing")
        log.result_message(ResultLevel.FAILED, "Process failed")

if __name__ == "__main__":
    main()