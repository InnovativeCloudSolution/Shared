import sys
import re
import random
import time
import requests
import subprocess
import os
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

cwpsa_base_url = "https://au.myconnectwise.net/v4_6_release/apis/3.0"

data_to_log = {}
bot_name = "AD User(s) Management"
log.info("Static variables set")

def record_result(log, level, message):
    log.result_message(level, f"[{bot_name}]: {message}")
    if level == ResultLevel.WARNING:
        data_to_log["status_result"] = "Fail"
    elif level == ResultLevel.SUCCESS:
        if "status_result" not in data_to_log or data_to_log["status_result"] != "Fail":
            data_to_log["status_result"] = "Success"

def execute_api_call(log, http_client, method, endpoint, data=None, retries=5, integration_name=None, headers=None, params=None):
    base_delay = 5
    log.info(f"Executing API call: {method.upper()} {endpoint}")
    for attempt in range(retries):
        try:
            if integration_name:
                response = (
                    getattr(http_client.third_party_integration(integration_name), method)(url=endpoint, json=data)
                    if data else getattr(http_client.third_party_integration(integration_name), method)(url=endpoint)
                )
            else:
                request_args = {"url": endpoint}
                if params:
                    request_args["params"] = params
                if headers:
                    request_args["headers"] = headers
                if data:
                    if (headers and headers.get("Content-Type") == "application/x-www-form-urlencoded"):
                        request_args["data"] = data
                    else:
                        request_args["json"] = data
                response = getattr(requests, method)(**request_args)

            if 200 <= response.status_code < 300:
                return response
            elif response.status_code in [429, 503]:
                retry_after = response.headers.get("Retry-After")
                wait_time = int(retry_after) if retry_after else base_delay * (2 ** attempt) + random.uniform(0, 3)
                log.warning(f"Rate limit exceeded. Retrying in {wait_time:.2f} seconds")
                time.sleep(wait_time)
            elif 400 <= response.status_code < 500:
                if response.status_code == 404:
                    log.warning(f"Skipping non-existent resource [{endpoint}]")
                    return None
                log.error(f"Client error Status: {response.status_code}, Response: {response.text}")
                return response
            elif 500 <= response.status_code < 600:
                log.warning(f"Server error Status: {response.status_code}, attempt {attempt + 1} of {retries}")
                time.sleep(base_delay * (2 ** attempt) + random.uniform(0, 3))
            else:
                log.error(f"Unexpected response Status: {response.status_code}, Response: {response.text}")
                return response
        
        except Exception as e:
            log.exception(e, f"Exception during API call to {endpoint}")
            return None
    return None

def post_ticket_note(log, http_client, cwpsa_base_url, ticket_number, note_type, note):
    log.info(f"Posting {note_type} note to ticket [{ticket_number}]")
    note_endpoint = f"{cwpsa_base_url}/service/tickets/{ticket_number}/notes"
    payload = {
        "text": note,
        "detailDescriptionFlag": False,
        "internalAnalysisFlag": False,
        "resolutionFlag": False,
        "issueFlag": False,
        "internalFlag": False,
        "externalFlag": False,
        "contact": {
            "id": 15655
        }
    }
    if note_type == "discussion":
        payload["detailDescriptionFlag"] = True
    elif note_type == "internal":
        payload["internalAnalysisFlag"] = True
    note_response = execute_api_call(log, http_client, "post", note_endpoint, integration_name="cw_psa", data=payload)
    if note_response and note_response.status_code == 200:
        log.info(f"{note_type} note posted successfully to ticket [{ticket_number}]")
        return True
    else:
        log.error(f"Failed to post {note_type} note to ticket [{ticket_number}] Status: {note_response.status_code}, Body: {note_response.text}")
    return False

def execute_powershell(log, command, shell="pwsh", debug_mode=False, timeout=None, ignore_stderr_warnings=False, log_command=True, log_output=True):
    try:
        if shell not in ("pwsh", "powershell"):
            raise ValueError("Invalid shell specified. Use 'pwsh' or 'powershell'")
        if log_command:
            sanitized_cmd = command
            sanitized_cmd = re.sub(r"('|\")(eyJ[a-zA-Z0-9_-]{5,}?\.[a-zA-Z0-9_-]{5,}?\.([a-zA-Z0-9_-]{5,})?)\1", r"\1***TOKEN-MASKED***\1", sanitized_cmd)
            sanitized_cmd = re.sub(r"('|\")([a-zA-Z0-9]{8,}-(clientid|clientsecret|password))\1", r"\1***SECRET-MASKED***\1", sanitized_cmd, flags=re.IGNORECASE)
            log.info(f"Executing {shell} command: {sanitized_cmd[:100]}{'...' if len(sanitized_cmd) > 100 else ''}")
        if debug_mode:
            log.info(f"Debug mode enabled for PowerShell execution")
        result = subprocess.run([shell, "-Command", command], capture_output=True, text=True, timeout=timeout)
        stdout = result.stdout.strip()
        stderr = result.stderr.strip()
        if log_output and stdout and debug_mode:
            log.info(f"{shell} stdout: {stdout[:500]}{'...' if len(stdout) > 500 else ''}")
        if stderr and debug_mode:
            log.info(f"{shell} stderr: {stderr}")
        success = result.returncode == 0
        if success and stderr and ignore_stderr_warnings:
            error_patterns = ["error:", "exception:", "fatal:", "failed:"]
            if any(pattern in stderr.lower() for pattern in error_patterns):
                success = False
                log.warning(f"{shell} command returned warnings: {stderr}")
            else:
                log.info(f"{shell} command executed with warnings: {stderr}")
        if success:
            if not stderr:
                log.info(f"{shell} command executed successfully")
            return True, stdout
        else:
            error_message = stderr or stdout
            log.error(f"{shell} execution failed (returncode={result.returncode}): {error_message}")
            return False, error_message
    except subprocess.TimeoutExpired:
        log.error(f"{shell} command timed out after {timeout} seconds")
        return False, f"Command timed out after {timeout} seconds"
    except Exception as e:
        log.exception(e, f"Exception occurred during {shell} execution")
        return False, str(e)

def get_ad_user_data(log, user_identifier):
    log.info(f"Resolving user SMTP address and SAMAccountName for [{user_identifier}] via PowerShell")
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
                Write-Log "User [$UserIdentifier] matched with SMTP [$($MatchedUser.UserPrincipalName)] and SAM [$($MatchedUser.sAMAccountName)]"
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
        log.error(f"Failed to resolve user SMTP address and SAMAccountName for [{user_identifier}]")
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

def create_user(log, user_details):
    log.info(f"Creating AD user [{user_details.get('UserLogonName')}] via PowerShell")
    ps_userdetails = "@{" + "; ".join([f'"{k}" = "{v}"' for k, v in user_details.items()]) + "}"

    ps_command = f"""
    $ErrorActionPreference = 'Stop'

    function Write-Log {{
        param ([string]$Message)
        Write-Output "`n$($Message)"
    }}

    function Test-Placeholder {{
        param ([string]$Value)
        return $Value -match "^@.+@$"
    }}

    function Resolve-ManagerDN {{
        param ([string]$Identifier)
        try {{
            if (-not $Identifier) {{ return $null }}
            $Identifier = $Identifier.Trim()
            if ($Identifier -match '\\s') {{
                $Identifier = $Identifier -replace '\\s+', '*'
            }}
            $ADManagers = @(Get-ADUser -Filter "sAMAccountName -eq '$Identifier' -or UserPrincipalName -eq '$Identifier' -or mail -eq '$Identifier' -or DisplayName -like '*$Identifier*'" -Properties DistinguishedName -ErrorAction SilentlyContinue)
            if ($ADManagers.Count -eq 1) {{
                return $ADManagers[0].DistinguishedName
            }} elseif ($ADManagers.Count -gt 1) {{
                Write-Log "Warning: Multiple users matched for Manager [$Identifier]; skipping assignment"
            }} else {{
                Write-Log "Warning: No user matched for Manager [$Identifier]; skipping assignment"
            }}
        }} catch {{
            Write-Log "Error: Exception while resolving Manager [$Identifier] $_"
        }}
        return $null
    }}

    $UserDetails = {ps_userdetails}
    $OriginalDomain = $UserDetails.Domain

    try {{
        Import-Module ActiveDirectory -ErrorAction Stop
        $Domain = (Get-ADDomain).DNSRoot
        Get-ADDomain -Server $Domain -ErrorAction Stop
    }} catch {{
        Write-Log "Error: Failed to load AD module or domain info"
        throw
    }}

    $NameValue = if ($UserDetails.DisplayName) {{ $UserDetails.DisplayName }} else {{ "$($UserDetails.FirstName) $($UserDetails.LastName)" }}

    $Attributes = @{{
        Enabled                = $true
        ChangePasswordAtLogon  = $false
        SAMAccountName         = $UserDetails.UserLogonName
        UserPrincipalName      = "$($UserDetails.UserLogonName)@$Domain"
        AccountPassword        = ConvertTo-SecureString -String $UserDetails.Password -AsPlainText -Force
        Server                 = $Domain
        GivenName              = $UserDetails.FirstName
        Surname                = $UserDetails.LastName
        Name                   = $NameValue
        DisplayName            = $NameValue
        EmailAddress           = if ($UserDetails.Email) {{ $UserDetails.Email }} else {{ "$($UserDetails.UserLogonName)@$OriginalDomain" }}
    }}

    $optionalAttributes = @(
        "Description", "Office", "Title", "Department", "Company",
        "Division", "Manager", "OrganizationalUnit", "StreetAddress", "POBox", "City", "State",
        "PostalCode", "Country", "HomePhone", "MobilePhone", "Fax", "OfficePhone",
        "HomePage", "ScriptPath", "ProfilePath", "HomeDrive", "HomeDirectory",
        "EmployeeID", "EmployeeNumber"
    )

    foreach ($attr in $optionalAttributes) {{
        if ($UserDetails.ContainsKey($attr) -and $UserDetails[$attr] -and -not (Test-Placeholder $UserDetails[$attr])) {{
            $Attributes[$attr] = $UserDetails[$attr]
        }}
    }}

    $maxLengthFields = @("POBox", "Fax", "HomePhone", "MobilePhone", "OfficePhone", "PostalCode")
    foreach ($field in $maxLengthFields) {{
        if ($Attributes.ContainsKey($field) -and $Attributes[$field].Length -gt 16) {{
            Write-Log "Trimming value for [$field] to 16 characters"
            $Attributes[$field] = $Attributes[$field].Substring(0,16)
        }}
    }}

    $countryIsoMap = @{{
        "australia" = "AU"; "new zealand" = "NZ"; "united states" = "US"; "canada" = "CA"
        "united kingdom" = "GB"; "germany" = "DE"; "france" = "FR"; "india" = "IN"
        "singapore" = "SG"; "philippines" = "PH"
    }}
    if ($Attributes.ContainsKey("Country")) {{
        $countryInput = $Attributes["Country"].Trim()
        $normalizedInput = $countryInput.ToLower()
        if ($countryIsoMap.ContainsKey($normalizedInput)) {{
            $iso = $countryIsoMap[$normalizedInput]
            Write-Log "Converting Country from [$countryInput] to ISO [$iso]"
            $Attributes["Country"] = $iso
        }} elseif ($countryInput.Length -gt 2) {{
            Write-Log "Removing Country [$countryInput] (not ISO/mapped)"
            $Attributes.Remove("Country")
        }}
    }}

    if ($Attributes.ContainsKey("Manager")) {{
        $resolvedManager = Resolve-ManagerDN -Identifier $Attributes["Manager"]
        if ($resolvedManager) {{
            $Attributes["Manager"] = $resolvedManager
        }} else {{
            Write-Log "Removing invalid Manager [$($Attributes["Manager"])]"
            $Attributes.Remove("Manager")
        }}
    }}

    $Attributes.GetEnumerator() | ForEach-Object {{
        if ($_.Key -ne "AccountPassword") {{
            Write-Log "$($_.Key): $($_.Value)"
        }}
    }}

    try {{
        if ($Attributes.OrganizationalUnit) {{
            if (-not (Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$($Attributes.OrganizationalUnit)'" -ErrorAction SilentlyContinue)) {{
                Write-Log "Error: OU [$($Attributes.OrganizationalUnit)] doesn't exist"
                throw
            }}
        }}

        if ($Attributes.ContainsKey("OrganizationalUnit")) {{
            $Attributes["Path"] = $Attributes["OrganizationalUnit"]
            $Attributes.Remove("OrganizationalUnit")
        }}

        if ($Attributes.Count -eq 0) {{
            Write-Log "Error: No valid attributes to create user"
            throw
        }}

        $user = New-ADUser @Attributes -PassThru
        if (-not $user) {{
            Write-Log "Error: New-ADUser did not return a user object"
            throw
        }}

        Write-Log "Success: User [$($user.UserPrincipalName)] created in domain [$Domain]"

        if ($Domain -ne $OriginalDomain) {{
            try {{
                Set-ADUser -Identity $user.SamAccountName -UserPrincipalName "$($UserDetails.UserLogonName)@$OriginalDomain"
                Write-Log "Info: Updated UPN to [$($UserDetails.UserLogonName)@$OriginalDomain]"
            }} catch {{
                Write-Log "Error: Failed to update UPN to [$OriginalDomain]"
            }}
        }}

        $confirm = Get-ADUser -Identity $user.SamAccountName -Server $Domain
        if (-not $confirm) {{
            Write-Log "Error: User creation succeeded but could not confirm with Get-ADUser"
            throw
        }}

    }} catch {{
        Write-Log "Error: Failed to create user in [$Domain]. Exception: $_"
        throw
    }}
    """

    success, output = execute_powershell(log, ps_command, shell="powershell")
    if not success:
        log.error(f"Failed to create AD user [{user_details.get('UserLogonName')}]: {output}")
        return False, output

    log.info(f"Successfully created AD user [{user_details.get('UserLogonName')}]")
    return True, output

def disable_user(log, user_sam):
    log.info(f"Disabling AD user [{user_sam}] via PowerShell")

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

    Disable-ADAccount -Identity '{user_sam}'
    """

    success, output = execute_powershell(log, ps_command, shell="powershell")
    if not success:
        log.error(f"Failed to disable AD user [{user_sam}]: {output}")
        return False, output

    log.info(f"Successfully disabled AD user [{user_sam}]")
    return True, output

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

def hide_user_from_gal(log, user_sam):
    log.info(f"Hiding AD user [{user_sam}] from GAL via PowerShell")

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
        $User = Get-ADUser -Identity "{user_sam}" -Properties mailNickname, msExchHideFromAddressLists

        Set-ADUser -Identity $User.DistinguishedName -Replace @{{mailNickname="{user_sam}"}} -ErrorAction Stop
        Write-Log "Set mailNickname on [{user_sam}] to [{user_sam}]"

        Set-ADUser -Identity $User.DistinguishedName -Replace @{{msExchHideFromAddressLists=$true}} -ErrorAction Stop

        # Set additional attribute to hide from GAL (required for Sherrin)
        Set-ADUser -Identity $User.DistinguishedName -Replace @{{"msDS-cloudExtensionAttribute1"="HideFromGAL"}} -ErrorAction Stop
        
        Write-Log "Success: Hid [{user_sam}] from GAL"
    }} catch {{
        Write-Log "Error: Failed to update [{user_sam}] $_" -IsError
    }}
    """

    success, output = execute_powershell(log, ps_command, shell="powershell")
    if not success:
        log.error(f"Failed to hide AD user [{user_sam}] from GAL: {output}")
        return False, output

    log.info(f"Successfully hid AD user [{user_sam}] from GAL")
    return True, output

def main():
    try:
        try:
            cwpsa_ticket = input.get_value("TicketNumber_1763362009760")
            operation = input.get_value("Operation_1756856396879")
            user = input.get_value("User_1756858724784")
            username = input.get_value("Username_1743998828806")
            first_name = input.get_value("FirstName_1743998899162")
            last_name = input.get_value("LastName_1743998912197")
            password = input.get_value("Password_1744056092639")
            primary_smtp_address = input.get_value("PrimarySMTPAddress_1743998875508")
            display_name = input.get_value("DisplayName_1743998931739")
            domain = input.get_value("Domain_1743998889244")
            description = input.get_value("Description_1743999027211")
            manager = input.get_value("Manager_1743999053200")
            title = input.get_value("JobTitle_1744055916566")
            department = input.get_value("Department_1744055920347")
            company = input.get_value("Company_1744055923880")
            division = input.get_value("Division_1744055970188")
            organizational_unit = input.get_value("OrganizationalUnit_1743999068183")
            office = input.get_value("Office_1744055984598")
            street_address = input.get_value("Street_1744055987546")
            po_box = input.get_value("POBox_1744056009095")
            city = input.get_value("City_1744055990479")
            state = input.get_value("State_1744056005330")
            postal_code = input.get_value("Postcode_1744056248485")
            country = input.get_value("Country_1744056077740")
            business_phone = input.get_value("BusinessPhone_1744056284223")
            mobile_phone = input.get_value("Mobile_1744056286371")
            fax = input.get_value("Fax_1744056326102")
            office_phone = input.get_value("OfficePhoneNumber_1744056315728")
            employee_id = input.get_value("EmployeeID_1744056507855")
            employee_number = input.get_value("EmployeeNumber_1744056318223")
            home_page = input.get_value("WebPage_1744056563329")
            script_path = input.get_value("LogonScriptPath_1744056572131")
            profile_path = input.get_value("ProfilePath_1744056569625")
            home_drive = input.get_value("HomeDriveLetter_1744056566969")
            home_directory = input.get_value("HomeDirectory_1744056574419")
            post_discussion_note = input.get_value("PostDiscussionNote_1763360733719")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        cwpsa_ticket = cwpsa_ticket.strip() if cwpsa_ticket else ""
        operation = operation.strip() if operation else ""
        username = username.strip() if username else ""
        user = user.strip() if user else ""
        first_name = first_name.strip() if first_name else ""
        last_name = last_name.strip() if last_name else ""
        password = password.strip() if password else ""
        primary_smtp_address = primary_smtp_address.strip() if primary_smtp_address else ""
        if not display_name:
            display_name = f"{first_name} {last_name}"
        domain = domain.strip() if domain else ""
        description = description.strip() if description else ""
        manager = manager.strip() if manager else ""
        title = title.strip() if title else ""
        department = department.strip() if department else ""
        company = company.strip() if company else ""
        division = division.strip() if division else ""
        organizational_unit = organizational_unit.strip() if organizational_unit else ""
        office = office.strip() if office else ""
        street_address = street_address.strip() if street_address else ""
        po_box = po_box.strip() if po_box else ""
        city = city.strip() if city else ""
        state = state.strip() if state else ""
        postal_code = postal_code.strip() if postal_code else ""
        country = country.strip() if country else ""
        business_phone = business_phone.strip() if business_phone else ""
        mobile_phone = mobile_phone.strip() if mobile_phone else ""
        fax = fax.strip() if fax else ""
        office_phone = office_phone.strip() if office_phone else ""
        employee_id = employee_id.strip() if employee_id else ""
        employee_number = employee_number.strip() if employee_number else ""
        home_page = home_page.strip() if home_page else ""
        script_path = script_path.strip() if script_path else ""
        profile_path = profile_path.strip() if profile_path else ""
        home_drive = home_drive.strip() if home_drive else ""
        home_directory = home_directory.strip() if home_directory else ""

        log.info(f"Requested operation = [{operation}]")

        if not cwpsa_ticket:
            record_result(log, ResultLevel.WARNING, "Ticket number is required but missing")
            return

        if not operation:
            record_result(log, ResultLevel.WARNING, "Operation is required.")
            return

        if operation == "Create user":
            user_details = {
                "UserLogonName": username,
                "FirstName": first_name,
                "LastName": last_name,
                "Password": password,
                "Email": primary_smtp_address,
                "DisplayName": display_name,
                "Domain": domain,
                "Description": description,
                "Manager": manager,
                "Title": title,
                "Department": department,
                "Company": company,
                "Division": division,
                "OrganizationalUnit": organizational_unit,
                "Office": office,
                "StreetAddress": street_address,
                "POBox": po_box,
                "City": city,
                "State": state,
                "PostalCode": postal_code,
                "Country": country,
                "HomePhone": business_phone,
                "MobilePhone": mobile_phone,
                "Fax": fax,
                "OfficePhone": office_phone,
                "EmployeeID": employee_id,
                "EmployeeNumber": employee_number,
                "HomePage": home_page,
                "ScriptPath": script_path,
                "ProfilePath": profile_path,
                "HomeDrive": home_drive,
                "HomeDirectory": home_directory,
            }
            log.info(f"Received input for user creation = [{user_details}]")
            required = {
                "UserLogonName": user_details.get("UserLogonName"),
                "Password": user_details.get("Password"),
                "Domain": user_details.get("Domain"),
                "FirstName": user_details.get("FirstName"),
                "LastName": user_details.get("LastName"),
            }
            missing = [k for k, v in required.items() if not v]
            if missing:
                record_result(log, ResultLevel.WARNING, f"Missing required inputs: {', '.join(missing)}")
                return
            success, output = create_user(log, user_details)
            if success:
                record_result(log, ResultLevel.SUCCESS, f"AD user [{user_details.get('UserLogonName')}] created successfully")
                post_ticket_note(log, http_client, cwpsa_base_url, cwpsa_ticket, "discussion", f"AD user [{user_details.get('UserLogonName')}] created successfully") if post_discussion_note == "Yes" else None
            else:
                record_result(log, ResultLevel.WARNING, f"Encountered issues while creating AD user [{user_details.get('UserLogonName')}]")

        elif operation == "Disable user":
            if not user:
                record_result(log, ResultLevel.WARNING, "User identifier is required")
                return
            user_primary_smtp_address, user_sam = get_ad_user_data(log, user)
            if not user_primary_smtp_address:
                record_result(log, ResultLevel.WARNING, f"Unable to resolve user SMTP address for [{user}]")
                return
            if not user_sam:
                record_result(log, ResultLevel.WARNING, f"No SAM account name found for [{user}] (cloud-only user)")
                return
            success, output = disable_user(log, user_sam)
            if success:
                record_result(log, ResultLevel.SUCCESS, f"AD user [{user_sam}] disabled successfully")
                post_ticket_note(log, http_client, cwpsa_base_url, cwpsa_ticket, "discussion", f"AD user [{user_sam}] disabled successfully") if post_discussion_note == "Yes" else None
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to disable AD user [{user_sam}]")

        elif operation == "Hide user from GAL":
            if not user:
                record_result(log, ResultLevel.WARNING, "User identifier is required")
                return
            user_primary_smtp_address, user_sam = get_ad_user_data(log, user)
            if not user_primary_smtp_address:
                record_result(log, ResultLevel.WARNING, f"Unable to resolve user SMTP address for [{user}]")
                return
            if not user_sam:
                record_result(log, ResultLevel.WARNING, f"No SAM account name found for [{user}]")
                return
            success, output = hide_user_from_gal(log, user_sam)
            if success:
                record_result(log, ResultLevel.SUCCESS, f"AD user [{user_sam}] hidden from GAL successfully")
                post_ticket_note(log, http_client, cwpsa_base_url, cwpsa_ticket, "discussion", f"AD user [{user_sam}] hidden from GAL successfully") if post_discussion_note == "Yes" else None
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to hide AD user [{user_sam}] from GAL")

        elif operation == "Move user to another OU":
            if not user:
                record_result(log, ResultLevel.WARNING, "User identifier is required")
                return
            if not organizational_unit:
                record_result(log, ResultLevel.WARNING, "OrganizationalUnit is required")
                return
            user_primary_smtp_address, user_sam = get_ad_user_data(log, user)
            if not user_primary_smtp_address:
                record_result(log, ResultLevel.WARNING, f"Unable to resolve user SMTP address for [{user}]")
                return
            if not user_sam:
                record_result(log, ResultLevel.WARNING, f"No SAM account name found for [{user}] (cloud-only user)")
                return
            success, output = move_user_ou(log, user_sam, organizational_unit)
            if success:
                record_result(log, ResultLevel.SUCCESS, f"User [{user_sam}] successfully moved to [{organizational_unit}]")
                post_ticket_note(log, http_client, cwpsa_base_url, cwpsa_ticket, "discussion", f"User [{user_sam}] successfully moved to [{organizational_unit}]") if post_discussion_note == "Yes" else None
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to move user [{user_sam}] to [{organizational_unit}]")

        elif operation == "Reset user password":
            if not user:
                record_result(log, ResultLevel.WARNING, "User identifier is required")
                return
            if not password:
                record_result(log, ResultLevel.WARNING, "Password is required")
                return
            user_primary_smtp_address, user_sam = get_ad_user_data(log, user)
            if not user_primary_smtp_address:
                record_result(log, ResultLevel.WARNING, f"Unable to resolve user SMTP address for [{user}]")
                return
            if not user_sam:
                record_result(log, ResultLevel.WARNING, f"No SAM account name found for [{user}] (cloud-only user)")
                return
            success, output = reset_user_password(log, user_sam, password)
            if success:
                record_result(log, ResultLevel.SUCCESS, f"Password reset successfully for user [{user_sam}]")
                post_ticket_note(log, http_client, cwpsa_base_url, cwpsa_ticket, "discussion", f"Password reset successfully for user [{user_sam}]") if post_discussion_note == "Yes" else None
            else:
                record_result(log, ResultLevel.WARNING, f"Failed to reset password for user [{user_sam}]")

        else:
            record_result(log, ResultLevel.WARNING, f"Unsupported Operation [{operation}]")

    except Exception as e:
        record_result(log, ResultLevel.WARNING, f"Unhandled error in main: {str(e)}")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
