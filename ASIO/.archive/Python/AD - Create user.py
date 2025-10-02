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
        "Initials", "Description", "Office", "Title", "Department", "Company",
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

def main():
    try:
        try:
            user_details = {
                "UserLogonName"             : input.get_value("UserLogonName_1743998828806"),
                "Password"                  : input.get_value("Password_1744056092639"),
                "Email"                     : input.get_value("Email_1743998875508"),
                "Domain"                    : input.get_value("DesiredUPNDomain_1743998889244"),
                "FirstName"                 : input.get_value("FirstName_1743998899162"),
                "LastName"                  : input.get_value("LastName_1743998912197"),
                "DisplayName"               : input.get_value("DisplayName_1743998931739"),
                "Initials"                  : input.get_value("Initials_1743998977523"),
                "Description"               : input.get_value("Description_1743999027211"),
                "Manager"                   : input.get_value("Manager_1743999053200"),
                "Title"                     : input.get_value("JobTitle_1744055916566"),
                "Department"                : input.get_value("Department_1744055920347"),
                "Company"                   : input.get_value("Company_1744055923880"),
                "Division"                  : input.get_value("Division_1744055970188"),
                "OrganizationalUnit"        : input.get_value("OrganizationalUnit_1743999068183"),
                "Office"                    : input.get_value("Office_1744055984598"),
                "StreetAddress"             : input.get_value("Street_1744055987546"),
                "POBox"                     : input.get_value("POBox_1744056009095"),
                "City"                      : input.get_value("City_1744055990479"),
                "State"                     : input.get_value("State_1744056005330"),
                "PostalCode"                : input.get_value("Zip_1744056248485"),
                "Country"                   : input.get_value("Country_1744056077740"),
                "HomePhone"                 : input.get_value("phone_1744056284223"),
                "MobilePhone"               : input.get_value("Mobile_1744056286371"),
                "Fax"                       : input.get_value("Fax_1744056326102"),
                "OfficePhone"               : input.get_value("OfficePhoneNumber_1744056315728"),
                "EmployeeID"                : input.get_value("EmployeeID_1744056507855"),
                "EmployeeNumber"            : input.get_value("EmployeeNumber_1744056318223"),
                "HomePage"                  : input.get_value("WebPage_1744056563329"),
                "ScriptPath"                : input.get_value("LogonScriptPath_1744056572131"),
                "ProfilePath"               : input.get_value("ProfilePath_1744056569625"),
                "HomeDrive"                 : input.get_value("HomeDriveLetter_1744056566969"),
                "HomeDirectory"             : input.get_value("HomeDirectory_1744056574419")
            }
        except Exception as e:
            log.exception(e, "Failed to fetch input values")
            log.result_message(ResultLevel.WARNING, "Failed to fetch input values")
            return

        for key in user_details:
            user_details[key] = user_details[key].strip() if user_details[key] else ""
        for key, value in user_details.items():
            if key.lower() == "password":
                continue
            if value:
                log.info(f"Field [{key}] has value [{value}]")
                log.result_message(ResultLevel.INFO, f"Field [{key}] has value [{value}]")
       
        log.info(f"Received input for user creation = [{user_details.get('UserLogonName')}]")

        success, _ = create_user(log, user_details)

        upn = f"{user_details.get('UserLogonName')}@{user_details.get('Domain')}"

        if success:
            log.result_message(ResultLevel.SUCCESS, f"AD user [{user_details.get('UserLogonName')}] created successfully")
            log.result_data({"UPN": upn, "Result": "Success"})
        else:
            log.result_message(ResultLevel.WARNING, f"Encountered issues while creating AD user [{user_details.get('UserLogonName')}]")
            log.result_data({"UPN": upn, "Result": "Fail"})

    except Exception:
        log.exception("An error occurred while processing")
        log.result_message(ResultLevel.WARNING, "Process failed with exception")
        log.result_data({"UPN": user_details.get("Email") or "Unknown", "Result": "Fail"})

if __name__ == "__main__":
    main()