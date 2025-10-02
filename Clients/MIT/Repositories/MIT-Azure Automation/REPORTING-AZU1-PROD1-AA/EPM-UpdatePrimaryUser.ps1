<#

Mangano IT - Endpoint Manager - Update Primary User
Created by: Gabriel Nugent
Version: 1.1.5

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [int]$TicketId,
    [string]$DeviceId,
    [string]$DeviceName,
    [string]$UserPrincipalName,
    [string]$UserId,
    [string]$BearerToken,
    [string]$TenantUrl,
    $ApiSecrets = $null,
    [bool]$Test = $false
)

## SCRIPT VARIABLES ##

$Device = $null
$TicketNote = "$($DeviceName): "

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

# Get CW Manage credentials

if ($null -eq $ApiSecrets -and $TicketId -ne 0) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## GRAB DEVICE ID IF NOT PROVIDED ##

if ($DeviceId -eq '' -or $null -eq $DeviceId) {
    Write-Warning "Device ID not provided for $DeviceName. Searching now..."

    $FindDeviceArguments = @{
        Uri = 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices'
        Method = 'GET'
        Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
        UseBasicParsing = $true
        Body = @{
            '$filter' = "deviceName eq '$DeviceName'"
        }
    }

    try {
        $DeviceList = Invoke-WebRequest @FindDeviceArguments | ConvertFrom-Json
        $Device = $DeviceList.Value[0]
        Write-Warning "SUCCESS: Device ID located - $($Device.id)"
    } catch {
        Write-Error "Unable to fetch $DeviceName : $($_)"
    }
    
}

## GRAB DEVICE DETAILS IF NEEDED ##

if ($null -eq $Device) {
    Write-Warning "Device ID provided. Fetching device details..."

    $GetDeviceArguments = @{
        Uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$DeviceId"
        Method = 'GET'
        Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
        UseBasicParsing = $true
    }

    try {
        $Device = Invoke-WebRequest @GetDeviceArguments | ConvertFrom-Json
        Write-Warning "SUCCESS: Fetched $($Device.deviceName)"
        $TicketNote = "$($Device.deviceName): "
    } catch {
        Write-Error "Unable to fetch $($DeviceId) : $($_)"
    }
}

## UPDATE PRIMARY USER ##

if ($null -ne $Device.id) {
    $DeviceUserPrincipalName = if ($Device.userPrincipalName) {$Device.userPrincipalName} else {"N/A (no primary user set)"}
    if ($DeviceUserPrincipalName.ToLower() -ne $UserPrincipalName.ToLower()) {
        Write-Warning "$($DeviceUserPrincipalName) does not match $UserPrincipalName. Updating primary user for $($Device.deviceName)..."

        # Grab user details if not provided
        if ($UserId -eq "") {
            $GetUserArguments = @{
                BearerToken = $BearerToken
                TenantUrl = $TenantUrl
                UserPrincipalName = $UserPrincipalName
            }
            $User = .\AAD-FindUser.ps1 @GetUserArguments | ConvertFrom-Json
            $UserId = $User.id
        }

        # Update device's primary user
        if ($null -ne $UserId -and $UserId -ne "") {
            $ApiArguments = @{
                Uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($Device.id)/users/`$ref"
                Method = 'POST'
                Headers = @{
                    'Content-Type' = "application/json"
                    Authorization = "Bearer $BearerToken"
                    ConsistencyLevel = 'eventual'
                }
                UseBasicParsing = $true
                Body = @{
                    '@odata.id' = "https://graph.microsoft.com/beta/users/$($UserId)"
                } | ConvertTo-Json -Depth 100
            }
        
            if (!$Test) {
                try {
                    $Request = Invoke-WebRequest @ApiArguments
                    if ($Request.StatusCode -eq 204) {
                        Write-Warning "SUCCESS: Primary user for $($Device.deviceName) has been set to $($UserPrincipalName)."
                        $TicketNote += "Primary user has been changed from $($DeviceUserPrincipalName) to $($UserPrincipalName)."
                    } else {
                        Write-Error "Status code indicates this was not successful. Request details: $($Request)"
                        $TicketNote += "Status code indicates this was not successful.`n`nRequest details: $($Request)"
                    }
                }
                catch {
                    Write-Error "Unable to update primary user for $($Device.deviceName) : $($_)"
                    $TicketNote += "Unable to update primary user.`n`nError details: $($_)"
                }
            } else {
                Write-Warning "Test flag has been enabled. Primary user would've been changed from $($DeviceUserPrincipalName) to $($UserPrincipalName)."
                $TicketNote += "Test flag has been enabled. Primary user would've been changed from $($DeviceUserPrincipalName) to $($UserPrincipalName)."
            }
        } else {
            # Skip updating if the user was not found
            Write-Warning "No valid User ID provided/located. Primary user for $($Device.deviceName) will not be updated."
            $TicketNote += "No valid User ID provided/located. Primary user will not be updated."
        }
    } else {
        Write-Warning "Device's UPN ($($DeviceUserPrincipalName)) already matches $($UserPrincipalName)."
        $TicketNote += "Device's UPN ($($DeviceUserPrincipalName)) already matches $($UserPrincipalName)."
    }
} else {
    Write-Error "No valid device object found for $($DeviceName)."
    $TicketNote = "No valid device object found for $($DeviceName)."
}

## UPDATE TICKET WITH RESULTS ##

if ($TicketId -ne 0) {
    if ($TicketNote -like "*Primary user has been changed*" -or $TicketNote -like "*Test flag has been enabled*") {
        $NoteArguments = @{
            TicketId = $TicketId
            Text = $TicketNote
            ResolutionFlag = $true
            ApiSecrets = $ApiSecrets
        }
    } else {
        $NoteArguments = @{
            TicketId = $TicketId
            Text = $TicketNote
            InternalFlag = $true
            ApiSecrets = $ApiSecrets
        }

        if ($TicketNote -like "*Error*" -or $TicketNote -like "*No valid device object*") {
            $NoteArguments += @{
                IssueFlag = $true
            }
        }
    }

    .\CWM-AddTicketNote.ps1 @NoteArguments | Out-Null
}

## WRITE OUTPUT ##

if ($TicketNote -like "*Error*") {
    $Output = @{
        Result = $false
    }
} else {
    $Output = @{
        Result = $true
    }
}

Write-Output $Output | ConvertTo-Json