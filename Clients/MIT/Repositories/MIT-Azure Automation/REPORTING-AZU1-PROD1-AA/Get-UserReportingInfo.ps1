<#

Mangano IT - Get User Info for Monthly Reporting
Created by: Gabriel Nugent
Version: 1.3

This runbook is designed to be run in conjunction with a Power Automate flow.

#>

param(
	[string]$BearerToken,
    [string]$TenantUrl
)

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    $Log += "Bearer token not supplied. Getting bearer token for $TenantUrl...`n`n"
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

# Form request headers with the acquired $AccessToken
$WebRequestHeaders = @{'Content-Type'="application\json";'Authorization'="Bearer $BearerToken"}
 
# This request get users list with signInActivity.
$ApiUrl = "https://graph.microsoft.com/beta/users?`$select=id,displayName,userPrincipalName,jobTitle,companyName,signInActivity,userType,assignedLicenses,streetAddress,city,state,postalCode,accountEnabled,department"
$Result = @()
while ($Null -ne $ApiUrl) { # Perform pagination if next page link (odata.nextlink) returned.
    try {
        $Response = Invoke-WebRequest -UseBasicParsing -Method GET -Uri $ApiUrl -ContentType "application\json" -Headers $WebRequestHeaders | ConvertFrom-Json
    } catch {
        # Run query again if it fails - most likely a tenancy without a Premium AAD license
        Write-Warning "Running query again - tenancy may not have required AAD Premium license."
        $ApiUrl = "https://graph.microsoft.com/beta/users?`$select=id,displayName,userPrincipalName,jobTitle,companyName,userType,assignedLicenses,streetAddress,city,state,postalCode,accountEnabled,department"
        $Response = Invoke-WebRequest -UseBasicParsing -Method GET -Uri $ApiUrl -ContentType "application\json" -Headers $WebRequestHeaders | ConvertFrom-Json
    }

    if ($Response.value) {
        $Users = $Response.value
        foreach ($User in $Users) {
            # Get ID to find manager
            $UserId = $User.id
            $ApiUrlManager = "https://graph.microsoft.com/v1.0/users/$UserId/manager"
			$Manager = $null
			try {
            	$Manager = Invoke-WebRequest -UseBasicParsing -Method GET -Uri $ApiUrlManager -ContentType "application\json" -Headers $WebRequestHeaders -ErrorAction 'SilentlyContinue' | ConvertFrom-Json
			} catch {}

            # Build address if all pieces exist
            $UserAddress = $null
            if ($User.streetAddress -and $User.city -and $User.state -and $User.postalCode) {
                $StreetAddress = $User.streetAddress
                $City = $User.city
                $State = $User.state
                $Postcode = $User.postalCode
                $UserAddress = "$StreetAddress, $City, $State $Postcode"
            }

            # Get license info
            $UserLicenses = @()
            if ($User.assignedLicenses.Count -ne 0) {
                $ApiUrlLicenses = "https://graph.microsoft.com/v1.0/users/$UserId/licenseDetails"
                try {
                    $Licenses = Invoke-WebRequest -UseBasicParsing -Method GET -Uri $ApiUrlLicenses -ContentType "application\json" -Headers $WebRequestHeaders -ErrorAction SilentlyContinue | ConvertFrom-Json
                    $LicensesValue = $Licenses.value
                } catch {}
                
                foreach ($License in $LicensesValue) {
                    $SkuPartNumber = $License.skuPartNumber
                    $LicenseName = $null
                    switch ($SkuPartNumber) {
                        "FLOW_FREE" { $LicenseName = "Microsoft Power Automate Free" }
                        "EMSPREMIUM" { $LicenseName = "Enterprise Mobility + Security E5" }
                        "SPE_E3" { $LicenseName = "Microsoft 365 E3" }
                        "SPE_E5" { $LicenseName = "Microsoft 365 E5" }
                        "STANDARDPACK" { $LicenseName = "Office 365 E1" }
                        "DESKLESSPACK" { $LicenseName = "Office 365 F3" }
                        "POWER_BI_STANDARD" { $LicenseName = "Power BI (Free)" }
                        "POWER_BI_PRO" { $LicenseName = "Power BI Pro" }
                        "PBI_PREMIUM_PER_USER" { $LicenseName = "Power BI Premium Per User" }
                        "VISIOCLIENT" { $LicenseName = "Visio Plan 2" }
                        "STREAM" { $LicenseName = "Microsoft Stream Trial" }
                        "PROJECTESSENTIALS" { $LicenseName = "Project Online Essentials" }
                        "PROJECTPROFESSIONAL" { $LicenseName = "Project Plan 3" }
                        "PROJECTPREMIUM" { $LicenseName = "Project Plan 5" }
                        "POWERAPPS_PER_APP_IW" { $LicenseName = "PowerApps per app baseline access" }
                        "OFFICESUBSCRIPTION" { $LicenseName = "Microsoft 365 Apps for enterprise" }
                        "POWERAPPS_VIRAL" { $LicenseName = "Microsoft Power Apps Plan 2 Trial" }
                        "EXCHANGESTANDARD" { $LicenseName = "Exchange Online (Plan 1)" }
                        "FORMS_PRO" { $LicenseName = "Dynamics 365 Customer Voice Trial" }
                        "CCIBOTS_PRIVPREV_VIRAL" { $LicenseName = "Power Virtual Agents Viral Trial" }
                        "MEETING_ROOM" { $LicenseName = "Microsoft Teams Rooms Standard" }
                        "INTUNE_A" { $LicenseName = "Intune" }
                        "O365_BUSINESS_ESSENTIALS" { $LicenseName = "Microsoft 365 Business Basic" }
                        "SPB" { $LicenseName = "Microsoft 365 Business Premium" }
                        "MCOPSTNEAU2" { $LicenseName = "Telstra Calling for O365" }
                        "PHONESYSTEM_VIRTUALUSER" { $LicenseName = "Microsoft Teams Phone Resoure Account" }
                        "MCOEV" { $LicenseName = "Microsoft Teams Phone Standard" }
                        "MCOCAP" { $LicenseName = "Microsoft Teams Shared Devices" }
                        "EXCHANGEENTERPRISE" { $LicenseName = "Exchange Online (Plan 2)" }
                        "INTUNE_A_D" { $LicenseName = "Microsoft Intune Plan 1 Device" }
                        "WINDOWS_STORE" { $LicenseName = "Windows Store for Business" }
                        "FLOW_PER_USER" { $LicenseName = "Power Automate per user plan" }
                        "SHAREPOINTSTORAGE" { $LicenseName = "Office 365 Extra File Storage" }
                        "DYN365_ENTERPRISE_P1_IW" { $LicenseName = "Dynamics 365 P1 Trial for Information Workers" }
                        "AAD_PREMIUM" { $LicenseName = "Azure Active Directory Premium P1" }
                        "SPE_F1" { $LicenseName = "Microsoft 365 F3" }
                        "DYN365_BUSCENTRAL_ADD_ENV_ADDON" { $LicenseName = "Dynamics 365 Business Central Additional Environment Addon" }
                        "DYN365_BUSCENTRAL_PREMIUM" { $LicenseName = "Dynamics 365 Business Central Premium" }
                        "DYN365_BUSCENTRAL_TEAM_MEMBER" { $LicenseName = "Dynamics 365 Business Central Team Members" }
                        "Power_Pages_vTrial_for_Makers" { $LicenseName = "Power Pages vTrial for Makers" }
                        "PROJECT_MADEIRA_PREVIEW_IW_SKU" { $LicenseName = "Dynamics 365 Business Central for IWs" }
                        "SKU_Dynamics_365_for_HCM_Trial" { $LicenseName = "Dynamics 365 for Talent" }
                        "TEAMS_EXPLORATORY" { $LicenseName = "Microsoft Teams Exploratory" }
                        "AX7_USER_TRIAL" { $LicenseName = "Microsoft Dynamics AX7 User Trial" }
                        "CDS_DB_CAPACITY" { $LicenseName = "Common Data Service Database Capacity" }
                        "MCOMEETADV" { $LicenseName = "Microsoft 365 Audio Conferencing" }
                        "MTR_PREM" { $LicenseName = "Teams Room Premium" }
                        "Microsoft_Teams_Rooms_Pro" { $LicenseName = "Microsoft Teams Rooms Pro" }
                        "POWERAPPS_PER_APP" { $LicenseName = "Power Apps per app Plan" }
                        "RMSBASIC" { $LicenseName = "Rights Management Service Basic Content Protection" }
                        Default { $LicenseName = $SkuPartNumber }
                    }
                    $UserLicenses += $LicenseName
                }
            }
			else { $UserLicenses = @("") }

            $Result += New-Object PSObject -property $([ordered]@{ 
                DisplayName = $User.displayName
                UserPrincipalName = $User.userPrincipalName
                JobTitle = if ($User.jobTitle) {$User.jobTitle} else {""}
                Company = if ($User.companyName) {$User.companyName} else {""}
				Department = if ($User.department) {$User.department} else {""}
                Address = if ($UserAddress) {$UserAddress} else {""}
                LastSignInDateTime = if ($User.signInActivity.lastSignInDateTime) { ([DateTime]$User.signInActivity.lastSignInDateTime).ToString("dddd dd/MM/yyyy HH:mm K") } else {""}
                LastSignInDateTimeExcel = if ($User.signInActivity.lastSignInDateTime) { ([DateTime]$User.signInActivity.lastSignInDateTime).ToString("dd/MM/yyyy HH:mm:ss") } else {""}
                Manager = if ($Manager.displayName) {$Manager.displayName} else {""}
                IsLicensed = if ($User.assignedLicenses.Count -ne 0) { $true } else { $false }
                Licenses = $UserLicenses
                IsGuestUser = if ($User.userType -eq 'Guest') { $true } else { $false }
                IsEnabled = $User.accountEnabled
            })
        }
    }
    $ApiUrl = $Response.'@odata.nextlink'
}

Write-Output $Result | ConvertTo-Json