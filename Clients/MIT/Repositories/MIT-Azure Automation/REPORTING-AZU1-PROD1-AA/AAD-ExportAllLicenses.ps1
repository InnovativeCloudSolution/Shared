<#

Mangano IT - Azure Active Directory - Export All Licenses
Created by: Gabriel Nugent
Version: 1.1.2

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$BearerToken,
	[string]$TenantUrl,
    [bool]$FilterOutZeroCountLicenses = $false
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

$Licenses = @()

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    $Log += "Bearer token not supplied. Getting bearer token for $TenantUrl...`n`n"
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

$ApiArguments = @{
    Uri = 'https://graph.microsoft.com/v1.0/subscribedSkus/?$select=id,skuId,skuPartNumber,consumedUnits,prepaidUnits'
    Method = 'GET'
    Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
	UseBasicParsing = $true
}

## PULL ALL LICENSES ##

try {
    $Log += "Attempting to pull all licenses for tenant...`n"
    $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
    $Log += "SUCCESS: License details have been pulled.`n`n"
	Write-Warning "SUCCESS: License details have been pulled."
} catch {
    $Log += "ERROR: License details not pulled.`nERROR DETAILS: " + $_
	Write-Error "License details not pulled : $_"
    $ApiResponse = $null
}

## ORGANISE LICENSES ##

foreach ($License in $ApiResponse.value) {
    # Set license name
    switch ($License.skuPartNumber) {
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
        Default { $LicenseName = $License.skuPartNumber }
    }

    # Get count of available licenses
    $AvailableUnits = ($License.prepaidUnits.enabled - $License.consumedUnits - $License.prepaidUnits.suspended - $License.prepaidUnits.warning)

    # Add license if there is at least one (or if filter is disabled)
    if (($FilterOutZeroCountLicenses -and $License.prepaidUnits.enabled -gt 0) -or !$FilterOutZeroCountLicenses) {
        $Licenses += @{
            Name = $LicenseName
            SkuPartNumber = $License.skuPartNumber
            Id = $License.id
            SkuId = $License.skuId
            ConsumedUnits = $License.consumedUnits
            EnabledUnits = $License.prepaidUnits.enabled
            SuspendedUnits = $License.prepaidUnits.suspended
            WarningUnits = $License.prepaidUnits.warning
            AvailableUnits = $AvailableUnits
        }
    }
}

## SEND DETAILS BACK TO FLOW ##

# Convert licenses to JSON (preserves array of one if it exists)

$LicensesJson = ConvertTo-Json -InputObject $Licenses -Depth 100

Write-Output $LicensesJson