<#

Seasons Living - Setup Teams Calling
Created by: Gabriel Nugent
Version: 1.0.5

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$EmailAddress,
    [int]$SiteId,
    [int]$TicketId,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

# Get CW Manage credentials
if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

# Grab connection variables
$TenantId = "746036af-6519-4420-ba95-65e211e271c6"

# Teams details by site
$SitePolicies = @(
    @{
        Name = 'Bribie Island'
        SiteId = 3659
        PhoneBlockLower = 61734784100
        PhoneNumberStart = 617347841
        CallerId = 'BRBE1-MainLine'
        DialingPlan = 'BRBE1'
    },
    @{
        Name = 'Brendale'
        SiteId = 3059
        PhoneBlockLower = 61734802900
        PhoneNumberStart = 617348029
        CallerId = 'BREND1-MainLine'
        DialingPlan = 'QLD'
    },
    @{
        Name = 'Caloundra'
        SiteId = 3658
        PhoneBlockLower = 61753419100
        PhoneNumberStart = 617534191
        CallerId = 'CDRA1-MainLine'
        DialingPlan = 'CDRA1'
    },
    @{
        Name = 'Eastern Heights'
        SiteId = 3663
        PhoneBlockLower = 61732013100
        PhoneNumberStart = 617320131
        CallerId = ''
        DialingPlan = 'QLD'
    },
    @{
        Name = 'Kallangur'
        SiteId = 3660
        PhoneBlockLower = 61738808500
        PhoneNumberStart = 617388085
        CallerId = 'KLGR1-MainLine'
        DialingPlan = 'QLD'
    },
    @{
        Name = 'Mango Hill'
        SiteId = 3661
        PhoneBlockLower = 617
        PhoneNumberStart = 0
        CallerId = 'MGOH1-MainLine'
        DialingPlan = 'QLD'
    },
    @{
        Name = 'Redbank Plains'
        SiteId = 3664
        PhoneBlockLower = 61734323100
        PhoneNumberStart = 617343231
        CallerId = 'RDBK1-MainLine'
        DialingPlan = 'QLD'
    },
    @{
        Name = 'Sinnamon Park'
        SiteId = 3274
        PhoneBlockLower = 61735655200
        PhoneNumberStart = 617356552
        CallerId = 'SNPK1-MainLine'
        DialingPlan = 'SNPK1'
    },
    @{
        Name = 'Waterford West'
        SiteId = 3662
        PhoneBlockLower = 61734420700
        PhoneNumberStart = 617344207
        CallerId = 'WTFD1-MainLine'
        DialingPlan = 'WTFD1'
    }
)

## RUN SUBSCRIPT ##

$ScriptParams = @{
    EmailAddress = $EmailAddress
    SiteId = $SiteId
    TicketId = $TicketId
    TenantId = $TenantId
    TeamsCallingSites = $SitePolicies
    ApiSecrets = $ApiSecrets
}
$Output = .\Teams-SetupCalling.ps1 @ScriptParams

# Send back to Power Automate
Write-Output $Output | ConvertTo-Json