<#

OPEC Systems - Setup Teams Calling
Created by: Gabriel Nugent
Version: 1.0.1

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
$TenantId = "e1841806-2216-4da6-9fd2-92b75bc65081"

# Teams details by site
$SitePolicies = @(
    @{
        Name = 'BELR1'
        SiteId = 2902
        PhoneBlockLower = 61294542510
        PhoneNumberStart = 612945425
        CallerId = 'Belrose Office'
        DialingPlan = 'NSW'
    },
    @{
        Name = 'HMNT1'
        SiteId = 2771
        PhoneBlockLower = 61730013710
        PhoneNumberStart = 617300137
        CallerId = 'Hemmant Office'
        DialingPlan = 'QLD'
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