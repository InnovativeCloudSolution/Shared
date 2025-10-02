<#

Mangano IT - Partner Center - Consent to Enterprise Application
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be run by a Power Automate flow.

#>

param (
    [Parameter(Mandatory)][string]$ApplicationId,
    [Parameter(Mandatory)][string]$ApplicationDisplayName,
    [Parameter(Mandatory)][string]$AccessToken,
    [Parameter(Mandatory)][string]$CustomerTenantId,
    [Parameter(Mandatory)][string]$RequiredGrantsJson
)

## CONNECT TO PARTNER CENTER ##

try {
    Connect-PartnerCenter -AccessToken $AccessToken | Out-Null
    Write-Warning "Connected to Partner Center."
} catch {
    Write-Error "Unable to connect to Partner Center : $($_)"
}

## CREATE CONSENT DETAILS ##

# Create array to house grants
$Grants = @()

# Convert grants from JSON object
$RequiredGrants = $RequiredGrantsJson | ConvertFrom-Json
foreach ($RequiredGrant in $RequiredGrants) {
    $Grant = New-Object -TypeName Microsoft.Store.PartnerCenter.Models.ApplicationConsents.ApplicationGrant
    $Grant.EnterpriseApplicationId = $RequiredGrant.EnterpriseApplicationId
    $Grant.Scope = $RequiredGrant.Scope
    $Grants += $Grant
}

## CONSENT TO APPLICATION ##

try {
    New-PartnerCustomerApplicationConsent -ApplicationGrants $Grants -CustomerId $CustomerTenantId -ApplicationId $ApplicationId -DisplayName $ApplicationDisplayName
    Write-Warning "Consent granted."
} catch {
    Write-Error "Consent not granted : $($_)"
}

## DISCONNECT FROM PARTNER CENTER ##

try {
    Disconnect-PartnerCenter | Out-Null
    Write-Warning "Disconnected from Partner Center."
}
catch {
    Write-Error "Unable to disconnect from Partner Center : $($_)"
}