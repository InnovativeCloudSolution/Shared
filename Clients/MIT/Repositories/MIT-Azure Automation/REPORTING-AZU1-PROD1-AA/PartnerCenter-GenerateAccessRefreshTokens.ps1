<#

Mangano IT - Partner Center - Generate Access and Refresh Tokens
Created by: Gabriel Nugent
Version: 1.0

This runbook has to be run manually.

Based on https://tminus365.com/my-automations-break-with-gdap-the-fix/

#>

param (
    [Parameter(Mandatory)][string]$ApplicationId,
    [Parameter(Mandatory)][string]$ApplicationSecret,
    [Parameter(Mandatory)][string]$TenantId
)

$Scopes = 'https://api.partnercenter.microsoft.com/user_impersonation'
$AppCredential = (New-Object System.Management.Automation.PSCredential ($ApplicationId, (ConvertTo-SecureString $ApplicationSecret -AsPlainText -Force)))

# Get PartnerAccessToken token
$PartnerAccessToken = New-PartnerAccessToken -serviceprincipal -ApplicationId $ApplicationId -Credential $AppCredential -Scopes $Scopes -tenant $Tenantid -UseAuthorizationCode

Write-Output $PartnerAccessToken