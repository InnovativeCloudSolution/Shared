<#

Mangano IT - Exchange Online - Get All Distribution Groups
Created by: Gabriel Nugent
Version: 1.0.1

This runbook is designed to be used in conjunction with another PowerShell script.

#>

param(
    [Parameter(Mandatory=$true)][string]$TenantUrl,
    [Parameter(Mandatory=$true)][string]$TenantSlug
)

## SCRIPT VARIABLES ##

# Script status variables
$EXOConnectionStatus = $false
$Mailboxes = @()

## EXCHANGE ##

# Grab connection variables
$ApplicationId = .\KeyVault-GetSecret.ps1 -SecretName "$TenantSlug-EXO-ApplicationID"
$CertificateThumbprint = .\KeyVault-GetSecret.ps1 -SecretName "$TenantSlug-EXO-CertificateThumbprint"

# Connect to Exchange Online
try {
    Connect-ExchangeOnline -CertificateThumbprint $CertificateThumbprint -AppID $ApplicationId -Organization $TenantUrl | Out-Null
    $EXOConnectionStatus = $true
}
catch {
    Write-Error "Unable to connect to Exchange Online : $_"
}

# Fetch all mailboxes

if ($EXOConnectionStatus) {
    try {
        $MailboxRequestParameters = @{ ResultSize = 'unlimited' }

        # Fetch all mailboxes
        $Mailboxes = Get-DistributionGroup @MailboxRequestParameters | Select-Object DisplayName, PrimarySmtpAddress
    }
    catch {
        Write-Error "Unable to fetch distribution lists : $_"
        $Mailboxes = $null
    }
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false | Out-Null

## SEND DETAILS TO SCRIPT ##

Write-Output $Mailboxes