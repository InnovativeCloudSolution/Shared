<#

Mangano IT - Exchange Online - Get All Mailboxes
Created by: Gabriel Nugent
Version: 1.0.1

This runbook is designed to be used in conjunction with another PowerShell script.

#>

param(
    [Parameter(Mandatory=$true)][string]$TenantUrl,
    [Parameter(Mandatory=$true)][string]$TenantSlug,
    [string]$MailboxType,
    [bool]$InactiveMailboxes = $false
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

        # Add filter if supplied
        if ($MailboxType -ne '') { $MailboxRequestParameters += @{ RecipientTypeDetails = $MailboxType } }

        # Search for inactive if requested
        if ($InactiveMailboxes) { $MailboxRequestParameters += @{ IncludeInactiveMailboxes = $true } }

        # Fetch all mailboxes
        $Mailboxes = Get-Mailbox @MailboxRequestParameters | Select-Object DisplayName, PrimarySmtpAddress
    }
    catch {
        Write-Error "Unable to fetch mailboxes : $_"
        $Mailboxes = $null
    }
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false | Out-Null

## SEND DETAILS TO SCRIPT ##

Write-Output $Mailboxes