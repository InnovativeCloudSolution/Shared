<#

Mangano IT - Exchange Online - Add to Safe Senders List
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory=$true)][string]$TenantUrl,
    [Parameter(Mandatory=$true)][string]$TenantSlug,
    [Parameter(Mandatory=$true)][string]$EmailAddress,
    [Parameter(Mandatory)][array]$Senders
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

# Script status variables
$EXOConnectionStatus = $false
$Mailbox = $null
$Result = $false

## EXCHANGE ##

# Grab connection variables
$ApplicationId = .\KeyVault-GetSecret.ps1 -SecretName "$TenantSlug-EXO-ApplicationID"
$CertificateThumbprint = .\KeyVault-GetSecret.ps1 -SecretName "$TenantSlug-EXO-CertificateThumbprint"

# Connect to Exchange Online
try {
    $Log += "Connecting to Exchange Online...`n"
    Connect-ExchangeOnline -CertificateThumbprint $CertificateThumbprint -AppID $ApplicationId -Organization $TenantUrl | Out-Null
    $EXOConnectionStatus = $true
    $Log += "SUCCESS: Connected to Exchange Online.`n`n"
}
catch {
    $Log += "ERROR: Unable to connect to Exchange Online.`nERROR DETAILS: " + $_
    Write-Error "Unable to connect to Exchange Online : $_"
}

# Fetch mailbox
if ($EXOConnectionStatus) {
    try {
        $Log += "Fetching mailbox for $EmailAddress...`n"
        $Mailbox = Get-Mailbox $EmailAddress
        $Log += "SUCCESS: Fetched mailbox for $EmailAddress.`n`n"
    }
    catch {
        $Log += "ERROR: Unable to fetch user's mailbox.`nERROR DETAILS: " + $_
        Write-Error "Unable to fetch user's mailbox : $_"
        $Mailbox = $null
    }
}

# Update safe senders list
if ($null -ne $Mailbox) {
    try {
        $Log += "Adding senders to $EmailAddress's safe senders list...`n"
        Set-MailboxJunkEmailConfiguration $EmailAddress -TrustedSendersAndDomains @{Add=$Senders} | Out-Null
        $Log += "SUCCESS: Added senders to $EmailAddress's safe senders list."
        Write-Warning "SUCCESS: Added senders to $EmailAddress's safe senders list."
        $Result = $true
    }
    catch {
        $Log += "ERROR: Unable to add senders to $EmailAddress's safe senders list.`nERROR DETAILS: " + $_
        Write-Error "Unable to add senders to $EmailAddress's safe senders list : $_"
        $Result = $false
    }
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false | Out-Null

## SEND DETAILS TO FLOW ##

$Output = @{
    Result = $Result
    Log = $Log
}

# Send back to Power Automate
Write-Output $Output | ConvertTo-Json