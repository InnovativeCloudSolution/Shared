<#

Mangano IT - Exchange Online - Set Default Font
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory=$true)][string]$TenantUrl,
    [Parameter(Mandatory=$true)][string]$TenantSlug,
    [Parameter(Mandatory=$true)][string]$EmailAddress,
    [Parameter(Mandatory=$true)][string]$FontName,
    [int]$FontSize
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
    $Log += "ERROR: Unable to connect to Exchange Online.`nERROR DETAILS: $($_)"
    Write-Error "Unable to connect to Exchange Online : $($_)"
}

# Fetch mailbox

if ($EXOConnectionStatus) {
    try {
        $Log += "Fetching mailbox for $EmailAddress...`n"
        $Mailbox = Get-Mailbox $EmailAddress
        $Log += "SUCCESS: Fetched mailbox for $EmailAddress.`n`n"
    }
    catch {
        $Log += "ERROR: Unable to fetch user's mailbox.`nERROR DETAILS: $($_)"
        Write-Error "Unable to fetch user's mailbox : $($_)"
        $Mailbox = $null
    }
}

# Set default font and size
if ($null -ne $Mailbox) {
    try {
        $Log += "Setting default font for $($EmailAddress)...`n"
        Set-MailboxMessageConfiguration $Mailbox.Id -DefaultFontName $FontName -DefaultFontSize $FontSize | Out-Null
        $Log += "SUCCESS: Default font has been set for $($EmailAddress)."
        Write-Warning "SUCCESS: Default font has been set for $($EmailAddress)."
        $Result = $true
    }
    catch {
        $Log += "ERROR: Unable to set default font for $($EmailAddress).`nERROR DETAILS: $($_)"
        Write-Error "Unable to set default font for $($EmailAddress) : $($_)"
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