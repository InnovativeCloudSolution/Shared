<#

Mangano IT - Exchange Online - Set Mailbox to Only Accept Messages From Given List
Created by: Gabriel Nugent
Version: 1.1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory=$true)][string]$TenantUrl,
    [Parameter(Mandatory=$true)][string]$TenantSlug,
    [Parameter(Mandatory=$true)][string]$EmailAddress,
    [Parameter(Mandatory=$true)][string]$AddressList,
    [int]$TicketId,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

# Script status variables
$EXOConnectionStatus = $false
$Mailbox = $null
$Result = $false

# Split array
$AddressListArray = $AddressList.Split(";")

## GET CREDENTIALS IF NOT PROVIDED AND REQUIRED ##

if ($null -eq $ApiSecrets -and $TicketId -ne 0) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## EXCHANGE ##

# Grab connection variables
$ApplicationId = .\KeyVault-GetSecret.ps1 -SecretName "$TenantSlug-EXO-ApplicationID"
$CertificateThumbprint = .\KeyVault-GetSecret.ps1 -SecretName "$TenantSlug-EXO-CertificateThumbprint"

# Connect to Exchange Online
try {
    Connect-ExchangeOnline -CertificateThumbprint $CertificateThumbprint -AppID $ApplicationId -Organization $TenantUrl | Out-Null
    $EXOConnectionStatus = $true
    Write-Warning "SUCCESS: Connected to Exchange Online."
}
catch {
    Write-Error "Unable to connect to Exchange Online : $($_)"
}

# Fetch mailbox

if ($EXOConnectionStatus) {
    try {
        $Mailbox = Get-Mailbox $EmailAddress
        Write-Warning "SUCCESS: Fetched mailbox for $EmailAddress."
    }
    catch {
        Write-Error "Unable to fetch user's mailbox : $($_)"
        $Mailbox = $null
    }
}

# Add email address list to accept messges from
if ($null -ne $Mailbox) {
    try {
        Set-Mailbox -Identity $Mailbox.Id -AcceptMessagesOnlyFrom @{add=$AddressListArray}
        Write-Warning "SUCCESS: $($EmailAddress) has been set to only accept messages from provided list ($($AddressList))."
        $Result = $true
    }
    catch {
        Write-Error "Unable to set accepted senders for $($EmailAddress) : $($_)"
        $Result = $false
    }
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false | Out-Null

## SEND DETAILS TO TICKET ##

if ($TicketId -ne 0) {
    # Set note details depending on result
    if ($Result) {
        $Text = "$($EmailAddress) has been set to only accept messages from the provided list ($($AddressList))."

        $NoteArguments = @{
            TicketId = $TicketId
            Text = $Text
            ResolutionFlag = $true
            ApiSecrets = $ApiSecrets
        }
    } else {
        $Text = "An error occurred while trying to set $($EmailAddress) to only accept messages from the provided list ($($AddressList))."

        $NoteArguments = @{
            TicketId = $TicketId
            Text = $Text
            InternalFlag = $true
            ApiSecrets = $ApiSecrets
        }
    }

    .\CWM-AddTicketNote.ps1 @NoteArguments | Out-Null
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Result = $Result
}

# Send back to Power Automate
Write-Output $Output | ConvertTo-Json