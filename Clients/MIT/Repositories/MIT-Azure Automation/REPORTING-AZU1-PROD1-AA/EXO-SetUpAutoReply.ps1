<#

Mangano IT - Exchange Online - Set up Auto Reply Email
Created by: Gabriel Nugent
Version: 1.2

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory=$true)][string]$TenantUrl,
    [Parameter(Mandatory=$true)][string]$TenantSlug,
    [Parameter(Mandatory=$true)][string]$EmailAddress,
    [Parameter(Mandatory=$true)][string]$InternalMessage,
    [Parameter(Mandatory=$true)][string]$ExternalMessage,
    [string]$EndTime,
    [int]$TicketId,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

$TaskNotes = 'Set Up Auto-Response Email'

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

## FETCH MAILBOX ##

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

## SET MAILBOX ##

if ($null -ne $Mailbox) {
    # Get distinguished name
    $DistinguishedName = $Mailbox.DistinguishedName

    # Add auto-response
    try {
        $Log += "Adding auto-response to $DistinguishedName's mailbox...`n"
        Write-Warning "Adding auto-response to $DistinguishedName's mailbox..."

        $Parameters = @{
            Identity = $DistinguishedName
            AutoReplyState = 'Enabled'
            InternalMessage = $InternalMessage
            ExternalMessage = $ExternalMessage
        }

        if ($EndTime -ne '') {
            $Parameters += @{
                StartTime = Get-Date
                EndTime = $EndTime
            }
            $Parameters.AutoReplyState = 'Scheduled'
        }

        Set-MailboxAutoReplyConfiguration @Parameters
        Write-Warning "SUCCESS: Added auto-response to $DistinguishedName's mailbox."
        $Result = $true
    } catch {
        $Log += "ERROR: Unable to add auto-response to $DistinguishedName's mailbox..`nERROR DETAILS: " + $_ + "`n`n"
        Write-Error "Unable to add auto-response to $DistinguishedName's mailbox : $_"
        $Result = $false
    }
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false | Out-Null

## UPDATE TICKET IF PROVIDED ##

if ($TicketId -ne 0) {
    # Define arguments for ticket note and task
    $TaskNoteArguments = @{
        Result = $Result
        TicketId = $TicketId
        TaskNotes = $TaskNotes
        TicketNote_Success = "An auto-response was set up for $EmailAddress."
        TicketNote_Failure = "An auto-response was not set up for $EmailAddress."
        ApiSecrets = $ApiSecrets
    }
    
    # Add note and update task if successful
    $TaskAndNote = .\CWM-UpdateTaskAddNoteForFlow.ps1 @TaskNoteArguments
    $Log += $TaskAndNote.Log
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Result = $Result
    Log = $Log
}

# Send back to Power Automate
Write-Output $Output | ConvertTo-Json