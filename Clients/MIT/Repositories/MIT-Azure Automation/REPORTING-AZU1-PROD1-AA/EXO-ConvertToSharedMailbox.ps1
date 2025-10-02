<#

Mangano IT - Exchange Online - Convert to Shared Mailbox
Created by: Gabriel Nugent
Version: 1.1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory=$true)][string]$TenantUrl,
    [Parameter(Mandatory=$true)][string]$TenantSlug,
    [Parameter(Mandatory=$true)][string]$EmailAddress,
    [string]$DelegateEmailAddress,
    [int]$TicketId,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

$TaskNotes = 'Convert To Shared Mailbox'

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

    # Set the mailbox to shared
    try {
        $Log += "Converting mailbox $DistinguishedName to shared...`n"
        Write-Warning "Converting mailbox $DistinguishedName to shared..."
        Set-Mailbox -Identity $DistinguishedName -Type 'Shared'
        $Log += "SUCCESS: Mailbox $DistinguishedName has been converted to shared.`n`n"
        Write-Warning "SUCCESS: Mailbox $DistinguishedName has been converted to shared."
        $MailboxResult = $true
    } catch {
        $Log += "ERROR: Unable to convert mailbox $DistinguishedName to shared.`nERROR DETAILS: " + $_ + "`n`n"
        Write-Error "Unable to convert mailbox $DistinguishedName to shared : $_"
        $MailboxResult = $false
    }

    # Assigns access to delegate
    if ($DelegateEmailAddress -ne '') {
        try {
            $Log += "Granting permissions to $DistinguishedName for $DelegateEmailAddress...`n"
            Write-Warning "Granting permissions to $DistinguishedName for $DelegateEmailAddress..."
            Add-MailboxPermission -Identity $DistinguishedName -AccessRights 'FullAccess' -User $DelegateEmailAddress
            $Log += "SUCCESS: Permissions granted to $DistinguishedName for $DelegateEmailAddress."
            Write-Warning "SUCCESS: Permissions granted to $DistinguishedName for $DelegateEmailAddress."
            $PermissionResult = $true
        } catch {
            $Log += "ERROR: Unable to grant permissions to $DistinguishedName for $DelegateEmailAddress.`nERROR DETAILS: " + $_
            Write-Error "Unable to grant permissions to $DistinguishedName for $DelegateEmailAddress : $_"
            $PermissionResult = $false
        }
    } else { $PermissionResult = $true }
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false | Out-Null

## UPDATE TICKET IF PROVIDED ##

if ($TicketId -ne 0) {
    # Define standard note text
    if ($MailboxResult) {
        $NoteText = "$EmailAddress was converted to a shared mailbox."
        if ($PermissionResult) { $Result = $true }
    } else { $NoteText = "$EmailAddress was not converted to a shared mailbox." }

    if ($DelegateEmailAddress -ne '') {
        if ($PermissionResult) { 
            $NoteText += "`n`nAccess to $EmailAddress was granted to $DelegateEmailAddress."
        } else { $NoteText += "`n`nAccess to $EmailAddress was not granted to $DelegateEmailAddress." }
    } else {
        $NoteText += "`n`nNo address was supplied to provide delegate access to."
    }

    # Define arguments for ticket note and task
    $TaskNoteArguments = @{
        Result = $Result
        TicketId = $TicketId
        TaskNotes = $TaskNotes
        TicketNote_Success = $NoteText
        TicketNote_Failure = $NoteText
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