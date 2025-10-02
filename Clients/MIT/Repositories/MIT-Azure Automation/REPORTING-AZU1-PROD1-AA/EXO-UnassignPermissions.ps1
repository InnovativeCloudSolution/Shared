<#

Mangano IT - Exchange Online - Unassign User Permissions in Exchange Online
Created by: Gabriel Nugent
Version: 1.8.2

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory=$true)][string]$TenantUrl,
    [Parameter(Mandatory=$true)][string]$TenantSlug,
    [Parameter(Mandatory=$true)][string]$EmailAddress,
    [int]$TicketId,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

$TaskNotes = 'Remove from Exchange Services'

# Script status variables
$EXOConnectionStatus = $false
$Mailbox = $null
$Result = $true
$SharedMailboxes = @()
$DistributionGroups = @()
$RemovedItems = ''
$FailedItems = ''
$NoteText = ''

## ADD NOTE AT START ##

if ($TicketId -ne 0) {
    $Text = "$TaskNotes [automated task]`n`n"
    $Text += "$EmailAddress will now be exited from Exchange Online.`n`n"
    $Text += "Please note that this particular part of the automation can take longer than the rest (5-10 minutes). "
    $Text += "If the automation breaks, the ticket will be updated accordingly. Otherwise, this step is likely still running."
    .\CWM-AddTicketNote.ps1 -TicketId $TicketId -Text $Text -ResolutionFlag $true -ApiSecrets $ApiSecrets | Out-Null
}

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

## GET ALL GROUPS AND MAILBOXES ##

if ($null -ne $Mailbox) {
    # Get distinguished name
    $DistinguishedName = $Mailbox.DistinguishedName

    # Get all distribution groups
    try {
        $Log += "Fetching all distribution groups that $EmailAddress has access to...`n"
        $Filter = "Members -like ""$DistinguishedName"""
        $DistributionGroupsList = Get-DistributionGroup -ResultSize Unlimited -Filter $Filter
        $Log += "SUCCESS: Fetched all distribution groups that $EmailAddress has access to.`n`n"
        Write-Warning "SUCCESS: Fetched all distribution groups that $EmailAddress has access to."
    }
    catch {
        $Log += "ERROR: Unable to fetch all distribution groups that $EmailAddress has access to.`nERROR DETAILS: " + $_
        Write-Error "Unable to fetch all distribution groups that $EmailAddress has access to : $_"
    }

    # Get all shared mailboxes
    try {
        $Log += "Fetching all shared mailboxes that $EmailAddress has access to...`n"
        $SharedMailboxList = Get-Mailbox -ResultSize unlimited | Get-MailboxPermission | Where-Object {($_.User -like $EmailAddress)}
        $Log += "SUCCESS: Fetched all shared mailboxes that $EmailAddress has access to.`n`n"
        Write-Warning "SUCCESS: Fetched all shared mailboxes that $EmailAddress has access to."
    }
    catch {
        $Log += "ERROR: Unable to fetch all shared mailboxes that $EmailAddress has access to.`nERROR DETAILS: " + $_
        Write-Error "Unable to fetch all shared mailboxes that $EmailAddress has access to : $_"
    }

    # Get all shared mailboxes with send as permissions
    try {
        $Log += "Fetching all shared mailboxes that $EmailAddress has access to with send as permissions...`n"
        $SharedSendAsMailboxList = Get-Mailbox -ResultSize unlimited | Get-RecipientPermission | Where-Object {($_.Trustee -like $EmailAddress)}
        $Log += "SUCCESS: Fetched all shared mailboxes that $EmailAddress has access to with send as permissions.`n`n"
        Write-Warning "SUCCESS: Fetched all shared mailboxes that $EmailAddress has access to with send as permissions."
    }
    catch {
        $Log += "ERROR: Unable to fetch all shared mailboxes that $EmailAddress has access to with send as permissions.`nERROR DETAILS: " + $_
        Write-Error "Unable to fetch all shared mailboxes that $EmailAddress has access to with send as permissions : $_"
    }
}

## REMOVE FROM MAILBOXES AND GROUPS ##

# Remove from all distribution groups
foreach ($DistributionGroup in $DistributionGroupsList) {
    $Identity = $DistributionGroup.PrimarySmtpAddress
    $DistributionGroups += [string]$Identity

    try {
        $Log += "Removing $EmailAddress from distribution group $Identity...`n"
        Remove-DistributionGroupMember -Identity $Identity -Member $DistinguishedName -BypassSecurityGroupManagerCheck -Confirm:$false
        $Log += "SUCCESS: Removed $EmailAddress from distribution group $Identity.`n`n"
        $RemovedItems += "`n- Distribution Group: $Identity"
    } catch {
        $Log += "ERROR: Unable to remove $EmailAddress from distribution group $Identity.`nERROR DETAILS: " + $_ + "`n`n"
        Write-Error "Unable to remove $EmailAddress from distribution group $Identity : $_"
        $FailedItems += "`n- Distribution Group: $Identity"
    }  
}

# Remove from all shared mailboxes
foreach ($SharedMailbox in $SharedMailboxList) {
    $Identity = $SharedMailbox.Identity
    $Address = $SharedMailbox.PrimarySmtpAddress
    $SharedMailboxes += [string]$Address
    $AccessRights = $SharedMailbox.AccessRights

    try {
        $Log += "Removing $EmailAddress from shared mailbox $Identity...`n"
        Remove-MailboxPermission -Identity $Identity -User $EmailAddress -AccessRights $AccessRights -InheritanceType All -Confirm:$false
        $Log += "SUCCESS: Removed $EmailAddress from shared mailbox $Identity.`n`n"
        $RemovedItems += "`n- Shared Mailbox: $Identity"
    } catch {
        $Log += "ERROR: Unable to remove $EmailAddress from shared mailbox $Identity.`nERROR DETAILS: " + $_ + "`n`n"
        Write-Error "Unable to remove $EmailAddress from shared mailbox $Identity : $_"
        $FailedItems += "`n- Shared Mailbox: $Identity"
    }  
}

# Remove from all shared mailboxes with send as permissions
foreach ($SharedMailboxSendAs in $SharedSendAsMailboxList) {
    $Identity = $SharedMailbox.Identity
    $Address = $SharedMailbox.PrimarySmtpAddress
    $SharedMailboxes += "$Address (send as)"
    $AccessRights = $SharedMailboxSendAs.AccessRights
    
    try {
        $Log += "Removing $EmailAddress from shared mailbox (send as permissions) $Identity...`n"
        Remove-RecipientPermission -Identity $Identity -Trustee $EmailAddress -AccessRights $AccessRights -Confirm:$false
        $Log += "SUCCESS: Removed $EmailAddress from shared mailbox (send as permissions) $Identity.`n`n"
        $RemovedItems += "`n- Shared Mailbox (Send As): $Identity"
    } catch {
        $Log += "ERROR: Unable to remove $EmailAddress from shared mailbox (send as permissions) $Identity.`nERROR DETAILS: " + $_ + "`n`n"
        Write-Error "Unable to remove $EmailAddress from shared mailbox (send as permissions) $Identity : $_"
        $FailedItems += "`n- Shared Mailbox (Send As): $Identity"
    }   
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false | Out-Null

## UPDATE TICKET IF PROVIDED ##

if ($TicketId -ne 0) {
    # Define standard note text
    if ($RemovedItems -ne '') {
        $NoteText += "$EmailAddress has been removed from the following items:$RemovedItems`n`n"
        $Result = $true
    }
    
    if ($FailedItems -ne '') {
        $NoteText += "$EmailAddress has not been removed from the following items:$FailedItems`n`n"
        $Result = $false
    }

    if ($RemovedItems -eq '' -and $FailedItems -eq '') {
        $NoteText += "$EmailAddress did not need to be removed from anything in Exchange."
        $Result = $true
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

    # Add missed groups to task if any
    if ($FailedItems -ne '') {
        $TaskNoteArguments += @{
            TaskResolution = "Missed items:$FailedItems"
        }
    }
    
    # Add note and update task if successful
    $TaskAndNote = .\CWM-UpdateTaskAddNoteForFlow.ps1 @TaskNoteArguments
    $Log += $TaskAndNote.Log
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Result = $Result
    DistributionGroups = $DistributionGroups
    SharedMailboxes = $SharedMailboxes
    Log = $Log
}

# Send back to Power Automate
Write-Output $Output | ConvertTo-Json