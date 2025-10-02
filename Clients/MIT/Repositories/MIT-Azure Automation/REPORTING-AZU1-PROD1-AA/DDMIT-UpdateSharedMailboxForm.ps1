<#

Mangano IT - DeskDirector - Update Shared Mailbox Form
Created by: Gabriel Nugent
Version: 1.2

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory)][string]$FormId,
    [Parameter(Mandatory=$true)][string]$TenantUrl,
    [Parameter(Mandatory=$true)][string]$TenantSlug,
    [string]$FormQuestionName = 'Shared mailboxes',
    [string]$BlacklistedAddresses
)

## SCRIPT VARIABLES ##

$Result = $false
$DesiredChoices = @()
$DisabledUsers = @()
$EntityId = [int]$FormId

## GET ALL MAILBOXES ##

# Fetch mailboxes
$Mailboxes = .\EXO-GetAllMailboxes.ps1 -TenantUrl $TenantUrl -TenantSlug $TenantSlug -MailboxType 'SharedMailbox'

# Fetch list of users
$Users = .\AAD-GetListOfUsers.ps1 -TenantUrl $TenantUrl -Mail

# Make list of disabled account UPNs
foreach ($User in $Users) {
    if (!$User.accountEnabled) {
        $DisabledUsers += $User.userPrincipalName
        $DisabledUsers += $User.mail
    }
}

# Sort through and make array
foreach ($Mailbox in $Mailboxes) {
    if (!$DisabledUsers.Contains($Mailbox.PrimarySmtpAddress) -and !$BlacklistedAddresses.Split(',').Contains($Mailbox.PrimarySmtpAddress)) {
        $MailboxName = $Mailbox.DisplayName
        $MailboxAddress = $Mailbox.PrimarySmtpAddress
        $DesiredChoices += "$MailboxName ($MailboxAddress)"
    }
}

## UPDATE FORM ##

$Request = .\DeskDirector-UpdateFormQuestionOptions.ps1 -EntityId $EntityId -FormQuestionName $FormQuestionName -DesiredChoices $DesiredChoices -DropDown | ConvertFrom-Json
$Result = $Request.Result

## SEND OUTPUT TO FLOW ##

$Output = @{
    Result = $Result
}

Write-Output $Output | ConvertTo-Json