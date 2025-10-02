<#

Seasons Living - Update Ticket for F3 User
Created by: Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow or an Azure Automation script.

#>

param(
    [int]$TicketId,
    [string]$GivenName,
    $ApiSecrets = $null
)

$TSBoardName = "HelpDesk (TS)"

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## GET TICKET ##

$Ticket = .\CWM-FindTicketDetails.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets | ConvertFrom-Json

## SET VARIABLES FOR TICKET BOARD UPDATE ##

$BoardParameters = @{
    TicketId = $TicketId
    BoardName = $TSBoardName
    ApiSecrets = $ApiSecrets
}

# Set parameters if ticket's on the triage board
if ($Ticket.board.name -eq 'Triage') {
    $BoardParameters += @{
        StatusName = 'Pre-Process'
        TypeName = 'Request'
    }
} else {
    if ($null -ne $Ticket.status) { $BoardParameters += @{ StatusName = $Ticket.status.name } }
    if ($null -ne $Ticket.type) { $BoardParameters += @{ TypeName = $Ticket.type.name } }
    if ($null -ne $Ticket.subType) { $BoardParameters += @{ SubtypeName = $Ticket.subType.name } }
    if ($null -ne $Ticket.item) { $BoardParameters += @{ ItemName = $Ticket.item.name } }
}

# Update board
.\CWM-UpdateTicketBoard.ps1 @BoardParameters | Out-Null

## REMOVE AGREEMENT ##

.\CWM-RemoveTicketAgreement.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets | Out-Null

## ADD NOTE ##

# Setup standard note
$NoteText = "The user $GivenName is not a supported staff member at Seasons. "
$NoteText += "As such, the agreement will be removed from this ticket, and the ticket will be moved to $TSBoardName.`n`n"
$NoteText += "F3/RSW user support info: https://mits.itglue.com/2642483/docs/11610348"

$NoteParameters = @{
    TicketId = $TicketId
    ApiSecrets = $ApiSecrets
    InternalFlag = $true
    Text = $NoteText
}

.\CWM-AddTicketNote.ps1 @NoteParameters | Out-Null