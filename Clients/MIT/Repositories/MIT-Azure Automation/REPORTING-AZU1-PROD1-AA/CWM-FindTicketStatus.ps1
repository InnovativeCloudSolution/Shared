<#

Mangano IT - ConnectWise Manage - Fetch Ticket Status
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow
for checking the ticket's current status. The output is designed to be as simple as
possible for ease of use.

#>

param (
    [Parameter(Mandatory=$true)][int]$TicketId,
    $ApiSecrets = $null
)

## GET TICKET ##

$Ticket = .\CWM-FindTicketDetails.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets | ConvertFrom-Json

## SEND DETAILS TO FLOW ##

Write-Output $Ticket.status.name