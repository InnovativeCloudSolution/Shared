<#

Mangano IT - ConnectWise Manage - Get Self-Bundled Tickets
Created by: Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

## SCRIPT VARIABLES ##

# Array of ticket IDs
$TicketIds = @()

## SET UP CONNECTION VAR AND CONNECT ##

# Get credentials
$Connection = .\CWM-CreateConnectionObject.ps1

# Connect to CW Manage
Connect-CWM @Connection

## GET ALL MATCHING TICKETS ##

# Get all child tickets
try {
    Write-Warning "Fetching all child tickets..."
    $Tickets = Get-CWMTicket -condition 'parentTicketId != null' -all
    Write-Warning "Fetched all child tickets. Will begin sorting now."
}
catch {
    Write-Error "Unable to fetch tickets : $($_)"
}

## SORT THROUGH TICKETS ##

# Check for tickets that are bundled into themselves
foreach ($Ticket in $Tickets) {
    if ($Ticket.id -eq $Ticket.parentTicketId) {
        Write-Warning "Found matching ticket: #$($Ticket.id)"
        $TicketIds += $Ticket.id
    }
}

## SEND BACK TO FLOW ##

$Output = @{
    TicketIds = $TicketIds
}

Write-Output $Output | ConvertTo-Json -Depth 100