<#

Mangano IT - ConnectWise Manage - Fetch Tickets (Ticket Number Only)
Created by: Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param (
    [Parameter(Mandatory)][string]$Summary,
    [Parameter(Mandatory)][ValidateSet('Equals','StartsWith','EndsWith')][string]$SummaryComparison,
    [string]$CompanyName,
    [string]$LastUpdatedDateTime,
    [bool]$IncludeClosed = $false
)

## SCRIPT VARIABLES ##

$Output = @()

## SET UP CONNECTION VAR AND CONNECT ##

# Get credentials
$Connection = .\CWM-CreateConnectionObject.ps1

# Connect to our CWM server
Connect-CWM @Connection

## SETUP REQUEST VARIABLES ##

# Condense summary length if required
if ($Summary.Length -gt 100) {
    $Summary = $Summary.Substring(0, 100)
}

# Format summary as needed
switch ($SummaryComparison) {
    'StartsWith' { $SummarySearch = "$Summary*" }
    'EndsWith' { $SummarySearch = "*$Summary" }
    Default { $SummarySearch = "$Summary" }
}

$Conditions = "summary like '$($SummarySearch)'"

# Add condition if closed tickets are to be skipped
if (!$IncludeClosed) {
    $Conditions += " AND closedFlag = false"
}

# Add extra conditions if provided
if ($CompanyName -ne '') {
    $Conditions += " AND company/name = '$CompanyName'"
}

if ($LastUpdatedDateTime -ne '') {
    $Conditions += " AND LastUpdated >= [$LastUpdatedDateTime]"
}

## FETCH DETAILS FROM TICKET ##

try { 
    $Tickets = Get-CWMTicket -condition $Conditions
    Write-Warning "SUCCESS: Tickets fetched."
} catch { 
    Write-Error "Tickets unable to be fetched : $($_)"
    $Tickets = $null
}

## SORT THROUGH TICKETS ##

# Add single 1-value to ticket if there's none located
if ($null -eq $Tickets) {
    Write-Warning "No tickets were located. Array will still be exported for Power Automate purposes."
    $Output = @(1)
} else {
    # Grab ticket ID from each ticket
    foreach ($Ticket in $Tickets) {
        if ($null -ne $Ticket.id) {
            $Output += $Ticket.id
        }
    }
}

## SEND DETAILS TO FLOW ##

# Convert output to JSON, allowing for single-value arrays
$JsonOutput = ConvertTo-Json -InputObject $Output -Depth 100

Write-Output $JsonOutput