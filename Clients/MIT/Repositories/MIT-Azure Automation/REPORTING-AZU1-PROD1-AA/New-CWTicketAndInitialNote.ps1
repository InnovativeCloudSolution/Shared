<#

Mangano IT - ConnectWise Manage - Create Ticket And Add Initial Note
Created by: Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
	[Parameter(Mandatory=$true)][string]$Summary,
	[Parameter(Mandatory=$true)][int]$CompanyId,
	[string]$BoardName = 'HelpDesk (MS)',
	[Parameter(Mandatory=$true)][string]$InitialDescriptionNote,
	[string]$InitialInternalNote,
	[string]$InitialResolutionNote,
	[int]$ContactId,
	[string]$ContactName,
	[int]$SiteId,
	[string]$StatusName = 'Pre-Process',
	[string]$TypeName = 'Request',
	[string]$SubtypeName,
	[string]$ItemName,
	[string]$PriorityName = 'P4 Normal Response',
	[string]$Level
)

## SET UP CONNECTION VAR AND CONNECT ##

# Get credentials
$Connection = .\CWM-CreateConnectionObject.ps1

# Connect to our CWM server
Connect-CWM @Connection

## MANAGE TICKET ##

# Condense summary if too long
if ($Summary.Length -gt 99) {
    $Summary = $Summary.Substring(0, 99)
}

# Create ticket info
# Some are parsed as hashtables based on whether or not they're sub-items in the API
$NewTicketParameters = @{
    summary = $Summary
    company = @{id = $CompanyId}
	board = @{name = $BoardName}
	initialDescription = $InitialDescriptionNote
	priority = @{name = $PriorityName}
}

# Add contact if one provided
if ($ContactId -ne 0) { $NewTicketParameters.Add('contact', @{id = $ContactId}) }
else { if ($ContactName -ne '') { $NewTicketParameters.Add('contactName', $ContactName) } }

# Add site if one provided
if ($SiteId -ne 0) { $NewTicketParameters.Add('site', @{id = $SiteId}) }

# Add status if one provided
if ($StatusName -ne '') { $NewTicketParameters.Add('status', @{name = $StatusName}) }

# Add internal and resolution note if one provided
if ($InitialInternalNote -ne '') { $NewTicketParameters.Add('initialInternalAnalysis', $InitialInternalNote) }
if ($InitialResolutionNote -ne '') { $NewTicketParameters.Add('initialResolution', $InitialResolutionNote) }

# Add type, subtype, and item if one provided (nested for efficiency)
if ($TypeName -ne '') { 
	$NewTicketParameters.Add('type', @{name = $TypeName})
	if ($SubtypeName -ne '') {
		$NewTicketParameters.Add('subType', @{name = $SubtypeName})
		if ($ItemName -ne '') { $NewTicketParameters.Add('item', @{name = $ItemName}) }
	}
}

# Add level if one provided
if ($Level -ne '') { 
	$NewTicketParameters.Add('customFields', @{
		id = 35
		value = $Level
	}) 
}

# Create ticket
$NewTicket = New-CWMTicket @NewTicketParameters

## SEND DETAILS TO FLOW ##

Write-Output $NewTicket.id