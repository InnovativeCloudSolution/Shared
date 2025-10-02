param(
	[Parameter(Mandatory=$true)][string]$Summary,
	[Parameter(Mandatory=$true)][int]$CompanyId,
	[string]$BoardName = "HelpDesk (MS)",
	[Parameter(Mandatory=$true)][string]$InitialDescriptionNote,
	[string]$InitialInternalNote = '',
	[string]$InitialResolutionNote = '',
	[int]$ContactId = 0,
	[string]$ContactName = '',
	[int]$SiteId = 0,
	[string]$StatusName = '',
	[string]$TypeName = '',
	[string]$SubtypeName = '',
	[string]$ItemName = ''
)

# Set CRM variables, connect to server
$Server = 'https://crm.manganoit.com.au'
$Company = 'mits'
$PublicKey = 'mBemoCno7IwHElgT'
$PrivateKey = 'BmC0AR9dN7eJFouT'
$ClientId = '1208536d-40b8-4fc0-8bf3-b4955dd9d3b7'

# Create a credential object
$Connection = @{
    Server = $Server
    Company = $Company
    pubkey = $PublicKey
    privatekey = $PrivateKey
    clientId = $ClientId
}

# Connect to Manage server
Connect-CWM @Connection

# Create ticket info
# Some are parsed as hashtables based on whether or not they're sub-items in the API
$NewTicketParameters = @{
    summary = $Summary
    company = @{id = $CompanyId}
	board = @{name = $BoardName}
	initialDescription = $InitialDescriptionNote
}

# Add contact if one provided
if ($ContactId -ne 0) { $NewTicketParameters.Add('contact', @{id = $ContactId}) }
else { if ($ContactName -ne '') { $NewTicketParameters.Add('contactName', $ContactName) } }

# Add site if one provided
if ($SiteId -ne 0) { $NewTicketParameters.Add('site', @{id = $SiteId}) }

# Add status if one provided
if ($StatusId -ne '') { $NewTicketParameters.Add('status', @{name = $StatusName}) }

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

# Create ticket
$NewTicket = New-CWMTicket @NewTicketParameters

Write-Output $NewTicket.id