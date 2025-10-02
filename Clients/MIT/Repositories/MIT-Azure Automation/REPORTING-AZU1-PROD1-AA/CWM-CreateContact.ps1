<#

Mangano IT - ConnectWise Manage - Create New ConnectWise Manage Contact
Created by: Gabriel Nugent
Version: 1.2.3

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [int]$TicketId,
    [Parameter(Mandatory=$true)][string]$FirstName,
    [Parameter(Mandatory=$true)][string]$LastName,
    [string]$Title,
    [Parameter(Mandatory=$true)][string]$Email,
    [Parameter(Mandatory=$true)][string]$Domain,
    [string]$MobilePhone,
    [string]$OfficePhone,
    [string]$AreaCode,
    [Parameter(Mandatory=$true)][string]$CompanyName,
    [string]$SiteId,
	[bool]$NonSupportUser,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

[string]$Log = ''

# Validate phone numbers
if ($MobilePhone -ne '') {
    $FormatMobilePhoneParams = @{
        PhoneNumber = $MobilePhone
        KeepSpaces = $false
        IncludePlus = $false
        IsMobileNumber = $true
    }
    if ($AreaCode -ne '') {
        $FormatMobilePhoneParams += @{ AreaCode = $AreaCode }
        $FormatMobilePhoneParams.IncludePlus = $true
    }
    $MobilePhone = .\SMS-FormatValidPhoneNumber.ps1 @FormatMobilePhoneParams
}
if ($OfficePhone -ne '') {
    $FormatOfficePhoneParams = @{
        PhoneNumber = $OfficePhone
        KeepSpaces = $false
        IncludePlus = $false
        IsMobileNumber = $false
    }
    if ($AreaCode -ne '') {
        $FormatOfficePhoneParams += @{ AreaCode = $AreaCode }
        $FormatOfficePhoneParams.IncludePlus = $true
    }
    $OfficePhone = .\SMS-FormatValidPhoneNumber.ps1 @FormatOfficePhoneParams
}

$ContentType = 'application/json'
$ContactId = 0

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## SETUP API VARIABLES ##

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

$ApiArguments = @{
    Uri = "$CWMApiUrl/company/contacts"
    Method = 'POST'
    ContentType = $ContentType
    Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
    UseBasicParsing = $true
}

## GET AND SET CONTACT DETAILS ##

# Get the company (id)
$CompanyId = .\CWM-FindCompanyId.ps1 -CompanyName $CompanyName -ApiSecrets $ApiSecrets

# Build communication items
$CommunicationItems = @()

$EmailItem = @{
    type = @{ id = 1 }
    value = $Email
    domain = $Domain
    defaultFlag = $True
    communicationType = 'Email'
}
$CommunicationItems += $EmailItem

# If given an office phone number, add that to the ConnectWise contact
if (($OfficePhone -ne "") -and ($null -ne $OfficePhone) -and ($OfficePhone -ne "null") -and ($OfficePhone -ne 0)) {
    $DirectItem = @{
        type = @{ id = 2 }
        value = $OfficePhone
        defaultFlag = $False
        communicationType = 'Phone'
    }
    $CommunicationItems += $DirectItem
}

# If given a mobile phone number, add that to the ConnectWise contact
if (($MobilePhone -ne "") -and ($null -ne $MobilePhone) -and ($MobilePhone -ne "null") -and ($MobilePhone -ne 0)) {
    $MobileItem = @{
        type = @{ id = 4 }
        value = $MobilePhone
        defaultFlag = $True
        communicationType = 'Phone'
    }
    $CommunicationItems += $MobileItem  
}

# Sets the contact type based on the bool given
$ContactTypes = @()
if ($NonSupportUser) { 
    $ContactType = @{ id = 27 }
    $ContactTypes += $ContactType
} else {
    $ContactType = @{ id = 10 }
    $ContactTypes += $ContactType
}

# Creates the body for the API request
$ApiBody = @{
    firstName = $FirstName
    lastName = $LastName
    company = @{ id = $CompanyId }
    site = @{ id = $SiteId }
    types = $ContactTypes
}

# If given a title, add that to the contact
if ($Title -ne '') { $ApiBody += @{ title = $Title } }

## CREATE CONTACT ##

# Checks to see if the contact already exists
$Log += "Checking to see if CWM Contact for $FirstName $LastName - $Title exists...`n"

# Enclose in quotes so that the contact check accepts the string
$GetUserArguments = @{
    Uri = "$CWMApiUrl/company/contacts"
    Method = 'GET'
    Headers = $ApiArguments.Headers
    Body = @{ conditions = 'firstName like "'+$FirstName+'" AND lastName like "'+$LastName+'" AND company/id = '+$CompanyId+' AND site/id = '+$SiteId }
    ContentType = 'application/json'
    UseBasicParsing = $true
}
$GetUserResponse = Invoke-WebRequest @GetUserArguments | ConvertFrom-Json

# Updates the API call if the contact already exists
if ($null -eq $GetUserResponse.id) { 
    Write-Warning "CWM Contact for $FirstName $LastName - $Title does not already exist."
    $Log += "Duplicate contact does not exist. Creating CWM Contact for $FirstName $LastName - $Title...`n" 
    $ApiBody += @{ communicationItems = $CommunicationItems } # Cannot have comm items for when the contact already exists
} 
else {
    Write-Warning "CWM Contact for $FirstName $LastName - $Title exists."
    $Log += "CWM Contact for $FirstName $LastName - $Title exists. Marking as active and updating details...`n"
    $ContactId = $GetUserResponse.id
    $ApiArguments.Method = 'PUT'
    $ApiArguments.Uri = "$CWMApiUrl/company/contacts/$ContactId"
}

# Add body to API arguments
$ApiArguments += @{ Body = $ApiBody | ConvertTo-Json -Depth 100 } # Depth required to hit submost hashtables

try {
    $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
	$Log += "SUCCESS: Created/updated CWM Contact for $FirstName $LastName."
    Write-Warning "SUCCESS: Created/updated CWM Contact for $FirstName $LastName."
    $ContactId = $ApiResponse.id
} catch { 
    $Log += "ERROR: Failed to create/update CWM Contact for $FirstName $LastName.`nERROR DETAILS: " + $_ 
    Write-Error "Failed to create/update CWM Contact for $FirstName $LastName : $_"
}

## UPDATE RELATED TASK IF TICKET ID PROVIDED ##

if ($TicketId -ne 0 -and $ContactId -ne 0) {
    # Collate info for task resolution
    $TaskOutput = @{
        CompanyId = $CompanyId
	    ContactId = $ContactId
    }

    # Convert info to JSON
    [string]$TaskOutputJson = $TaskOutput | ConvertTo-Json

    # Update task with info
    $Task = .\CWM-UpdateTask.ps1 -TicketId $TicketId -Note 'Create ConnectWise Contact' -Resolution $TaskOutputJson -ClosedStatus $true `
    -ApiSecrets $ApiSecrets
    $Log += "`n`n" + $Task.Log

    # Add note to ticket
    $Note = .\CWM-AddTicketNote.ps1 -TicketId $TicketId -Text "A contact has been created for $FirstName $LastName.`n`nContact ID: $ContactId" `
    -ResolutionFlag $true -ApiSecrets $ApiSecrets
    $Log += "`n`n" + $Note.Log
} elseif ($TicketId -ne 0 -and $ContactId -eq 0) {
    # Add note to ticket
    $Note = .\CWM-AddTicketNote.ps1 -TicketId $TicketId -Text "A contact has not been created for $FirstName $LastName." -InternalFlag $true `
    -IssueFlag $true -ApiSecrets $ApiSecrets
    $Log += "`n`n" + $Note.Log
}

## SEND DETAILS TO FLOW ##

$Output = @{	
	CompanyId = $CompanyId
	ContactId = $ContactId
	Log = $Log
}

Write-Output ($Output | ConvertTo-Json)