<#

Mangano IT - ConnectWise Manage - Find Contact (Enhanced Edition)
Created by: Gabriel Nugent
Version: 1.0.2

This runbook is designed to be used in conjunction with a Power Automate flow.

This will spit out easier to read values, so that the output doesn't have to be iterated over for sub-items (e.g. mobile numbers).

#>

param (
    [int]$TicketId,
	[string]$CompanyName,
    [string]$EmailAddress,
    [string]$FirstName,
    [string]$LastName,
    $ApiSecrets = $null
)

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## DEFINE POSSIBLE VARIABLES ##

$Contact = [ordered]@{
    Id = 0
    Active = $false
    FirstName = ''
    LastName = ''
    EmailAddress = ''
    Title = ''
    Company = ''
    CompanyId = 0
    Site = ''
    SiteId = 0
    MobilePhone = ''
    DirectPhone = ''
    ContactTypes = ''
}

## FETCH CONTACT DETAILS DEPENDING ON INFO PROVIDED ##

if ($TicketId -ne 0) {
    $FetchedContact = .\CWM-FindContactFromTicket.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets | ConvertFrom-Json
} elseif ($EmailAddress -ne '') {
    $FetchedContact = .\CWM-FindContactByEmail.ps1 -EmailAddress $EmailAddress -CompanyName $CompanyName -ApiSecrets $ApiSecrets | ConvertFrom-Json
} elseif ($FirstName -ne '' -and $LastName -ne '') {
    $FetchedContact = .\CWM-FindContact.ps1 -FirstName $FirstName -LastName $LastName -CompanyName $CompanyName -ApiSecrets $ApiSecrets
} else {
    $FetchedContact = $null
}

## SIFT THROUGH CONTACT INFO AND UPDATE OUTPUT ##

if ($null -ne $FetchedContact.id) {
    $Contact.Id = $FetchedContact.id
}

if ($null -ne $FetchedContact.inactiveFlag) {
    if (!$FetchedContact.inactiveFlag) {
        $Contact.Active = $true
    } else {
        $Contact.Active = $false
    }
}

if ($null -ne $FetchedContact.firstName) {
    $Contact.FirstName = $FetchedContact.firstName
}

if ($null -ne $FetchedContact.lastName) {
    $Contact.LastName = $FetchedContact.lastName
}

if ($null -ne $FetchedContact.title) {
    $Contact.Title = $FetchedContact.title
}

if ($null -ne $FetchedContact.company.id) {
    $Contact.CompanyId = $FetchedContact.company.id
    $Contact.Company = $FetchedContact.company.name
}

if ($null -ne $FetchedContact.site.id) {
    $Contact.SiteId = $FetchedContact.site.id
    $Contact.Site = $FetchedContact.site.name
}

# Set phone numbers and email from communication items
foreach ($Item in $FetchedContact.communicationItems) {
    if ($Item.type.name -match 'Email' -and $Item.defaultFlag) {
        $Contact.EmailAddress = $Item.value
        continue
    }

    if ($Item.type.name -eq 'Mobile') {
        $Contact.MobilePhone = $Item.value
        continue
    }

    if ($Item.type.name -eq 'Direct') {
        $Contact.DirectPhone = $Item.value
        continue
    }
}

# Add types to list
foreach ($Type in $FetchedContact.types) {
    if ($Contact.ContactTypes -ne '') {
        $Contact.ContactTypes += ', '
    }

    $Contact.ContactTypes += $Type.Name
}

## OUTPUT TO FLOW ##

Write-Output $Contact | ConvertTo-Json -Depth 100