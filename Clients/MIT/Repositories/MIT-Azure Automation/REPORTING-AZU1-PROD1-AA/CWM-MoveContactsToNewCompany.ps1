<#

Mangano IT - ConnectWise Manage - Move Contacts From One ConnectWise Manage Company To Another
Created by: Gabriel Nugent
Version: 1.3

This runbook is designed to be used independently.

#>

param(
    [string]$OldCompanyName,
    [string]$NewCompanyName,
    $ApiSecrets = $null,
    [bool]$ActiveOnly = $true
)

## SCRIPT VARIABLES ##

$ContentType = 'application/json'

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## SETUP API VARIABLES ##

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

## GET COMPANY IDS ##

$OldCompanyId = .\CWM-FindCompanyId.ps1 -CompanyName $OldCompanyName -ApiSecrets $ApiSecrets
Write-Output "Old company ID: $OldCompanyId"

$NewCompanyId = .\CWM-FindCompanyId.ps1 -CompanyName $NewCompanyName -ApiSecrets $ApiSecrets
Write-Output "New company ID: $NewCompanyId`n"

## GET ALL ACTIVE CONTACTS ##

$GetContactsArguments = @{
    Uri = "$CWMApiUrl/company/contacts"
    Method = 'GET'
    Body = @{ conditions = "company/id = $OldCompanyId" }
    ContentType = $ContentType
    Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
    UseBasicParsing = $true
}

# If only active contacts, add to request
if ($ActiveOnly) {
    $GetContactsArguments.Body.conditions += ' AND inactiveFlag = false'
}

try {
    Write-Output "Getting active contacts at $OldCompanyName..."
    $Contacts = Invoke-WebRequest @GetContactsArguments | ConvertFrom-Json
    Write-Output "SUCCESS: Fetched contacts at $OldCompanyName."
} catch {
    Write-Output "ERROR: Unable to fetch contacts at $OldCompanyName."
    Write-Output "ERROR DETAILS: " $_
}

## MOVE CONTACTS ##

foreach ($Contact in $Contacts) {
    $Id = $Contact.id
    $Name = $Contact.firstName + ' ' + $Contact.lastName

    # Create request body - replace company ID with new ID
    $ApiBody = @(
        @{
            op = 'replace'
            path = '/company/id'
            value = $NewCompanyId
        }
    )

    $ApiArguments = @{
        Uri = "$CWMApiUrl/company/contacts/$Id"
        Method = 'PATCH'
        Body = ConvertTo-Json -InputObject $ApiBody -Depth 100
        ContentType = $ContentType
        Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
        UseBasicParsing = $true
    }

    try {
        Write-Output "`nAttempting to move $Name to $NewCompanyName..."
        Invoke-WebRequest @ApiArguments | Out-Null
        Write-Output "SUCCESS: Moved $Name to $NewCompanyName."
    } catch {
        Write-Output "ERROR: Unable to move $Name to $NewCompanyName."
        Write-Output "ERROR DETAILS: " $_
        break
    }
}