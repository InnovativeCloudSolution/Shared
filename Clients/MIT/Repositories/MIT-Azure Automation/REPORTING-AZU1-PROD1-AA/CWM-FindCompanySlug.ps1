<#

Mangano IT - ConnectWise Manage - Find Company Identifier (Slug)
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param (
	[string]$CompanyName,
    [int]$TicketId = 0,
    $ApiSecrets = $null
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

## GET COMPANY IDENTIFIER FROM TICKET ##

if ($TicketId -ne 0) {
    $Ticket = .\CWM-FindTicketDetails.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets | ConvertFrom-Json
    Write-Output $Ticket.company.identifier
}

## GET COMPANY IDENTIFIER FROM NAME ##
else {
    $ApiArguments = @{
        Uri = "$CWMApiUrl/company/companies"
        Method = 'GET'
        Body = @{ conditions = 'name like "'+$CompanyName+'" AND deletedFlag=false' }
        ContentType = $ContentType
        Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
        UseBasicParsing = $true
    }

    try { $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json } catch {}

    # Send slug (or lack of slug) back to flow
    if ($null -eq $ApiResponse) { Write-Output 'null' }
    else {
        if ($null -eq $ApiResponse.identifier) { 
            $ApiArguments = @{
                Uri = "$CWMApiUrl/company/companies"
                Method = 'GET'
                Body = @{ conditions = 'name = "'+$CompanyName+'" AND deletedFlag=false' }
                ContentType = $ContentType
                Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
                UseBasicParsing = $true
            }
            try { $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json } catch {}

            # Write new version of the API pull
            if ($null -eq $ApiResponse.identifier) { Write-Output 'null' }
            else { Write-Output $ApiResponse.identifier }
        } else { Write-Output $ApiResponse.identifier }
    }
}