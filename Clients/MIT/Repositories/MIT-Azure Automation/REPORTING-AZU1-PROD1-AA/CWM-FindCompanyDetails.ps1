<#

Mangano IT - ConnectWise Manage - Find Company Details
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

## GET COMPANY DETAILS FROM TICKET ##

if ($TicketId -ne 0) {
    $Ticket = .\CWM-FindTicketDetails.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets | ConvertFrom-Json
    Write-Output $Ticket.company
}

## GET COMPANY DETAILS FROM NAME ##

else {
    $CWMApiUrl = $ApiSecrets.Url
    $CWMApiClientId = $ApiSecrets.ClientId
    $CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

    $ApiArguments = @{
        Uri = "$CWMApiUrl/company/companies"
        Method = 'GET'
        Body = @{ conditions = 'name = "'+$CompanyName+'" AND deletedFlag=false' }
        ContentType = $ContentType
        Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
        UseBasicParsing = $true
    }

    ## GET COMPANY ##

    try { $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json } catch {}

    # Send details (or lack of) back to flow
    if ($null -eq $ApiResponse) { Write-Output $null }
    else { Write-Output $ApiResponse }
}