<#

Mangano IT - ConnectWise Manage - Find Company ID
Created by: Gabriel Nugent
Version: 1.1

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

## GET COMPANY ID FROM TICKET ##

if ($TicketId -ne 0) {
    $Ticket = .\CWM-FindTicketDetails.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets | ConvertFrom-Json
    Write-Output $Ticket.company.id
}

## GET COMPANY ID FROM NAME ##
else {
	$CWMApiUrl = $ApiSecrets.Url
	$CWMApiClientId = $ApiSecrets.ClientId
	$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

	$ApiArguments = @{
		Uri = "$CWMApiUrl/company/companies"
		Method = 'GET'
		Body = @{ conditions = 'name like "'+$CompanyName+'" AND deletedFlag=false' }
		ContentType = $ContentType
		Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
		UseBasicParsing = $true
	}

	## GET COMPANY ID ##

	try { $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json } catch {}

	# Send ID (or lack of ID) back to flow
	if ($null -eq $ApiResponse) { Write-Output 0 }
	else {
		if ($null -eq $ApiResponse.id) { 
			$ApiArguments = @{
				Uri = "$CWMApiUrl/company/companies"
				Method = 'GET'
				Body = @{ conditions = 'name = "'+$CompanyName+'" AND deletedFlag=false' }
				ContentType = $ContentType
				Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$ApiAuthentication }
				UseBasicParsing = $true
			}
			try { $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json } catch {}

			# Write new version of the API pull
			if ($null -eq $ApiResponse.id) { Write-Output 0 }
			else { Write-Output $ApiResponse.id }
		} else { Write-Output $ApiResponse.id }
	}
}