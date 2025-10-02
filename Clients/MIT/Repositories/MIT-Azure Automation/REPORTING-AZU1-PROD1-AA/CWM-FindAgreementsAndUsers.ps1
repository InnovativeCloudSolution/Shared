<#

Mangano IT - ConnectWise Manage - Find ConnectWise Manage Agreement by Type
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param (
	[Parameter(Mandatory)][string]$AgreementTypes,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$ContentType = 'application/json'
$Output = @()

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## GET AGREEMENT BY TYPE ##

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

foreach ($AgreementType in $AgreementTypes.split(';')) {
	$ApiArguments = @{
		Uri = "$CWMApiUrl/finance/agreements"
		Method = 'GET'
		Body = @{ 
			conditions = "type/name = '$AgreementType' AND agreementStatus = 'Active'"
			fields = 'id,name,type,company,contact,agreementStatus'
		}
		ContentType = $ContentType
		Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
		UseBasicParsing = $true
	}
	
	# Get agreements
	try { $Agreements = Invoke-WebRequest @ApiArguments | ConvertFrom-Json } catch {}

	# Get user count per agreement, then add to big list
	foreach ($Agreement in $Agreements) {
		# Sculpt output
		$AgreementUpdated = @{
			Name = $Agreement.name
			Id = $Agreement.id
			Status = $Agreement.agreementStatus
			Company = $Agreement.company.name
			CompanyId = $Agreement.company.id
			Contact = $Agreement.contact.name
			ContactId = $Agreement.contact.id
			UserAdditions = @()
		}

		# Get active additions
		$AgreementId = $Agreement.id
		$GetAdditionsArguments = @{
			Uri = "$CWMApiUrl/finance/agreements/$AgreementId/additions"
			Method = 'GET'
			Body = @{ 
				conditions = "product/identifier like '*user*' AND agreementStatus = 'Active'"
				fields = 'id,product,quantity,billedQuantity,agreementStatus'
			}
			ContentType = $ContentType
			Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
			UseBasicParsing = $true
		}
		$Additions = Invoke-WebRequest @GetAdditionsArguments | ConvertFrom-Json

		# Add additions to sculpted output
		foreach ($Addition in $Additions) {
			$AgreementUpdated.UserAdditions += @{
				Product = $Addition.product.identifier
				AdditionId = $Addition.Id
				Quantity = $Addition.quantity
				BilledQuantity = $Addition.billedQuantity
				AdditionStatus =  $Addition.agreementStatus
			}
		}

		# Add agreement to output if it has users
		if ($AgreementUpdated.UserAdditions -ne @()) { $Output += $AgreementUpdated }
	}
}

## WRITE OUTPUT TO FLOW ##

Write-Output $Output | ConvertTo-Json -Depth 100