<#

Mangano IT - ConnectWise Manage - Batch Update Board Agreement Default
Created by: Gabriel Nugent
Version: 1.0.2

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param (
	[Parameter(Mandatory)][string]$AgreementTypes,
	[Parameter(Mandatory)][int]$BoardId,
	[Parameter(Mandatory)][string]$BoardName,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$ContentType = 'application/json'

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

## GET AGREEMENTS BY TYPE ##

$Agreements = .\CWM-FindAgreementsByType.ps1 -AgreementTypes $AgreementTypes -ApiSecrets $ApiSecrets | ConvertFrom-Json

## ADD BOARD DEFAULTS FOR EACH AGREEMENT ##

foreach ($Agreement in $Agreements) {
	# Variable to determine if default already exists
	$DefaultAlreadyExists = $false

	# Get board defaults for this agreement
	$GetBoardDefaultArguments = @{
		Uri = "$CWMApiUrl/finance/agreements/$($Agreement.id)/boardDefaults"
		Method = 'GET'
		ContentType = $ContentType
		Headers = @{
			'clientId' = $CWMApiClientId
			'Authorization' = $CWMApiAuthentication
		}
		UseBasicParsing = $true
	}
	
	try {
		$AgreementBoardDefaults = Invoke-WebRequest @GetBoardDefaultArguments | ConvertFrom-Json
		Write-Warning "SUCCESS: Fetched board defaults for $($Agreement.name) ($($Agreement.company.name))"
	} catch {
		Write-Error "Unable to fetch board defaults for $($Agreement.name) ($($Agreement.company.name)) : $($_)"
	}

	# Check to see if board requested is already covered
	foreach ($AgreementBoardDefault in $AgreementBoardDefaults) {
		if ($AgreementBoardDefault.board.id -eq $BoardId) {
			$DefaultAlreadyExists = $true
		}
	}

	# Add requested board default if not already covered
	if (!$DefaultAlreadyExists) {
		$AddBoardDefaultArguments = @{
			Uri = "$CWMApiUrl/finance/agreements/$($Agreement.id)/boardDefaults"
			Method = 'POST'
			Body = @{ 
				agreementId = $Agreement.id
				board = @{
					id = $BoardId
					name = $BoardName
				}
				defaultFlag = $true
			} | ConvertTo-Json -Depth 100
			ContentType = $ContentType
			Headers = @{
				'clientId' = $CWMApiClientId
				'Authorization' = $CWMApiAuthentication
			}
			UseBasicParsing = $true
		}

		try {
			Invoke-WebRequest @AddBoardDefaultArguments | Out-Null
			Write-Warning "SUCCESS: $($BoardName) has been added as a default board for the agreement $($Agreement.name) ($($Agreement.company.name))."
		}
		catch {
			Write-Error "Unable to add $($BoardName) as a default board for the agreement $($Agreement.name) ($($Agreement.company.name)) : $($_)"
		}
	}
}