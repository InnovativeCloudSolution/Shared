<#

Mangano IT - Pax8 - Get Bearer Token
Created by: Gabriel Nugent
Version: 1.0

#>

## SCRIPT VARIABLES ##

$AzKeyVaultName = Get-AutomationVariable -Name 'AzKeyVaultName'
$ContentType = 'application/json'
$Audience = 'api://p8p.client'

## CONNECT TO AZURE KEY VAULT ##

try {
    # Connect to Azure using Managed Identity
    Connect-AzAccount -Identity | Out-Null
} catch {
    Write-Error -Message $_.Exception.Message
    throw $_.Exception
}

## VARIABLES ##

# Get key vault variables
$ClientId = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'Pax8-ClientId' -AsPlainText
$ClientSecret = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'Pax8-ClientSecret' -AsPlainText

# Build API body
$ApiBody = @{
    clientId = $ClientId
    clientSecret = $ClientSecret
    audience = $Audience       
} | ConvertTo-Json

# Build request variables
$ApiArguments = @{
	Uri = "https://token-manager.pax8.com/auth/token"
	Method = 'POST'
	ContentType = $ContentType
	Body = $ApiBody
    UseBasicParsing = $true
}

## GET TOKEN ##

# Completes an API call to get the bearer token
try {
	$ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
	$Output = $ApiResponse.token
} catch { $Output = $_}

# Writes token as output
Write-Output $Output