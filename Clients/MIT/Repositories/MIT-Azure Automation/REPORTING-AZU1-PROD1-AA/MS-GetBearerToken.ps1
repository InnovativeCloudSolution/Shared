<#

Mangano IT - Microsoft Graph API - Get Bearer Token for Client Tenancy
Created by: Gabriel Nugent
Version: 1.2

#>

param(
    [Parameter(Mandatory)][string]$TenantUrl,
    [bool]$ForHybridWorker = $false
)

## SCRIPT VARIABLES ##

$AzKeyVaultName = Get-AutomationVariable -Name 'AzKeyVaultName'
$ContentType = 'application/x-www-form-urlencoded'

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
$ClientId = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-PartnerAppClientID' -AsPlainText
$ClientSecret = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-PartnerAppSecret' -AsPlainText
$Scope = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-Audience' -AsPlainText
$GrantType = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-GrantType' -AsPlainText

# Build request variables
$ApiArguments = @{
	Uri = "https://login.microsoftonline.com/$TenantUrl/oauth2/v2.0/token"
	Method = 'POST'
	Headers = @{'Content-Type'=$ContentType}
	Body = "client_id=$ClientId&client_secret=$ClientSecret&scope=$Scope&grant_type=$GrantType"
    UseBasicParsing = $true
}

## GET TOKEN ##

# Completes an API call to get the bearer token
try { $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json } catch {}

# Convert access token to secure string
if ($ForHybridWorker) { $AccessToken = $ApiResponse.access_token }
else { $AccessToken = .\AzAuto-EncryptString.ps1 -String $ApiResponse.access_token }

# Writes token as output
Write-Output $AccessToken