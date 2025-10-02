<#

Mangano IT - ConnectWise Automate - Get Bearer Token
Created by: Gabriel Nugent
Version: 1.0.1

#>

param(
    [bool]$Encrypted = $true
)

## SCRIPT VARIABLES ##

$AzKeyVaultName = Get-AutomationVariable -Name 'AzKeyVaultName'
$ContentType = 'application/json'
$Accept = 'application/json'

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
$ClientId = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWAApi-ClientId' -AsPlainText
$Url = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWAApi-Url' -AsPlainText
$Username = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWAAPI-Username' -AsPlainText
$Password = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWAApi-Password' -AsPlainText

# Build request variables
$ApiArguments = @{
	Uri = "$($Url)/apitoken"
	Method = 'POST'
	Headers = @{
        'Content-Type' = $ContentType
        Accept = $Accept
        ClientId = $ClientId
    }
	Body = @{
        Username = $Username
        Password = $Password
    } | ConvertTo-Json -Depth 100
    UseBasicParsing = $true
}

## GET TOKEN ##

# Completes an API call to get the bearer token
try { $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json } catch {}

# Convert access token to secure string
if (!$Encrypted) { $AccessToken = $ApiResponse.AccessToken }
else { $AccessToken = .\AzAuto-EncryptString.ps1 -String $ApiResponse.AccessToken }

# Writes token as output
Write-Output $AccessToken