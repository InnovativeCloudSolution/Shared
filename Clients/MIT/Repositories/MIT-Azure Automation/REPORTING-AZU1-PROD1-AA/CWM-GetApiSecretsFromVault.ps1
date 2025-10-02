<#

Mangano IT - ConnectWise Manage - Get ConnectWise Manage Secrets from Azure Key Vault
Created by: Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

## SCRIPT VARIABLES ##

$AzKeyVaultName = Get-AutomationVariable -Name 'AzKeyVaultName'

## CONNECT TO AZURE KEY VAULT ##

try {
    # Connect to Azure using Managed Identity
    Connect-AzAccount -Identity | Out-Null
} catch {
    Write-Error -Message $_.Exception.Message
    throw $_.Exception
}

## GET API VARIABLES ##

# Keys from the Azure key vault
$CWMApiUrl = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWMApi-Url' -AsPlainText
$CWMApiPublicKey = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWMApi-PublicKey' -AsPlainText
$CWMApiPrivateKey = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWMApi-PrivateKey' -AsPlainText
$CWMApiClientId = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWMApi-ClientId' -AsPlainText

# API call variables
$CWMApiCredentials = "$($CWMApiPublicKey):$($CWMApiPrivateKey)"
$CWMApiEncodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($CWMApiCredentials))
$CWMApiAuthentication = "Basic $CWMApiEncodedCredentials"
$EncryptedAuthentication = .\AzAuto-EncryptString.ps1 -String $CWMApiAuthentication

## SEND BACK TO FLOW ##

$Output = @{
    Url = $CWMApiUrl
    ClientId = $CWMApiClientId
    Authentication = $EncryptedAuthentication
}

Write-Output $Output | ConvertTo-Json