<#

Mangano IT - Azure Key Vault - Get Secret
Created by: Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param (
    [Parameter(Mandatory=$true)]
    [Alias("Name")]
    [string]$SecretName
)

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

## VARIABLES ##

# Get key vault variables
$Secret = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name $SecretName -AsPlainText
Write-Output $Secret