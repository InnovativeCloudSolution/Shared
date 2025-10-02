<#

Mangano IT - Azure Automation - Encrypt String
Created by: Gabriel Nugent
Version: 1.0

This script is not to be called from Power Automate, as it defeats the purpose.

#>

param(
    [Parameter(Mandatory=$true)][string]$String
)

## ENCRYPTION VARIABLES ##

$ByteArray = Get-AutomationVariable -Name "Encryption-ByteArray"
[Byte[]]$Key = $ByteArray -split ','

## ENCRYPT STRING ##

try {
    $SecureString = ConvertTo-SecureString $String -AsPlainText -Force
    $EncryptedString = ConvertFrom-SecureString -SecureString $SecureString -Key $Key
} catch { throw $_ }

## SEND BACK TO MAIN SCRIPT ##

Write-Output $EncryptedString