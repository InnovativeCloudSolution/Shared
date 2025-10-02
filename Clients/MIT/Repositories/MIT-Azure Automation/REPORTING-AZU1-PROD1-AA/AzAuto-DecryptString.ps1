<#

Mangano IT - Azure Automation - Decrypt String
Created by: Gabriel Nugent
Version: 1.0

This script is not to be called from Power Automate, as it defeats the purpose.

#>

param(
    [Parameter(Mandatory=$true)][Alias("String")][string]$EncryptedString
)

## ENCRYPTION VARIABLES ##

$ByteArray = Get-AutomationVariable -Name "Encryption-ByteArray"
[Byte[]]$Key = $ByteArray -split ','

## DECRYPT STRING ##

try {
    $SecureString = ConvertTo-SecureString $EncryptedString -Key $Key
    $BSTR = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    [string]$String = [Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
} catch { throw $_ }

## SEND BACK TO MAIN SCRIPT ##

Write-Output $String