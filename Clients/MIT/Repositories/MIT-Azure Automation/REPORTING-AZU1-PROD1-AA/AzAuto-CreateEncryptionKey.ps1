<#

Mangano IT - Azure Automation - Create Encryption Key
Created by: Gabriel Nugent
Version: 1.0

Exports the key in an encrypted state, using the old key.

#>

## SCRIPT VARIABLES ##

$MinimumValue = 0
$MaximumValue = 250
[string]$NewKey = Get-Random -Minimum $MinimumValue -Maximum $MaximumValue
$KeyValidation = 0

## CREATE KEY ##

# Adds random numbers until there are 32 separate numbers (31 commas)
while ($KeyValidation -lt 31) {
    $NewKey += ',' + (Get-Random -Minimum $MinimumValue -Maximum $MaximumValue)
    $KeyValidation = ($NewKey.ToCharArray() | Where-Object {$_ -eq ','} | Measure-Object).Count
}

## ENCRYPT NEW KEY ##

$EncryptedKey = .\AzAuto-EncryptString.ps1 -String $NewKey
Write-Output $EncryptedKey