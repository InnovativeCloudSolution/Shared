<#

Mangano IT - Azure Automation - Refresh Encryption Key
Created by: Gabriel Nugent
Version: 1.0

#>

param(
    [Parameter(Mandatory=$true)][string]$EncryptedKey
)

## DECRYPTS KEY WITH OLD KEY ##

[string]$NewKey = .\AzAuto-DecryptString.ps1 -String $EncryptedKey

## SAVES NEW KEY IF VALID INPUT ##

# Removes invalid characters
$NewKey = $NewKey -replace '[^0-9,]', ''

# Counts comma in string
$KeyValidation = ($NewKey.ToCharArray() | Where-Object {$_ -eq ','} | Measure-Object).Count

if ($KeyValidation -eq 31) {
    try {
        Set-AutomationVariable -Name "Encryption-ByteArray" -Value $NewKey
        Write-Output $true
    } catch { Write-Output $false }
} else {
    Write-Output $false
}