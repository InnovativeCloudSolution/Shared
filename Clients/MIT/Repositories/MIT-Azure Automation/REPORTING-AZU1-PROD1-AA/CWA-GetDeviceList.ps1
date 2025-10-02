<#

Mangano IT - ConnectWise Automate - Get List of Devices
Created by: Gabriel Nugent
Version: 1.0.3

#>

param(
    [Parameter(Mandatory)][string]$ClientName,
    [string]$IncludeFields = 'Id,Client,Location,ComputerName,LastUserName,Type',
    [string]$DeviceType = 'Workstation',
    [string]$BearerToken
)

## SCRIPT VARIABLES ##

$AzKeyVaultName = Get-AutomationVariable -Name 'AzKeyVaultName'

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    Write-Warning "Bearer token not supplied. Getting bearer token..."
	$EncryptedBearerToken = .\CWA-GetBearerToken.ps1
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

## CONNECT TO AZURE KEY VAULT ##

try {
    # Connect to Azure using Managed Identity
    Connect-AzAccount -Identity | Out-Null
} catch {
    Write-Error -Message $_.Exception.Message
    throw $_.Exception
}

# Keys from the Azure key vault
$CWAApiUrl = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWAApi-Url' -AsPlainText
$CWAApiClientId = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWAApi-ClientId' -AsPlainText

## GET LIST OF DEVICES ##

$Arguments = @{
    Uri = "$($CWAApiUrl)/computers"
    Method = 'GET'
    Headers = @{
        'Content-Type' = "application/json"
        'ClientId' = $CWAApiClientId
        'Authorization' = "Bearer $BearerToken"
    }
    Body = @{
        condition = "Client/Name eq '$($ClientName)' AND Type eq '$($DeviceType)'"
        includefields = $IncludeFields
        pagesize = 1000000
    }
    UseBasicParsing = $true
}

try {
    $ApiResponse = Invoke-WebRequest @Arguments
    Write-Warning "SUCCESS: Fetched device list from CW Automate."
    $DeviceList = $ApiResponse.Content  # This export is already formatted as JSON, no need to convert
} catch {
    Write-Error "Unable to fetch device list from Automate : $($_)"
}

## SEND OUTPUT TO FLOW ##

if ($null -ne $DeviceList) {
    Write-Output $DeviceList
}