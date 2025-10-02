<#

Mangano IT - ConnectWise Manage - Fetch Tickets
Created by: Gabriel Nugent
Version: 1.0.2

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param (
    [Parameter(Mandatory)][string]$Summary,
    [Parameter(Mandatory)][ValidateSet('Equals','StartsWith','EndsWith')][string]$SummaryComparison,
    [string]$CompanyName,
    [string]$LastUpdatedDateTime,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$ContentType = 'application/json'

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## SETUP API VARIABLES ##

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

# Format summary as needed
switch ($SummaryComparison) {
    'StartsWith' { $SummarySearch = "$Summary*" }
    'EndsWith' { $SummarySearch = "*$Summary" }
    Default { $SummarySearch = "$Summary" }
}

$ApiArguments = @{
    Uri = "$CWMApiUrl/service/tickets"
    Method = 'GET'
    ContentType = $ContentType
    Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
    UseBasicParsing = $true
    Body = @{
        conditions = "summary like '$SummarySearch'"
        pageSize = 1000
    }
}

# Add extra conditions if provided
if ($CompanyName -ne '') {
    $ApiArguments.Body.conditions += " AND company/name = '$CompanyName'"
}

if ($LastUpdatedDateTime -ne '') {
    $ApiArguments.Body.conditions += " AND LastUpdated >= [$LastUpdatedDateTime]"
}

## FETCH DETAILS FROM TICKET ##

try { 
    $Ticket = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
    Write-Warning "SUCCESS: Tickets fetched."
} catch { 
    Write-Error "Tickets unable to be fetched : $($_)"
    $Ticket = $null
}

## SEND DETAILS TO FLOW ##

Write-Output $Ticket | ConvertTo-Json -Depth 100