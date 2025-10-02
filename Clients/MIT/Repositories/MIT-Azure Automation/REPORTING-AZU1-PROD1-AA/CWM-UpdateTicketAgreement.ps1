<#

Mangano IT - ConnectWise Manage - Update Ticket Agreement
Created by: Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow or an Azure Automation script.

#>

param (
    [Parameter(Mandatory=$true)][int]$TicketId,
    [Parameter(Mandatory=$true)][string]$AgreementName,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

[string]$Log = ''
$ContentType = 'application/json'
$AgreementId = 0
$Result = $false

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## SETUP API VARIABLES ##

$CWMApiUrl = $ApiSecrets.Url
$CWMApiClientId = $ApiSecrets.ClientId
$CWMApiAuthentication = .\AzAuto-DecryptString.ps1 -String $ApiSecrets.Authentication

## GET TICKET DETAILS ##

$Ticket = .\CWM-FindTicketDetails.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets | ConvertFrom-Json

## GET COMPANY AGREEMENTS ##

$Agreements = .\CWM-FindAgreementsForCompany.ps1 -CompanyName $Ticket.company.name -ApiSecrets $ApiSecrets | ConvertFrom-Json

# Find right agreement
foreach ($Agreement in $Agreements) {
    if ($Agreement.name -eq $AgreementName) {
        $AgreementId = $Agreement.id
        $Log += "INFO: Found matching agreement for $AgreementName.`nAgreement ID: $AgreementId`n`n"
        Write-Warning "Agreement ID: $AgreementId"
    }
}

## UPDATE TICKET AGREEMENT ##

# If ticket isn't empty, update the agreement
if ($null -ne $Ticket -and $AgreementId -ne 0 -and $Ticket.agreement.id -ne $AgreementId) {
    $ApiBody = @(
        @{
            op = 'replace'
            path = '/agreement/id'
            value = $AgreementId
        }
    )

    $ApiArguments = @{
        Uri = "$CWMApiUrl/service/tickets/$TicketId"
        Method = 'PATCH'
        Body = ConvertTo-Json -InputObject $ApiBody -Depth 100
        ContentType = $ContentType
        Headers = @{ 'clientId'=$CWMApiClientId;'Authorization'=$CWMApiAuthentication }
        UseBasicParsing = $true
    }

    try {
        $Log += "Updating #$TicketId's agreement to $AgreementId...`n"
        Invoke-WebRequest @ApiArguments | Out-Null
        $Log += "SUCCESS: #$TicketId's agreement has been set to $AgreementId."
        Write-Warning "SUCCESS: #$TicketId's agreement has been set to $AgreementId."
        $Result = $true
    } catch {
        $Log += "ERROR: Unable to set #$TicketId's agreement to $AgreementId.`nERROR DETAILS: " + $_
        Write-Error "Unable to set #$TicketId's agreement to $AgreementId : $_"
        $Result = $false
    }
} elseif ($Ticket.agreement.id -eq $AgreementId) {
    $Log += "INFO: Agreement already matches."
    Write-Warning "Agreement already matches."
    $Result = $true
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Result = $Result
    Log = $Log
}

Write-Output $Output | ConvertTo-Json