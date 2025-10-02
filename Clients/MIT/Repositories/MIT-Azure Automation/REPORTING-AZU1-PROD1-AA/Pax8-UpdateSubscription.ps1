<#

Mangano IT - Pax8 - Update Subscription
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$BearerToken = '',
    [Parameter(Mandatory=$true)][string]$SubscriptionId,
    [bool]$AddSubscriptions = $true,
    [int]$Quantity = 1,
    [string]$BillingTerm = 'Monthly',
    [bool]$Testing = $false
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    $Log += "Bearer token not supplied. Getting bearer token for Pax8...`n`n"
	$BearerToken = .\Pax8-GetBearerToken.ps1
}

$ContentType = 'application/json'
$ApiUrl = .\KeyVault-GetSecret.ps1 -SecretName 'Pax8-ApiUrl'

## GET CURRENT SUBSCRIPTION QUANTITY ##

$GetSubscriptionArguments = @{
    Uri = "$ApiUrl/subscriptions/$SubscriptionId"
	Method = 'GET'
    ContentType = $ContentType
    Headers = @{'Authorization'="Bearer $BearerToken"}
    UseBasicParsing = $true
}

try {
    $Log += "Fetching subscription $SubscriptionId...`n"
    $Subscription = Invoke-WebRequest @GetSubscriptionArguments | ConvertFrom-Json
    $Log += "SUCCESS: Fetched subscription $SubscriptionId.`n`n"
    Write-Warning "SUCCESS: Fetched subscription $SubscriptionId."
}
catch {
    $Log += "ERROR: Unable to fetch subscription from Pax8.`nERROR DETAILS: " + $_
    Write-Error "Unable to fetch subscription from Pax8 : $_"
    $Subscription = $null
}

## PREPARE API REQUEST ##

if ($null -ne $Subscription) {
    $CurrentQuantity = $Subscription.quantity
    if ($AddSubscriptions) { $NewQuantity = $CurrentQuantity + $Quantity }
    else { $NewQuantity = $CurrentQuantity - $Quantity }
    $Log += "INFO: Subscription quantity will change from $CurrentQuantity to $NewQuantity.`n`n"
    Write-Warning "INFO: Subscription quantity will change from $CurrentQuantity to $NewQuantity."

    # If testing, add mock query to script
    if ($Testing) { $Url = "$ApiUrl/subscriptions/$SubscriptionId?isMock=true" }
    else { $Url = "$ApiUrl/subscriptions/$SubscriptionId" }

    $ApiArguments = @{
        Uri = $Url
        Method = 'PUT'
        ContentType = $ContentType
        Headers = @{'Authorization'="Bearer $BearerToken"}
        Body = @{
            quantity = $NewQuantity
            billingTerm = $BillingTerm
            startDate = Get-Date -Format 'yyyy-MM-dd'
        } | ConvertTo-Json
        UseBasicParsing = $true
    }

    ## UPDATE CURRENT SUBSCRIPTION ##

    try {
        $Log += "Updating subscription $SubscriptionId...`n"
        Invoke-WebRequest @ApiArguments | Out-Null
        $Result = $true
        $Log += "SUCCESS: Updated subscription $SubscriptionId.`n`n"
        Write-Warning "SUCCESS: Updated subscription $SubscriptionId."
    }
    catch {
        $Log += "ERROR: Unable to update subscription.`nERROR DETAILS: " + $_
        Write-Error "Unable to update subscription : $_"
        $Result = $false
    }
} else { $Result = $false }

## SEND DETAILS TO FLOW ##

$Output = @{
    Result = $Result
    Log = $Log
}

Write-Output $Output | ConvertTo-Json