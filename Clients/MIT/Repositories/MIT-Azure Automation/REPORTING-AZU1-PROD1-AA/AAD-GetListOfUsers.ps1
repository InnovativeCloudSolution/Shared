<#

Mangano IT - Azure Active Directory - Get List of Users
Created by: Gabriel Nugent
Version: 1.2

This runbook is designed to be used in conjunction with a PowerShell script.

#>

param(
    [string]$BearerToken,
	[string]$TenantUrl,
    [switch]$Mail,
    [switch]$JobTitle,
    [switch]$Department,
    [switch]$CompanyName,
    [switch]$SignInActivity,
    [switch]$UserType,
    [switch]$AssignedLicenses,
    [switch]$Address
)

## SCRIPT VARIABLES ##

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

# List of users
$Users = @()

## GET LIST OF USERS ##

$ApiUrl = 'https://graph.microsoft.com/beta/users?$select=id,displayName,userPrincipalName,accountEnabled'

# Add extra params based on switch statements
if ($Mail) { $ApiUrl += ',mail' }
if ($JobTitle) { $ApiUrl += ',jobTitle' }
if ($Department) { $ApiUrl += ',department' }
if ($CompanyName) { $ApiUrl += ',companyName' }
if ($SignInActivity) { $ApiUrl += ',signInActivity' }
if ($UserType) { $ApiUrl += ',userType' }
if ($AssignedLicenses) { $ApiUrl += ',assignedLicenses' }
if ($Address) { $ApiUrl += ',streetAddress,city,state,postalCode' }

while ($null -ne $ApiUrl) {
    $ApiArguments = @{
        Uri = $ApiUrl
        Method = 'GET'
        Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
        UseBasicParsing = $true
    }

    try {
        $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
        Write-Warning "SUCCESS: Fetched list of users for $TenantUrl."
        $Users += $ApiResponse.value
        $ApiUrl = $ApiResponse.'@odata.nextlink'
    } catch {
        Write-Error "Unable to fetch list of users for $TenantUrl : $_"
    }
}

## SEND DETAILS BACK TO SCRIPT ##

Write-Output $Users