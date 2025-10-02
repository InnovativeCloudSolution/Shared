<#

Mangano IT - Azure Active Directory - Get All Groups
Created by: Gabriel Nugent
Version: 1.1.6

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$BearerToken,
	[string]$TenantUrl,
    [bool]$GroupsWithOwners = $false
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

$Output = @()

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    $Log += "Bearer token not supplied. Getting bearer token for $TenantUrl...`n`n"
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

## GET ALL GROUPS ##

$ApiUrl = 'https://graph.microsoft.com/v1.0/groups?$select=id,displayName,groupTypes,membershipRule,onPremisesSyncEnabled'

while ($null -ne $ApiUrl) {
    $ApiArguments = @{
        Uri = $ApiUrl
        Method = 'GET'
        Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
        UseBasicParsing = $true
    }

    try {
        $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
        Write-Warning "SUCCESS: Fetched list of groups for $TenantUrl."
        $Groups += $ApiResponse.value
        $ApiUrl = $ApiResponse.'@odata.nextlink'
        Start-Sleep -Seconds 3
    } catch {
        Write-Error "Unable to fetch list of groups for $TenantUrl : $_"
    }
}

# List only groups with owners if required
if ($GroupsWithOwners) {
    foreach ($Group in $Groups) {
        $ApiArguments = @{
            Uri = "https://graph.microsoft.com/v1.0/groups/$($Group.id)/owners"
            Method = 'GET'
            Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
            UseBasicParsing = $true
        }

        # Fetch group's owners
        $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json

        # If group has owners, add to list
        if ($ApiResponse.value -ne @()) {
            $Output += @{
                id = $Group.id
                displayName = $Group.displayName
                groupTypes = $Group.groupTypes
                membershipRule = $Group.membershipRule
                onPremisesSyncEnabled = $Group.onPremisesSyncEnabled
                owners = $ApiResponse.value
            }
        }
    }
} else {
    $Output = $Groups
}

## WRITE OUTPUT TO FLOW ##

Write-Output $Output | ConvertTo-Json -Depth 100