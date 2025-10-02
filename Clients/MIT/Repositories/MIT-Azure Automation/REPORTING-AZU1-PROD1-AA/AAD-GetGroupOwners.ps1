<#

Mangano IT - Azure Active Directory - Get Group Owners
Created by: Gabriel Nugent
Version: 1.0.2

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$BearerToken,
	[string]$TenantUrl,
    [string]$DisplayName,
    [string]$GroupId
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    $Log += "Bearer token not supplied. Getting bearer token for $TenantUrl...`n`n"
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

## GET GROUP ##

$Group = .\AAD-GetGroup.ps1 -BearerToken $BearerToken -GroupId $GroupId -DisplayName $DisplayName | ConvertFrom-Json

## GET OWNERS FOR GROUP ##

$ApiArguments = @{
    Uri = "https://graph.microsoft.com/v1.0/groups/$($Group.id)/owners"
    Method = 'GET'
    Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken";'ConsistencyLevel'='eventual'}
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
} else {
    $Output += @{
        id = $Group.id
        displayName = $Group.displayName
        groupTypes = $Group.groupTypes
        membershipRule = $Group.membershipRule
        onPremisesSyncEnabled = $Group.onPremisesSyncEnabled
        owners = $null
    }
}

## WRITE OUTPUT TO FLOW ##

Write-Output $Output | ConvertTo-Json -Depth 100