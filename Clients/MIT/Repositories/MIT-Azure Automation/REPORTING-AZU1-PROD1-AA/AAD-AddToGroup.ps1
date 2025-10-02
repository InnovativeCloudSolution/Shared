<#

Mangano IT - Azure Active Directory - Add User to Group (AAD)
Created by: Gabriel Nugent, Liam Adair
Version: 1.5

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
	[string]$BearerToken,
    [string]$TenantUrl,
    [string]$UserPrincipalName = '',
    [string]$UserId = '',
    [string]$SecurityGroupName = '',
    [string]$SecurityGroupId = ''
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

## CHECKS AND BALANCES FOR USER/GROUP DETAILS ##

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    $Log += "Bearer token not supplied. Getting bearer token for $TenantUrl...`n`n"
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

# If UserId is null, fetch UserId via Graph API request
if ($UserId -eq '') {
    $GetUserArguments = @{
        Uri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName"
        Method = 'GET'
        Headers = @{ 'Authorization'="Bearer $BearerToken";'ConsistencyLevel'='eventual' }
        UseBasicParsing = $true
    }
    
    try {
        $Log += "Attempting to get user details for $UserPrincipalName..."
        $GetUserResponse = Invoke-WebRequest @GetUserArguments | ConvertFrom-Json
        $UserId = $GetUserResponse.id
        $Log += "`nSUCCESS: Fetched user details for $UserPrincipalName.`nUser ID: $UserId`n`n"
        Write-Warning "User ID: $UserId"
    } catch {
        $Log += "`nERROR: Unable to fetch user details for $UserPrincipalName.`nERROR DETAILS: " + $_ + "`n`n"
        Write-Error "Unable to fetch user details for $UserPrincipalName : $_"
    }
}

# If SecurityGroupId is null, fetch SecurityGroupId via Graph API request
if ($SecurityGroupId -eq '') {
    $GetSecurityGroupArguments = @{
        Uri = 'https://graph.microsoft.com/v1.0/groups/?$search="displayName:' + $SecurityGroupName + '"&$select=id,displayName'
        Method = 'GET'
        Headers = @{ 'Content-Type'="application/json";'Authorization'="Bearer $BearerToken";'ConsistencyLevel'='Eventual' }
        UseBasicParsing = $true
    }
    
    try {
        $Log += "Attempting to get group details for $SecurityGroupName..."
        $GetSecurityGroupResponse = Invoke-WebRequest @GetSecurityGroupArguments | ConvertFrom-Json

        # If value returned has multiple IDs, take the first one
        foreach ($Group in $GetSecurityGroupResponse.value) {
            if ($Group.displayName -eq $SecurityGroupName) {
                $SecurityGroupId = $Group.id
                $Log += "`nSUCCESS: Fetched group details for $SecurityGroupName.`nGroup ID: $SecurityGroupId`n`n"
                Write-Warning "Group ID: $SecurityGroupId"
            }
        }
    } catch {
        $Log += "`nERROR: Unable to fetch group details for $SecurityGroupName.`nERROR DETAILS:" + $_ + "`n`n"
        Write-Error "Unable to fetch group details for $SecurityGroupName : $_"
    }
}

# Check if user is already in group
$CheckUserArguments = @{
    Uri = "https://graph.microsoft.com/v1.0/groups/$SecurityGroupId/members/$UserId"
    Method = 'GET'
    Headers = @{ 'Content-Type'="application/json";'Authorization'="Bearer $BearerToken";'ConsistencyLevel'='Eventual' }
    UseBasicParsing = $true
}

try {
    $Log += "Checking to see if $UserId is a member of $SecurityGroupId...`n"
    $CheckUserResponse = Invoke-WebRequest @CheckUserArguments | ConvertFrom-Json
    $Log += "INFO: $UserId is already a member of $SecurityGroupId.`n"
    Write-Warning "INFO: $UserId is already a member of $SecurityGroupId."
} catch {
    $Log += "INFO: $UserId is not a member of $SecurityGroupId.`n"
    Write-Warning "INFO: $UserId is not a member of $SecurityGroupId."
    $CheckUserResponse = $null
}

## SCRIPT VARIABLES ##
$Result = $false

$ApiArguments = @{
    Uri = "https://graph.microsoft.com/v1.0/groups/$SecurityGroupId/members/" + '$ref'
    Method = 'POST'
    Headers = @{ 'Content-Type'="application/json";'Authorization'="Bearer $BearerToken";'ConsistencyLevel'='eventual' }
    Body = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId" } | ConvertTo-Json
    UseBasicParsing = $true
}

## ADD TO SECURITY GROUP ##

# Contacts Graph API to add member to group - only runs if user was not already found
if ($null -eq $CheckUserResponse) {
    try {
        $ApiResponse = Invoke-WebRequest @ApiArguments
    } catch {
        $Result = $false
        $Log += "ERROR: Unable to add $UserId to $SecurityGroupId.`nERROR DETAILS:" + $_
        Write-Error "Unable to add $UserId to $SecurityGroupId : $_"
    }
    if ($ApiResponse.StatusCode -eq 204) { 
        $Log += "SUCCESS: $UserId has been added to $SecurityGroupId."
        Write-Warning "SUCCESS: $UserId has been added to $SecurityGroupId."
        $Result = $true 
    }
} else {
    $Log += "INFO: $UserId is already a member of $SecurityGroupId."
    $Result = $true 
}

## SEND DETAILS BACK TO FLOW ##

$Output = @{
    Result = $Result
    SecurityGroupId = $SecurityGroupId
    UserId = $UserId
    Log = $Log
}

Write-Output $Output | ConvertTo-Json