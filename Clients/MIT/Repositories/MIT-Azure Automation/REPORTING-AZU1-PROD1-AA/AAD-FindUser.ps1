<#

Mangano IT - Azure Active Directory - Find User
Created by: Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$BearerToken,
	[string]$TenantUrl,
    [string]$UserId,
    [string]$UserPrincipalName,
    [bool]$LastSignInDateTime = $false,
    [bool]$SanitisedOutput = $false
)

## SCRIPT VARIABLES ##

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

## FIND USER ##

if ($UserId -ne "" -and $null -ne $UserId) {
    $ApiArguments = @{
        Uri = "https://graph.microsoft.com/v1.0/users/$UserId"
        Method = 'GET'
        Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
        UseBasicParsing = $true
    }
} else {
    $ApiArguments = @{
        Uri = "https://graph.microsoft.com/v1.0/users"
        Method = 'GET'
        Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
        UseBasicParsing = $true
        Body = @{
            '$filter' = "startswith(userPrincipalName,'$($UserPrincipalName)')"
            '$select' = 'displayName,givenName,jobTitle,mail,mobilePhone,surname,userPrincipalName,id,accountEnabled'
        }
    }
}

try {
    $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
    if ($ApiResponse.value -ne @()) {
        $User = $ApiResponse.value[0]
        Write-Warning "SUCCESS: $($User.userPrincipalName) ($($User.id) has been located."
    }
    
} catch {
    Write-Warning "INFO: $UserPrincipalName not located."
}

## GRAB LAST SIGN IN DATE AND TIME ##

if ($null -ne $User.id) {
    $ApiArguments = @{
        Uri = "https://graph.microsoft.com/beta/users/$($User.id)"
        Method = 'GET'
        Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
        UseBasicParsing = $true
        Body = @{
            '$select' = 'id,signInActivity'
        }
    }

    try {
        $SignInActivity = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
        Write-Warning "SUCCESS: $($User.userPrincipalName)'s sign in activity has been fetched.'"
    } catch {
        Write-Error "Unable to fetch $($User.userPrincipalName)'s sign in activity : $($_)"
    }
    
}

## SEND OUTPUT TO FLOW ##

# Create output that flow can always accept with no issues
if ($SanitisedOutput) {
    $Output = [ordered]@{
        UserId = if ($User.id) {$User.id} else {""}
        UserPrincipalName = if ($User.userPrincipalName) {$User.userPrincipalName} else {""}
        GivenName = if ($User.givenName) {$User.givenName} else {""}
        Surname = if ($User.surname) {$User.surname} else {""}
        JobTitle = if ($User.jobTitle) {$User.jobTitle} else {""}
        Mail = if ($User.mail) {$User.mail} else {""}
        MobilePhone = if ($User.mobilePhone) {$User.mobilePhone} else {""}
        AccountEnabled = if ($null -ne $User.accountEnabled) {$User.accountEnabled} else {$false}
        LastSignInDateTime = if ($null -ne $SignInActivity.signInActivity.lastSignInDateTime) {$SignInActivity.signInActivity.lastSignInDateTime} else {""} 
    }
} else {
    $Output = $User
}

Write-Output $Output | ConvertTo-Json -Depth 100