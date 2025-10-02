param(
    [Parameter(Mandatory=$true)][string]$UserPrincipalName,
    [string]$BearerToken,
    [string]$TenantUrl,
    [int]$PasswordLength = 12,
    [int]$DinopassCalls = 1
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

# Generate a new password
$Password = .\GenerateSecurePassword.ps1 -Length $PasswordLength -DinopassCalls $DinopassCalls
$Log += "INFO: Password has been generated, and will be only provided via script output.`n`n"
Write-Warning "Password: $Password"

## RESET PASSWORD ##

# Create the body for the password reset
$ApiBody = @{
    passwordProfile = @{
        forceChangePasswordNextSignIn = $true
        forceChangePasswordNextSignInWithMfa = $true
        password = $Password
    }
}

# API request arguments
$ApiArguments = @{
    Uri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName"
    Method = 'PATCH'
    Headers = @{
        'Content-Type' = "application/json"
        'Authorization' = "Bearer $BearerToken"
    }
    Body = $ApiBody | ConvertTo-Json -Depth 10
    UseBasicParsing = $true
}

try {
    $Log += "Attempting to reset password for $UserPrincipalName...`n"
    Invoke-WebRequest @ApiArguments | Out-Null
    $Log += "SUCCESS: Password for $UserPrincipalName has been reset.`n"
    Write-Warning "SUCCESS: Password for $UserPrincipalName has been reset."
    $Result = $true
} catch {
    $Log += "ERROR: Failed to reset password for $UserPrincipalName.`nERROR DETAILS: $_"
    Write-Error "ERROR: Failed to reset password for $UserPrincipalName : $_"
    $Result = $false
}

## OUTPUT RESULTS ##

$Output = @{
    Result = $Result
    UserPrincipalName = $UserPrincipalName
    Password = if ($Result) { $Password } else { "" }
    Log = $Log
}

Write-Output $Output | ConvertTo-Json
