<#

Mangano IT - Azure Active Directory - Replace Part of Titles
Created by: Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a PowerShell script.

#>

param(
    [string]$BearerToken,
    [string]$TenantUrl,
    [string]$OldTitleSegment,
    [string]$NewTitleSegment
)

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

# Get users
$Users = .\AAD-GetListOfUsers.ps1 -BearerToken $BearerToken -TenantUrl $TenantUrl -JobTitle

# Sort through users and replace part of title
foreach ($User in $Users) {
    if ($User.jobTitle -like "*$OldTitleSegment*") {
        $UserDisplayName = $User.displayName
        $UserId = $User.id
        Write-Warning "Old title located for $UserDisplayName"
        $NewTitle = $User.jobTitle.replace($OldTitleSegment, $NewTitleSegment)

        # Build API arguments and complete request
        $ApiArguments = @{
            Uri = "https://graph.microsoft.com/v1.0/users/$UserId"
            Method = 'PATCH'
            Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
            UseBasicParsing = $true
            Body = @{
                jobTitle = $NewTitle
            } | ConvertTo-Json -Depth 100
        }

        try {
            Invoke-WebRequest @ApiArguments
            Write-Warning "SUCCESS: Changed title for $UserDisplayName to $NewTitle."
        } catch {
            Write-Error "Unable to change title for $UserDisplayName to $NewTitle : $_"
        }
    }
}