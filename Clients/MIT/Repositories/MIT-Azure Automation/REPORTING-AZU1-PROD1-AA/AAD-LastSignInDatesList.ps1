<#

Mangano IT - Azure Active Directory - Get Last Sign In Dates
Created by: Gabriel Nugent
Version: 1.0.5

This runbook is designed to be run in conjunction with a Power Automate flow.

#>

param(
	[string]$BearerToken,
    [string]$TenantUrl,
    [bool]$ExcludeRecentlyCreated=$false
)

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    $Log += "Bearer token not supplied. Getting bearer token for $TenantUrl...`n`n"
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

# Form request headers with the acquired $AccessToken
$WebRequestHeaders = @{'Content-Type'="application\json";'Authorization'="Bearer $BearerToken"}
 
# This request get users list with signInActivity.
$ApiUrl = "https://graph.microsoft.com/beta/users?`$select=id,displayName,userPrincipalName,signInActivity,accountEnabled,createdDateTime,jobTitle,department,companyName&`$expand=manager(`$levels=1)"
$Result = @()
while ($Null -ne $ApiUrl) { # Perform pagination if next page link (odata.nextlink) returned.
    $Response = Invoke-WebRequest -UseBasicParsing -Method GET -Uri $ApiUrl -ContentType "application\json" -Headers $WebRequestHeaders | ConvertFrom-Json
    if ($Response.value) {
        $Users = $Response.value
        foreach ($User in $Users) {
            # Format date for PowerShell check
            if ($null -ne $User.signInActivity.lastSignInDateTime) {
                $LastSignInDate = [DateTime]$User.signInActivity.lastSignInDateTime
            } else { $LastSignInDate = $null }

            if ($null -ne $User.createdDateTime) {
                $CreatedDate = [DateTime]$User.createdDateTime
            } else { $CreatedDate = $null }

            # Add user to list if they haven't signed in in the last three months
            if (($LastSignInDate -lt (Get-Date).AddMonths(-3) -or $null -eq $LastSignInDate) -and $User.accountEnabled -eq $True) {
                if ($ExcludeRecentlyCreated){
                    if ($CreatedDate -lt (Get-Date).AddMonths(-3) -or $null -eq $CreatedDate){
                        $Result += New-Object PSObject -property $([ordered]@{ 
                            DisplayName = $User.displayName
                            UserPrincipalName = $User.userPrincipalName
                            LastSignInDateTime = if ($LastSignInDate) {
                                ($LastSignInDate.ToString("yyyy-MM-dd HH:mm:ss"))
                            } else {""}
                            CreatedDateTime = if ($CreatedDate) {
                                ($CreatedDate.ToString("yyyy-MM-dd HH:mm:ss"))
                            } else {""}
                            IsEnabled = $User.accountEnabled
                            Manager = $User.manager.userPrincipalName
                            Title = $User.jobTitle
                            Department = $User.Department
                            CompanyName = $User.CompanyName
                        })
                    }
                }else{
                    $Result += New-Object PSObject -property $([ordered]@{ 
                        DisplayName = $User.displayName
                        UserPrincipalName = $User.userPrincipalName
                        LastSignInDateTime = if ($LastSignInDate) {
                            ($LastSignInDate.ToString("yyyy-MM-dd HH:mm:ss"))
                        } else {""}
                        CreatedDateTime = if ($CreatedDate) {
                            ($CreatedDate.ToString("yyyy-MM-dd HH:mm:ss"))
                        } else {""}
                        IsEnabled = $User.accountEnabled
                        Manager = $User.manager.userPrincipalName
                        Title = $User.jobTitle
                        Department = $User.Department
                        CompanyName = $User.CompanyName
                    })
                }
            }
        }
    }
    $ApiUrl = $Response.'@odata.nextlink'
}

Write-Output $Result | ConvertTo-Json