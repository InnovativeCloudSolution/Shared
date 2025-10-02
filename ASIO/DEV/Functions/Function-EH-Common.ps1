function Get-EH-Headers {
    param (
        [Parameter(Mandatory=$true)]
        [string]$EHAuthorization
    )
    return @{
        'Authorization' = $EHAuthorization
        'Content-Type'  = 'application/json'
    }
}

function Get-EH-Secrets {
    param (
        [Parameter(Mandatory=$true)]
        [string]$AzKeyVaultName
    )

    try {
        Connect-AzAccount -Identity | Out-Null
        return @{
            EHclient_Id      = (Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'SEA-EH-ClientID' -AsPlainText)
            EHclient_secret  = (Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'SEA-EH-ClientSecret' -AsPlainText)
            EHcode           = (Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'SEA-EH-Code' -AsPlainText)
            EHrefresh_token  = (Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'SEA-EH-RefreshToken' -AsPlainText)
            EHOrganizationId = (Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'SEA-EH-OrganizationId' -AsPlainText)
            EHRedirectUri    = (Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'SEA-EH-RedirectUri' -AsPlainText)
        }
    }
    catch {
        Write-MessageLog "Failed to retrieve EH Secrets: $_" -LogType "Error"
        throw $_
    }
}

function Get-EH-Authorization {
    param (
        [Parameter(Mandatory=$true)]
        [string]$EHclient_Id,
        [Parameter(Mandatory=$true)]
        [string]$EHclient_secret,
        [Parameter(Mandatory=$true)]
        [string]$EHcode,
        [Parameter(Mandatory=$true)]
        [string]$EHrefresh_token,
        [Parameter(Mandatory=$true)]
        [string]$EHRedirectUri
    )

    $ApiUri = "https://oauth.employmenthero.com/oauth2/token" + `
              "?client_id=$($EHclient_Id)" + `
              "&client_secret=$($EHclient_secret)" + `
              "&grant_type=refresh_token" + `
              "&code=$($EHcode)" + `
              "&redirect_uri=$($EHRedirectUri)" + `
              "&refresh_token=$($EHrefresh_token)"
    
    $ApiArguments = @{
        Uri             = $ApiUri
        Method          = 'POST'
        ContentType     = 'application/json'
        UseBasicParsing = $true
    }

    try {
        $Response = Invoke-RestMethod @ApiArguments
        Write-MessageLog "API Response: $($Response | ConvertTo-Json -Depth 5)"
        if (-not $Response.access_token) {
            Write-MessageLog "Access token is null or missing in the response." -LogType "Error"
            return $null
        }
        return "Bearer $($Response.access_token)"
    }
    catch {
        Write-MessageLog "Failed to retrieve access token: $_" -LogType "Error"
        return $null
    }
}

function Get-EH-EmployeeAll {
    param (
        [Parameter(Mandatory=$true)]
        [string]$EHOrganizationId,
        [Parameter(Mandatory=$true)]
        [string]$EHAuthorization
    )

    $UriString = "https://api.employmenthero.com/api/v1/organisations/$EHOrganizationId/employees"
    $AllEmployeeArguments = @{
        Uri             = $UriString
        Headers         = Get-EH-Headers -EHAuthorization $EHAuthorization
        Method          = 'GET'
        UseBasicParsing = $true
    }

    try {
        $Response = Invoke-RestMethod @AllEmployeeArguments
        if (-not $Response -or -not $Response.data.items) {
            Write-MessageLog "No employees returned." -LogType "Error"
            return $null
        }
        return $Response.data.items
    }
    catch {
        Write-MessageLog "Failed to fetch employees: $_" -LogType "Error"
        return $null
    }
}

function Get-EH-Employee {
    param (
        [Parameter(Mandatory=$true)]
        [string]$EHOrganizationId,
        [Parameter(Mandatory=$true)]
        [string]$EHAuthorization,
        [Parameter(Mandatory=$true)]
        [string]$EmployeeId
    )

    $UriString = "https://api.employmenthero.com/api/v1/organisations/$EHOrganizationId/employees/$EmployeeId"
    $EmployeeArguments = @{
        Uri             = $UriString
        Headers         = Get-EH-Headers -EHAuthorization $EHAuthorization
        Method          = 'GET'
        UseBasicParsing = $true
    }

    try {
        $Response = Invoke-RestMethod @EmployeeArguments
        if (-not $Response -or -not $Response.data) {
            Write-MessageLog "No data returned for Employee details." -LogType "Error"
            return $null
        }
        return $Response.data
    }
    catch {
        Write-MessageLog "Failed to fetch Employee details: $_" -LogType "Error"
        return $null
    }
}

function Get-EH-TeamDetails {
    param (
        [Parameter(Mandatory=$true)]
        [string]$EHOrganizationId,
        [Parameter(Mandatory=$true)]
        [string]$EHAuthorization
    )

    $UriString = "https://api.employmenthero.com/api/v1/organisations/$EHOrganizationId/teams"
    $TeamArguments = @{
        Uri             = $UriString
        Headers         = Get-EH-Headers -EHAuthorization $EHAuthorization
        Method          = 'GET'
        UseBasicParsing = $true
    }

    try {
        $Response = Invoke-RestMethod @TeamArguments
        if (-not $Response -or -not $Response.data.items) {
            Write-MessageLog "No team details returned." -LogType "Error"
            return $null
        }
        return $Response.data.items
    }
    catch {
        Write-MessageLog "Failed to fetch team details: $_" -LogType "Error"
        return $null
    }
}

function Get-EH-EmployeeCustomfields {
    param (
        [Parameter(Mandatory=$true)]
        [string]$EHOrganizationId,
        [Parameter(Mandatory=$true)]
        [string]$EHAuthorization,
        [Parameter(Mandatory=$true)]
        [string]$EmployeeId
    )

    $UriString = "https://api.employmenthero.com/api/v1/organisations/$EHOrganizationId/employees/$EmployeeId/custom_fields"
    $EmployeeArguments = @{
        Uri             = $UriString
        Headers         = Get-EH-Headers -EHAuthorization $EHAuthorization
        Method          = 'GET'
        UseBasicParsing = $true
    }

    try {
        $Response = Invoke-RestMethod @EmployeeArguments
        if (-not $Response -or -not $Response.data) {
            Write-MessageLog "No custom fields returned for Employee." -LogType "Error"
            return $null
        }
        return $Response.data
    }
    catch {
        Write-MessageLog "Failed to fetch Employee custom fields: $_" -LogType "Error"
        return $null
    }
}
