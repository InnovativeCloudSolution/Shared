function Get-EH-Headers {
    param (
        [Parameter(Mandatory=$true)]
        $EHAuthorization
    )
    return @{
        'Authorization' = "Bearer $EHAuthorization"
    }
}

function Get-EH-Authorization {
    param (
        [string]$AzKeyVaultName,
        [string]$ClientIdSecretName,
        [string]$ClientSecretSecretName,
        [string]$CodeSecretName,
        [string]$RefreshTokenSecretName,
        [string]$RedirectUriSecretName
    )

    Write-MessageLog -Message "Fetching Employment Hero Authorization Token using secrets from Azure Key Vault."

    # Fetch secrets from Azure Key Vault
    $SecretNames = @($ClientIdSecretName, $ClientSecretSecretName, $CodeSecretName, $RefreshTokenSecretName, $RedirectUriSecretName)
    $Secrets = Get-MSGraph-Secrets -AzKeyVaultName $AzKeyVaultName -SecretNames $SecretNames

    # Map secrets to variables
    $ClientId = $Secrets[$ClientIdSecretName]
    $ClientSecret = $Secrets[$ClientSecretSecretName]
    $Code = $Secrets[$CodeSecretName]
    $RefreshToken = $Secrets[$RefreshTokenSecretName]
    $RedirectUri = $Secrets[$RedirectUriSecretName]

    # Construct the API URI for the OAuth token request
    $ApiUri = "https://oauth.employmenthero.com/oauth2/token"

    # Prepare the request body
    $Body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        grant_type    = "refresh_token"
        code          = $Code
        redirect_uri  = $RedirectUri
        refresh_token = $RefreshToken
    }

    $ApiArguments = @{
        Uri             = $ApiUri
        Method          = 'POST'
        Headers         = @{ 'Content-Type' = 'application/x-www-form-urlencoded' }
        Body            = $Body
        UseBasicParsing = $true
    }

    try {
        $Response = Invoke-WebRequest @ApiArguments | ConvertFrom-Json

        if (-not $Response.access_token) {
            Write-ErrorLog -Message "Authorization token is null or missing in the response."
            throw "Authorization token missing"
        }

        return $Response.access_token
    }
    catch {
        Write-ErrorLog -Message "Failed to retrieve Authorization Token: $_"
        throw $_
    }
}

function Get-EH-AuthorizationTest {

    # Map secrets to variables
    $ClientId = "_C_6tY11_apcWCims66eaesh0A9ZxY5hCN1EUXnsiPg"
    $ClientSecret = "TKibvtBB6uZ86Nxo9ODPuEM4ZGX2Q4aFTW3nv_LT42Y"
    $Code = "JTnk3VwY51lUDCb2G5XutpVxnRWkQLT9EMV-6eyuzCY"
    $RefreshToken = "CucABqgAj0NkTzm0IFrUFnFeF-YHZcI5NQraJZYioRk"
    $RedirectUri = "https://seasonsliving.com.au/"

    # Construct the API URI for the OAuth token request
    $ApiUri = "https://oauth.employmenthero.com/oauth2/token"

    # Prepare the request body
    $Body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        grant_type    = "refresh_token"
        code          = $Code
        redirect_uri  = $RedirectUri
        refresh_token = $RefreshToken
    }

    $ApiArguments = @{
        Uri             = $ApiUri
        Method          = 'POST'
        Headers         = @{ 'Content-Type' = 'application/x-www-form-urlencoded' }
        Body            = $Body
        UseBasicParsing = $true
    }

    try {
        $Response = Invoke-WebRequest @ApiArguments | ConvertFrom-Json

        if (-not $Response.access_token) {
            Write-ErrorLog -Message "Authorization token is null or missing in the response."
            throw "Authorization token missing"
        }

        return $Response.access_token
    }
    catch {
        Write-ErrorLog -Message "Failed to retrieve Authorization Token: $_"
        throw $_
    }
}

function Get-EH-EmployeeAll {
    param (
        $EHOrganizationId,
        $EHAuthorization
    )

    $UriBase = "https://api.employmenthero.com/api/v1/organisations/$EHOrganizationId/employees"
    $Headers = Get-EH-Headers -EHAuthorization $EHAuthorization

    $AllEmployees = @()
    $PageIndex = 1
    $TotalPages = 1

    do {
        $UriString = "$($UriBase)?page=$($PageIndex)"
        $AllEmployeeArguments = @{
            Uri             = $UriString
            Headers         = $Headers
            Method          = 'GET'
            UseBasicParsing = $true
        }

        try {
            $Response = Invoke-WebRequest @AllEmployeeArguments | ConvertFrom-Json

            if ($Response -and $Response.data.items) {
                $AllEmployees += $Response.data.items

                $TotalPages = $Response.data.total_pages
                $PageIndex++
            } else {
                Write-MessageLog "No employees returned on page $PageIndex." -LogType "Error"
                break
            }
        }
        catch {
            Write-MessageLog "Failed to fetch employees on page $($PageIndex): $_" -LogType "Error"
            break
        }
    } while ($PageIndex -le $TotalPages)

    # Return the aggregated list of all employees
    if ($AllEmployees.Count -eq 0) {
        Write-MessageLog "No employees found across all pages." -LogType "Error"
        return $null
    }

    return $AllEmployees
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

function Get-EH-Manager {
    param (
        [Parameter(Mandatory=$true)]
        [string]$EHOrganizationId,
        [Parameter(Mandatory=$true)]
        [string]$EHAuthorization,
        [Parameter(Mandatory=$true)]
        [string]$ManagerId
    )
    try {
        $ManagerDetails = Get-EH-Employee -EHOrganizationId $EHOrganizationId -EmployeeId $ManagerId -EHAuthorization $EHAuthorization
        return $ManagerDetails
    } catch {
        Write-Error -Message "Failed to fetch manager details. Error: $_"
        return $null
    }
}

function Get-EH-TeamDetails {
    param (
        [Parameter(Mandatory = $true)]
        [string]$EHOrganizationId,
        [Parameter(Mandatory = $true)]
        [string]$EHAuthorization
    )

    # Updated URL with item_per_page=50
    $UriString = "https://api.employmenthero.com/api/v1/organisations/$EHOrganizationId/teams?item_per_page=100"
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

function Get-EH-TeamMembers {
    param (
        [Parameter(Mandatory=$true)]
        $EHOrganizationId,
        [Parameter(Mandatory=$true)]
        $EHAuthorization,
        [Parameter(Mandatory=$true)]
        $TeamId
    )
    $UriString = "https://api.employmenthero.com/api/v1/organisations/$EHOrganizationId/teams/$TeamId/employees?item_per_page=100"
    $TeamMembersArguments = @{
        Uri             = $UriString
        Headers         = Get-EH-Headers -EHAuthorization $EHAuthorization
        Method          = 'GET'
        UseBasicParsing = $true
    }
    try {
        $TeamMembersResponse = Invoke-WebRequest @TeamMembersArguments | ConvertFrom-Json
        return $TeamMembersResponse.data.items
    }
    catch {
        Write-Error -Message "Failed to fetch team members for team $TeamId. Error: $_"
        return $null
    }
}

function Get-EH-TeamNames {
    param (
        [Parameter(Mandatory=$true)]
        $EHOrganizationId,
        [Parameter(Mandatory=$true)]
        $EHAuthorization,
        [Parameter(Mandatory=$true)]
        $EmployeeId,
        [Parameter(Mandatory=$true)]
        $Teams
    )
    try {
        $TeamMembersWithTeamName = @()
        foreach ($TeamItem in $Teams) {
            $TeamId = $TeamItem.id
            $TeamMembers = Get-EH-TeamMembers -EHOrganizationId $EHOrganizationId -TeamId $TeamId -EHAuthorization $EHAuthorization
            if ($null -eq $TeamMembers) { continue }

            foreach ($TeamMember in $TeamMembers) {
                if ($TeamMember.id -eq $EmployeeId) {
                    $TeamMembersWithTeamName += [PSCustomObject]@{
                        TeamMember = $TeamMember
                        TeamName   = $TeamItem.name
                    }
                }
            }
        }
        return $TeamMembersWithTeamName | Select-Object -ExpandProperty TeamName
    } catch {
        Write-Error -Message "Failed to fetch team names. Error: $_"
        return $null
    }
}

function Get-EH-EmployeeTeamNames {
    param (
        [Parameter(Mandatory=$true)]
        $EHOrganizationId,
        [Parameter(Mandatory=$true)]
        $EHAuthorization,
        [Parameter(Mandatory=$true)]
        $EmployeeId,
        [Parameter(Mandatory=$true)]
        $Teams
    )
    try {
        $TeamNames = Get-EH-TeamNames -EHOrganizationId $EHOrganizationId -EHAuthorization $EHAuthorization -EmployeeId $EmployeeId -Teams $Teams
        if ($null -eq $TeamNames) {
            Write-Error -Message "No team names found for employee."
            return $null
        }
        return $TeamNames
    } catch {
        Write-Error -Message "Failed to fetch employee team names. Error: $_"
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
