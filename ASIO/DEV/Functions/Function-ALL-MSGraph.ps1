function Get-MSGraph-Secrets {
    param (
        [string]$AzKeyVaultName,
        [string[]]$SecretNames
    )

    Connect-AzAccount -Identity | Out-Null

    $secrets = @{}
    foreach ($secretName in $SecretNames) {
        try {
            $secretValue = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name $secretName -AsPlainText
            $secrets[$secretName] = $secretValue
        }
        catch {
            $secrets[$secretName] = $null
        }
    }
    return $secrets
}

function Get-MSGraph-BearerToken {
    param (
        [string]$AzKeyVaultName,
        [string]$TenantUrl,
        [string]$ClientIdSecretName,
        [string]$ClientSecretSecretName
    )

    $Secrets = Get-MSGraph-Secrets -AzKeyVaultName $AzKeyVaultName -SecretNames @($ClientIdSecretName, $ClientSecretSecretName)

    $ClientId = $Secrets[$ClientIdSecretName]
    $ClientSecret = $Secrets[$ClientSecretSecretName]
    $Scope = "https://graph.microsoft.com/.default"
    $ContentType = 'application/x-www-form-urlencoded'
    $GrantType = "client_credentials"

    $ApiArguments = @{
        Uri             = "https://login.microsoftonline.com/$TenantUrl/oauth2/v2.0/token"
        Method          = 'POST'
        Headers         = @{ 'Content-Type' = $ContentType }
        Body            = "client_id=$ClientId&client_secret=$ClientSecret&scope=$Scope&grant_type=$GrantType"
        UseBasicParsing = $true
    }

    try {
        $BearerToken = (Invoke-WebRequest @ApiArguments | ConvertFrom-Json).access_token -as [string]
        if (-not $BearerToken) {
            throw "Bearer token missing"
        }
        return $BearerToken
    } catch {
        throw $_
    }
}

function ConnectMSGraph {
    param (
        [string]$AzKeyVaultName,
        [string]$TenantUrl,
        [string]$ClientIdSecretName,
        [string]$ClientSecretSecretName
    )

    try {
        $BearerToken = Get-MSGraph-BearerToken -AzKeyVaultName $AzKeyVaultName -TenantUrl $TenantUrl -ClientIdSecretName $ClientIdSecretName -ClientSecretSecretName $ClientSecretSecretName | ConvertTo-SecureString -AsPlainText -Force
        Connect-MgGraph -AccessToken $BearerToken -NoWelcome
    }
    catch {
        throw $_
    }
}

function Get-MSGraph-BearerTokenTest {
    param (
        [string]$TenantUrl
    )
    
    $ClientId = "812408b3-6b87-418b-95ed-b036e2a67402"
    $ClientSecret = "Pmk8Q~FOA3Kd5QOR~9Bm.xiRBWNyJh1.TxjEcbIV"
    $Scope = "https://graph.microsoft.com/.default"
    $ContentType = 'application/x-www-form-urlencoded'
    $GrantType = "client_credentials"

    $ApiArguments = @{
        Uri             = "https://login.microsoftonline.com/$TenantUrl/oauth2/v2.0/token"
        Method          = 'POST'
        Headers         = @{ 'Content-Type' = $ContentType }
        Body            = "client_id=$ClientId&client_secret=$ClientSecret&scope=$Scope&grant_type=$GrantType"
        UseBasicParsing = $true
    }

    try {
        $BearerToken = (Invoke-WebRequest @ApiArguments | ConvertFrom-Json).access_token -as [string]
        if (-not $BearerToken) {
            throw "Bearer token missing"
        }
        return $BearerToken
    } catch {
        throw $_
    }
}

function ConnectMSGraphTest {
    param (
        [string]$TenantUrl
    )

    try {
        $BearerToken = Get-MSGraph-BearerTokenTest -TenantUrl $TenantUrl | ConvertTo-SecureString -AsPlainText -Force
        Connect-MgGraph -AccessToken $BearerToken -NoWelcome
    }
    catch {
        throw $_
    }
}

function Get-MSGraph-BearerTokenClients {
    param (
        [string]$TenantUrl
    )
    
    $ClientId = "e05b08c8-78aa-4bc2-9d3c-f99bca9bf5f6"
    $ClientSecret = "6a._9At5Yh3~-Bg1-j1iaHqKRR_6TsiEOm"
    $Scope = "https://graph.microsoft.com/.default"
    $ContentType = 'application/x-www-form-urlencoded'
    $GrantType = "client_credentials"

    $ApiArguments = @{
        Uri             = "https://login.microsoftonline.com/$TenantUrl/oauth2/v2.0/token"
        Method          = 'POST'
        Headers         = @{ 'Content-Type' = $ContentType }
        Body            = "client_id=$ClientId&client_secret=$ClientSecret&scope=$Scope&grant_type=$GrantType"
        UseBasicParsing = $true
    }

    try {
        $BearerToken = (Invoke-WebRequest @ApiArguments | ConvertFrom-Json).access_token -as [string]
        if (-not $BearerToken) {
            throw "Bearer token missing"
        }
        return $BearerToken
    } catch {
        throw $_
    }
}

function ConnectMSGraphClients {
    param (
        [string]$TenantUrl
    )

    try {
        $BearerToken = Get-MSGraph-BearerTokenClients -TenantUrl $TenantUrl | ConvertTo-SecureString -AsPlainText -Force
        Connect-MgGraph -AccessToken $BearerToken -NoWelcome
    }
    catch {
        throw $_
    }
}
