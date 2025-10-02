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