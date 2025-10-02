<#

Mangano IT - Partner Center - Refresh Access Token
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be run on a schedule via Power Automate.

Based on https://tminus365.com/my-automations-break-with-gdap-the-fix/

#>

param (
    [Parameter(Mandatory)][string]$ApplicationId,
    [Parameter(Mandatory)][string]$ApplicationSecret,
    [Parameter(Mandatory)][string]$RefreshToken,
    [Parameter(Mandatory)][string]$TenantId,
    [string]$AccessTokenSecretName,
    [string]$RefreshTokenSecretName
)

## SCRIPT VARIABLES ##

$SecureApplicationSecret = ConvertTo-SecureString -String $ApplicationSecret -AsPlainText -Force
$Credential = [PSCredential]::new($ApplicationId, $SecureApplicationSecret)

## GET PARTNER ACCESS TOKEN ##

$PartnerAccessParameters = @{
    ApplicationId = $ApplicationId
    Credential = $Credential
    RefreshToken = $RefreshToken
    Scopes = "https://api.partnercenter.microsoft.com/user_impersonation"
    ServicePrincipal = $true
    TenantId = $TenantId
}

$PartnerAccessToken = New-PartnerAccessToken @PartnerAccessParameters

## UPDATE KEY VAULT SECRET IF PROVIDED ##

if ($AccessTokenSecretName -ne '' -or $RefreshTokenSecretName -ne '') {
    $AzKeyVaultName = Get-AutomationVariable -Name 'AzKeyVaultName'

    # Connect to Azure using Managed Identity
    try {
        Connect-AzAccount -Identity | Out-Null
    } catch {
        Write-Error -Message $_.Exception.Message
        throw $_.Exception
    }

    # Update secrets
    $Expires = (Get-Date).AddDays(90).ToUniversalTime()

    if ($AccessTokenSecretName -ne '') {
        $AccessTokenSecret = ConvertTo-SecureString -String $PartnerAccessToken.AccessToken -AsPlainText -Force
        Set-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name $KeyVaultSecretName -SecretValue $AccessTokenSecret -Expires $Expires
    }
    
    if ($RefreshTokenSecretName -ne '') {
        $RefreshTokenSecret = ConvertTo-SecureString -String $PartnerAccessToken.RefreshToken -AsPlainText -Force
        Set-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name $KeyVaultSecretName -SecretValue $RefreshTokenSecret -Expires $Expires
    }
} else {
    Write-Output $PartnerAccessToken
}