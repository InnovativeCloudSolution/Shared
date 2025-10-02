<#

Mangano IT - Generate New Partner Center Refresh Tokens
Created by: Gabriel Nugent
Version: 1.0

#>

## CONNECT TO AZURE KEY VAULT ##

$AzKeyVaultName = Get-AutomationVariable -Name 'AzKeyVaultName'
$AzKeyVaultConnectionName = 'AzureRunAsConnection'
try {
    # Get the connection properties
    $ServicePrincipalConnection = Get-AutomationConnection -Name $AzKeyVaultConnectionName
    $null = Connect-AzAccount `
        -ServicePrincipal `
        -TenantId $ServicePrincipalConnection.TenantId `
        -ApplicationId $ServicePrincipalConnection.ApplicationId `
        -CertificateThumbprint $ServicePrincipalConnection.CertificateThumbprint 
} catch {
    if (!$ServicePrincipalConnection) {
        # Azure Run As Account has not been enabled
        $ErrorMessage = "Connection $AzKeyVaultConnectionName not found."
        throw $ErrorMessage
    } else {
        # Something else went wrong
        Write-Error -Message $_.Exception.Message
        throw $_.Exception
    }
}

## SCRIPT VARIABLES ##

$ApplicationId = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-PartnerAppId' -AsPlainText
$ApplicationSecret = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-PartnerAppSecret' -AsPlainText
$ApplicationSecretSecure = ConvertTo-SecureString -String $ApplicationSecret -AsPlainText
$PartnerCenterTenantId = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-PartnerAppClientID' -AsPlainText
$PartnerCenterCredential = [PSCredential]::new($ApplicationId, $ApplicationSecretSecure)

# Build connection arguments
$TokenArguments = @{
    ApplicationId        = $ApplicationId
    Credential           = $PartnerCenterCredential
    Scopes               = "https://api.partnercenter.microsoft.com/user_impersonation"
    ServicePrincipal     = $true
    TenantId             = $PartnerCenterTenantId
    UseAuthorizationCode = $true
}

$NewToken = New-PartnerAccessToken @$TokenArguments