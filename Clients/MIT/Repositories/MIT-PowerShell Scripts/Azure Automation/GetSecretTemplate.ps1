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
    if (!$ServicePrincipalConnection)     {
        # You forgot to turn on 'Create Azure Run As account' 
        $ErrorMessage = "Connection $AzKeyVaultConnectionName not found."
        throw $ErrorMessage
    } else {
        # Something else went wrong
        Write-Error -Message $_.Exception.Message
        throw $_.Exception
    }
}
 
# Retrieve value from Key Vault
$Secret = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name $SecretName -AsPlainText