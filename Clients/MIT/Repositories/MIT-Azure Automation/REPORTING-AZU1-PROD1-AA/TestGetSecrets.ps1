# Log in to Azure with AZ (standard code)
########################################################
  
# Name of the Azure Run As connection
$ConnectionName = 'AzureRunAsConnection'
try {
    # Get the connection properties
    $ServicePrincipalConnection = Get-AutomationConnection -Name $ConnectionName      
   
    'Log in to Azure...'
    $null = Connect-AzAccount `
        -ServicePrincipal `
        -TenantId $ServicePrincipalConnection.TenantId `
        -ApplicationId $ServicePrincipalConnection.ApplicationId `
        -CertificateThumbprint $ServicePrincipalConnection.CertificateThumbprint 
} catch {
    if (!$ServicePrincipalConnection)     {
        # You forgot to turn on 'Create Azure Run As account' 
        $ErrorMessage = "Connection $ConnectionName not found."
        throw $ErrorMessage
    } else {
        # Something else went wrong
        Write-Error -Message $_.Exception.Message
        throw $_.Exception
    }
}
########################################################

# Variables for retrieving the correct secret from the correct vault
$VaultName = "MIT-AZU1-PROD1-AKV1"
$SecretName = "CleanCoTenantID"
 
# Retrieve value from Key Vault
$MySecret = Get-AzKeyVaultSecret -VaultName $VaultName -Name $SecretName -AsPlainText