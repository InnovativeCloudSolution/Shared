# Grab required access tokens
$SecretString = Get-AutomationVariable -Name ClientAppSecret
$SecretValue = ConvertTo-SecureString $SecretString -AsPlainText -Force
$ClientAppId = Get-AutomationVariable -Name ClientAppId
$TenantId = Get-AutomationVariable -Name TenantId
$CQLTenant = Get-AutomationVariable -Name CQL-TenantId
$Credential = New-Object System.Management.Automation.PSCredential ($ClientAppId, $SecretValue)
$RefreshToken = Get-AutomationVariable -Name PartnerCenterAppRefreshToken

# Connect to Partner Center
Connect-PartnerCenter -ApplicationId $ClientAppId -RefreshToken $RefreshToken -Credential $Credential | Out-Null

# Make Partner Center tokens as required
$AadAccessToken = New-PartnerAccessToken -ApplicationId $ClientAppId -Credential $Credential -RefreshToken $RefreshToken `
-Scopes 'https://graph.windows.net/.default' -ServicePrincipal -Tenant $TenantId
$MsAccessToken = New-PartnerAccessToken -ApplicationId $ClientAppId -Credential $Credential -RefreshToken $RefreshToken `
-Scopes 'https://graph.microsoft.com/.default' -ServicePrincipal -Tenant $TenantId

# Connect to Azure AD
Connect-AzureAD -TenantId $CQLTenant -AadAccessToken $AadAccessToken -MsAccessToken $MsAccessToken