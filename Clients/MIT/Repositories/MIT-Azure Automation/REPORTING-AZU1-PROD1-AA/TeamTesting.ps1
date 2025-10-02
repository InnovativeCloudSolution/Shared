
$PartnerCenterApp = Get-AutomationPSCredential -Name 'PartnerCenterApp'
$ClientAppId = $PartnerCenterApp.UserName
$ClientSecret = $PartnerCenterApp.Password
$refreshToken = Get-AutomationVariable -Name 'PartnerCenterAppRefreshToken'
$Credential = New-Object System.Management.Automation.PSCredential ($ClientAppId, $ClientSecret)
$aadGraphToken = New-PartnerAccessToken -ApplicationId $ClientAppId -Credential $Credential -RefreshToken $refreshToken -Scopes 'https://graph.windows.net/.default' -ServicePrincipal -Tenant $MITSTenant

Connect-MicrosoftTeams -AadAccessToken $aadGraphToken.AccessToken -AccountId 'workflows@manganoit.com.au'