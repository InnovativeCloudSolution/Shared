Import-Module MSOnline
Import-Module PartnerCenter

# Define date range
$DateNewest = (Get-Date).AddDays(-89)
$DateOldest = (Get-Date).AddDays(-90)

# Grab required access tokens
$SecretString = Get-AutomationVariable -Name ClientAppSecret
$SecretValue = ConvertTo-SecureString $SecretString -AsPlainText -Force
$ClientAppId = Get-AutomationVariable -Name ClientAppId
$TenantID = Get-AutomationVariable -Name TenantId
$SeasonsTenant = Get-AutomationVariable -Name SEA-TenantId
$Credential = New-Object System.Management.Automation.PSCredential ($ClientAppId, $SecretValue)
$RefreshToken = Get-AutomationVariable -Name PartnerCenterAppRefreshToken

# Connect to Partner Center
Connect-PartnerCenter -ApplicationId $ClientAppId -RefreshToken $RefreshToken -Credential $Credential | Out-Null

# Make Partner Center tokens as required
$aadGraphToken = New-PartnerAccessToken -ApplicationId $ClientAppId -Credential $Credential -RefreshToken $RefreshToken `
-Scopes 'https://graph.windows.net/.default' -ServicePrincipal -Tenant $TenantID
$graphToken = New-PartnerAccessToken -ApplicationId $ClientAppId -Credential $Credential -RefreshToken $RefreshToken `
-Scopes 'https://graph.microsoft.com/.default' -ServicePrincipal -Tenant $TenantID

# Connect to MSOL
Connect-MsolService -AdGraphAccessToken $aadGraphToken.AccessToken -MsGraphAccessToken $graphToken.AccessToken

# Fetch array of users who last set their passwords in the given range
$Users = Get-MsolUser -EnabledFilter EnabledOnly -MaxResults 9999 -TenantId $SeasonsTenant | Where-Object {$_.LastPasswordChangeTimeStamp -lt $DateNewest -and `
$_.LastPasswordChangeTimeStamp -gt $DateOldest -and $_.PasswordNeverExpires -eq $false} | `
Select UserPrincipalName, DisplayName, FirstName

# Old version for reference

#$Users = Get-ADUser -filter {pwdLastSet -lt $DateNewest -and pwdLastSet -gt $DateOldest -and passwordNeverExpires -eq $false -and Enabled -eq $true} `
#-SearchBase "OU=Users,OU=CleanCoQld,DC=internal,DC=cleancoqld,DC=com,DC=au" -Properties mail,PasswordLastSet,displayName,givenName,userPrincipalName | `
#Select-Object mail,PasswordLastSet,displayName,givenName,userPrincipalName

# Check if list is empty
if ($null -eq $Users) { Write-Output "No users in list" }

# Output list as parseable JSON
else { Write-Output $Users | ConvertTo-Json }