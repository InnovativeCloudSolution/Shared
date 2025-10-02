Import-Module MSOnline
Import-Module PartnerCenter

# Define date range
$DateNewest = (Get-Date).AddDays(-89)
$DateOldest = (Get-Date).AddDays(-90)

# Grab required access tokens
$SecretValue = ConvertTo-SecureString "zCzfaJSW1_AKqcfcJI5f6tkPMWJU3]=[" -AsPlainText -Force
$ClientAppId = "7adad76f-0bc2-45d5-8cee-d82429dbd377"
$SeasonsTenant = "746036af-6519-4420-ba95-65e211e271c6"
$MitTenant = "5792a6c1-f4fe-466b-b97c-10eaf4fb3122" #The ManganoIT Tenancy
$refreshToken = "0.AUEAwaaSV_70a0a5fBDq9PsxIm_X2nrCC9VFjO7YJCnb03dBAEA.AgABAAAAAAD--DLA3VO7QrddgJg7WevrAgDs_wQA9P_5AxmuV7SehPuVufN-wBT5xc-_NAUAVoCcp8bXTLnEU7f6AJT1u511JI_EHpxDvVMogMuZpuMGMY4ZfZyvjbM0AdaPSSHxoV4D3sJVuJCAVLs04kNnM1AXC-YQxEk6KURvKN4hK6gKeOj-jAlOfDUG3M-RIo5eEAQQO34ogWD9zu25ik4wyqBjManSr22sNvv271qocjnrLgtRihjaDbeGI1BAIn_7DUtNZoqROJ0cz7-Cm2Y4g38S9ugaWEPsTGUuNjv1QStM-PFdy__16sql3rQ-Uxa8LpIYRWfzRQroknLJDwXPF8BybheEuFR9r-aJpaLHKtaElyoTZUevxEWtyU-ryGKVktyCMgk5u54lhv2aUaM2jRd6lX-fG7KQhWO0JeafInuY921GUpwGxjaJjppjkhft6gctCLrZfI-foVEcCzyMjXVKX8-P5Bqy7TvIc8W1o1C536R7c5NHNc9tBBZOoK845VMTIaAPN9oaPmIdarsDqiOijTNDjfP-K84OHiSoUtc-amhKF0oVjRLToO-UV-hHLSipjPA4azGh6Yfe-Aq_yMlVLqxHyeR0g7EUmNPEHeTzkOt2CVHY6V_VBl9kp0AhEm-jATq8CBbZIsBAczdSR1i6gczdLG74Hwx1lCRSeTKlQWPcSTWOUO9NGDzKuONFmq2k0ts-nxEaN3R3IcJlg6h2RboQl4NBMepY99BFDuRdSSDHdTcwiUEkrpObLkDu2lalrjP23uSFcHWnyJzx9K5m9vyIjp5YV6uwkqr4hrsJ2FI2QZjH-G1I3ftkz6X5CEloRkMIYOQVYYSmXtrRILZWNVYgY_jvxL7unvYd55oUsaR9TTCvOIC3cdvbIGEQyVp67K-yQemM1ZXbUdH48WuTQBGixpgmoyiSJsAkXnajmpck0RSSSLDfUyt8oeq_yWXjbdQkcHccoxPlznXtbD1UjjGTFyrUlualn4IIF89Cqi8HcfZzB32pdZQ2seKnmLToFodhrzbi7IRzXHqASgkPtVbGNQACB1Q-Tt_c_13dXnHnjoGM9xguVNfHDHkkOLeiguJGTjHoTieDtx2kI3yXa5wszv7wlhLbJFXYji1epypdFv91keUdUlsSpw1Dvl7qrw3-5YGPNuivJDQv9on-7nU1Wvtgx_xwBs-9sjotcWvBwLfhTzGpm8Vnkg6YTfBzZEEDlIaP"
$Credential = New-Object System.Management.Automation.PSCredential ($ClientAppId, $SecretValue)

# Connect to Partner Center
Connect-PartnerCenter -ApplicationId $ClientAppId -RefreshToken $RefreshToken -Credential $Credential

# Make Partner Center tokens as required
$aadGraphToken = New-PartnerAccessToken -ApplicationId $ClientAppId -Credential $Credential -RefreshToken $RefreshToken `
-Scopes 'https://graph.windows.net/.default' -ServicePrincipal -Tenant $MitTenant
$graphToken = New-PartnerAccessToken -ApplicationId $ClientAppId -Credential $Credential -RefreshToken $RefreshToken `
-Scopes 'https://graph.microsoft.com/.default' -ServicePrincipal -Tenant $MitTenant

# Connect to MSOL
Connect-MsolService -AdGraphAccessToken $aadGraphToken.AccessToken -MsGraphAccessToken $graphToken.AccessToken

# Fetch array of users who last set their passwords in the given range
$Users = Get-MsolUser -EnabledFilter EnabledOnly -MaxResults 9999 -TenantId $SeasonsTenant | Where-Object {$_.LastPasswordChangeTimeStamp -lt $DateNewest -and `
$_.LastPasswordChangeTimeStamp -gt $DateOldest -and $_.PasswordNeverExpires -eq $false} | `
Select-Object UserPrincipalName, DisplayName, FirstName

# Old version for reference

#$Users = Get-ADUser -filter {pwdLastSet -lt $DateNewest -and pwdLastSet -gt $DateOldest -and passwordNeverExpires -eq $false -and Enabled -eq $true} `
#-SearchBase "OU=Users,OU=CleanCoQld,DC=internal,DC=cleancoqld,DC=com,DC=au" -Properties mail,PasswordLastSet,displayName,givenName,userPrincipalName | `
#Select-Object mail,PasswordLastSet,displayName,givenName,userPrincipalName

# Check if list is empty
if ($null -eq $Users) { Write-Output "No users in list" }

# Output list as parseable JSON
else { Write-Output $Users | ConvertTo-Json }