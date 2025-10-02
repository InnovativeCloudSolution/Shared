#Partner Center
$PartnerCenterApp = Get-AutomationPSCredential -Name 'PartnerCenterApp'

$PartnerCenterAppRefreshToken = Get-AutomationVariable -Name 'PartnerCenterAppRefreshToken'
$PartnerCenterAppAccessToken = New-PartnerAccessToken -Credential $PartnerCenterApp -ApplicationId '7adad76f-0bc2-45d5-8cee-d82429dbd377' -Scopes 'https://api.partnercenter.microsoft.com/user_impersonation' -RefreshToken $PartnerCenterAppRefreshToken

try{
    $Connection = Connect-PartnerCenter -ApplicationId '7adad76f-0bc2-45d5-8cee-d82429dbd377' -RefreshToken $PartnerCenterAppAccessToken.RefreshToken -Credential $PartnerCenterApp
    if($Connection.Account.Tenant -eq '5792a6c1-f4fe-466b-b97c-10eaf4fb3122'){
        $PartnerCenter = "Success"
        Set-AutomationVariable -Name 'PartnerCenterAppRefreshToken' -Value $PartnerCenterAppAccessToken.RefreshToken
    }else{
        $PartnerCenter = "Failure"
    }
}catch{
    $PartnerCenter = "Failure"
    $ErrorLogs += $_.Exception.Message
}

Write-Output $PartnerCenter

#Exchange
$ExchangeAppRefreshToken = Get-AutomationVariable -Name 'ExchangeAppRefreshToken'
$upn = "workflows@manganoit.com.au"
$ExchangeAppAccessToken = New-PartnerAccessToken -RefreshToken $ExchangeAppRefreshToken -Scopes 'https://outlook.office365.com/.default' -Tenant '5792a6c1-f4fe-466b-b97c-10eaf4fb3122' -ApplicationId 'a0c73c16-a7e3-4564-9a95-2bdf47383716'

#try{
#    $Connection = Connect-PartnerCenter -ApplicationId '7adad76f-0bc2-45d5-8cee-d82429dbd377' -RefreshToken $PartnerCenterAppAccessToken.RefreshToken -Credential $PartnerCenterApp
#    if($Connection.Account.Tenant -eq '5792a6c1-f4fe-466b-b97c-10eaf4fb3122'){
#        $ExchangeOnline = "Success"
        Set-AutomationVariable -Name 'ExchangeAppRefreshToken' -Value $ExchangeAppAccessToken.RefreshToken
#    }else{
#        $ExchangeOnline = "Failure"
#    }
#}catch{
#    $ExchangeOnline = "Failure"
#    $ErrorLogs += $_.Exception.Message
#}

#Write-Output $ExchangeOnline


# Email Result
$sendGridApiKey = Get-AutomationVariable -Name 'SendGrid Azure Automation API'

$Subject = "Partner Center: "+$PartnerCenter
$Body = "Refresh token for Partner Centre token has been refreshed: "+$PartnerCenter+" - "+$ErrorLogs
$SendGridEmail = @{
	From = 'workflows@manganoit.com.au'
	To = 'InternalAdmins@manganoit.com.au'
	Subject = $Subject
	Body = $Body
	
	SmtpServer = 'smtp.sendgrid.net'
	Port = 587
	UseSSL = $true
	Credential = New-Object PSCredential 'apikey', (ConvertTo-SecureString $sendGridApiKey -AsPlainText -Force)	
}
Send-MailMessage @SendGridEmail