$env:CW_URL = "https://your-instance.connectwisemanagedservices.com"
$env:CW_COMPANY = "YourCompanyID"
$env:CW_PUBLIC_KEY = "YourPublicKey"
$env:CW_PRIVATE_KEY = "YourPrivateKey"

$authString = "$env:CW_COMPANY+$env:CW_PUBLIC_KEY:$env:CW_PRIVATE_KEY"
$env:CW_AUTH = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($authString))

Write-Host "ConnectWise API configuration loaded"
Write-Host "URL: $env:CW_URL"
Write-Host "Company: $env:CW_COMPANY"

