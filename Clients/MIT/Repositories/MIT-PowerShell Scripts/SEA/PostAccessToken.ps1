$Headers = @{
    'Content-Type' = 'application/x-www-form-urlencoded'
    'Host' = 'login.microsoftonline.com'
}

$Body = @{
    'client_id' = 'e05b08c8-78aa-4bc2-9d3c-f99bca9bf5f6'
    'client_secret' = '6a._9At5Yh3~-Bg1-j1iaHqKRR_6TsiEOm'
    'scope' = 'https://graph.microsoft.com/.default'
    'grant_type' = 'client_credentials'
}

$BearerToken = Invoke-RestMethod -Method Get -Uri 'https://login.microsoftonline.com/seasonsagedcare.com.au/oauth2/v2.0/token' `
-Headers $Headers -Body $Body

Write-Host $BearerToken