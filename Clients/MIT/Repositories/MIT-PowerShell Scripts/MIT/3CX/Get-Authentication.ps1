# Requires PS full language mode/administrative prompt or dev machine.
# $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
# $headers.Add("Content-Type", "application/json")
# Created by Joshua Ceccato and Juan Moredo. 

# Run this script first to get authorised headers.
$password="put me in keyvault?" # https://mits.au.itglue.com/797747/passwords/3817176368334888

# Updated to use hash table instead of Dictionary Object.
$headers = @{
    "Content-Type" = "application/json"
}

$body = @"
{`"SecurityCode`":`"`",`"Password`":`"$password`",`"Username`":`"7000`"}
"@

$response = Invoke-RestMethod 'https://voip.manganoit.com.au:5001/webclient/api/Login/GetAccessToken' -Method 'POST' -Headers $headers -Body $body

$ACCESS_TOKEN = $response.token.access_token

# Not used at this time
# $REFRESH_TOKEN = $response.token.refresh_token

# $headers.Add("Authorization", "Bearer $ACCESS_TOKEN") # For dictionary object. 
$headers["Authorization"] = "Bearer $ACCESS_TOKEN" # For Hash table.