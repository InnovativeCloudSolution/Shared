$URLprefix = "api"
$Headers = @{
    'Content-Type' = 'application/json'
    'Accept' = 'application/json'
    'X-Cisco-Meraki-API-Key' = 'b93fe98f872f27dab09549ef5ffcb42d50b6d03b'
}

Clear-Host

# Program details - please remember to update the version number when making changes
Write-Host "`nMangano IT - Get Meraki Organisation Access" -ForegroundColor Yellow
Write-Host "Version: " -ForegroundColor yellow -NoNewLine; Write-Host "1.0"
Write-Host "Created by: " -ForegroundColor yellow -NoNewLine; Write-Host "Gabriel Nugent`n"
#Write-Host "Maintained by: " -ForegroundColor yellow -NoNewLine; Write-Host "Gabriel Nugent"

# User details
$Name = $(Write-Host "Your name: " -ForegroundColor yellow -NoNewLine; Read-Host)
$Email = $(Write-Host "Your email address (non-admin): " -ForegroundColor yellow -NoNewLine; Read-Host)

# Line break
Write-Host

# Get list of organisations and display them
$Orgs = Invoke-RestMethod -Uri "https://$URLprefix.meraki.com/api/v1/organizations" -Method Get -Headers $Headers
$Count = 1
foreach ($Name in $Orgs.name) {
    Write-Host "$Count. $Name"
    $Count++
}

# User chooses org they want access to
$Choice = $(Write-Host "`nSelect the organisation that you'd like access to: " `
-ForegroundColor yellow -NoNewLine; Read-Host)

# Fixes choice to match place in array, and gets related variables
$ArraySpot = [int]$Choice - 1
$Org = $Orgs.id[$ArraySpot]
$Body = @{
    "name" = $Name
    "email" = $Email
    "orgAccess" = "full"
}

# Makes admin account
$StatusCode = Invoke-RestMethod -Uri "https://$URLprefix.meraki.com/api/v1/organizations/$Org/admins" -Headers $Headers -Body $Body -Method Post

if ($StatusCode.status -eq 201) {
    Write-Host "User created successfully."
}

else {
    Write-Error "User not created successfully. Error code $StatusCode"
}