param(
    [string]$FirstName = '',
    [string]$LastName = '',
    [string]$MobilePhone = '',
    [string]$Title = '',
    [switch]$AdminRequired
)

$AzureADScript = '.\Dependencies\MIT-NewUser-Azure.ps1'
$CWMScript = '.\Dependencies\MIT-NewUser-CWM.ps1'
$ExchangeScript = '.\Dependencies\MIT-NewUser-Exchange.ps1'
$TeamsScript = '.\Dependencies\MIT-NewUser-Teams.ps1'

## Program details - please remember to update the version number when making changes
Write-Host "`nMangano IT - Internal New User" -ForegroundColor Yellow
Write-Host "Version: " -ForegroundColor yellow -NoNewLine; Write-Host "2.0"
Write-Host "Created by: " -ForegroundColor yellow -NoNewLine; Write-Host "Gabriel Nugent"

# Confirms the phone number starts with a 0
if ($MobilePhone[0] -eq "4") { $MobilePhone="0"+$MobilePhone }

$Log="The parameters that have been provided are:`r`n"
$Log+="First Name - $FirstName`r`n"
$Log+="Last Name - $LastName`r`n"
$Log+="Title - $Title`r`n"
$Log+="Mobile Phone - $MobilePhone`r`n"

Write-Host "`nThe parameters that have been provided are:" -ForegroundColor yellow
Write-Host "First Name - $FirstName"
Write-Host "Last Name - $LastName"
Write-Host "Title - $Title"
Write-Host "Mobile Phone - $MobilePhone"

#Company name in CWM
$CompanyName = 'Mangano IT'
#Email domain (for CWM and users)
$Domain = '@manganoit.com.au'

$TenantID = "5792a6c1-f4fe-466b-b97c-10eaf4fb3122" #The ManganoIT Tenancy

# Grab email address of person running script
$ScriptAdmin = $(Write-Host "`nAdmin login: " -ForegroundColor yellow -NoNewLine; Read-Host)
if ($ScriptAdmin -eq "gabe") { $ScriptAdmin = "adm.gabriel.nugent@manganoit.com.au" }

# Generate a password, convert it to a secure string
$Log+="`r`nGenerating password through DinoPass`r`n"

# Loops until the password is at least 12 characters long
while ($Password.Length -lt 12) { $Password = Invoke-restmethod -uri "https://www.dinopass.com/password/strong" }

# Create the user's details
$Username = $FirstName+'.'+$LastName
$Email = $Username+$Domain
$DisplayName = $FirstName+' '+ $LastName
$Log+="`r`nConstructing the username and email for the user`r`n"
$Log+="Username: $Username`r`n"
$Log+="Email: $Email`r`n"

# Give a list of teams for the new user to be in - asks for input
$Team = ''

Write-Host "`n1. Service Team Level 1 (including CSCs)"
Write-Host "2. Service Team Level 2"
Write-Host "3. Service Team Level 3"
Write-Host "4. Projects Team"
Write-Host "5. Sales Team (including Admin Assistants)"
Write-Host "6. Leadership Team"
$TeamSelect = $(Write-Host "Select the new employee's team (by number): " -ForegroundColor yellow -NoNewLine; Read-Host)
switch ($TeamSelect) {
    "1" { $Team = "Service Team Level 1" }
    "2" { $Team = "Service Team Level 2" }
    "3" { $Team = "Service Team Level 3" }
    "4" { $Team = "Projects Team" }
    "5" { $Team = "Sales Team" }
    "6" { $Team = "Leadership Team" }
}

# Create the admin user's details
if ($AdminRequired) {
    #Admin Firstname Surname (No firstname/surname)
    Write-Host "`nConstructing the username and email for the user`n"
    $AdminUsername = "adm."+$FirstName+'.'+ $LastName
    $AdminEmail = $AdminUsername+$Domain
    $AdminDisplayName = "Admin "+$FirstName+' '+ $LastName
    $Log+="`r`nConstructing the username and email for the user`r`n"
    $Log+="Username: $Username`r`n"
    $Log+="Email: $Email`r`n"
    Write-Host "Username: $Username`n"
    Write-Host "Email: $Email`n"

    # Build parameters to be used by Azure AD script
    $AzureADParameters = ' -FirstName '+$FirstName+' -LastName '+$LastName+' -DisplayName '+$DisplayName+' -Email '+$Email`
    +' -MobilePhone '+$MobilePhone+' -Title '+$Title+' -Team '+$Team+' -Password '+$Password+' -ScriptAdmin '+$ScriptAdmin+`
    ' -TenantId '+$TenantID+' -AdminRequired True -AdminEmail '+$AdminEmail+' -AdminDisplayName '+$AdminDisplayName
}

# Build parameters to be used by Azure AD script
else { 
    $AzureADParameters = ' -FirstName '+$FirstName+' -LastName '+$LastName+' -DisplayName '+$DisplayName+' -Email '+$Email`
    +' -MobilePhone '+$MobilePhone+' -Title '+$Title+' -Team '+$Team+' -Password '+$Password+' -ScriptAdmin '+$ScriptAdmin+`
    ' -TenantId '+$TenantID
 }

# Append parameters to script variable, then run script
$AzureADScript += $AzureADParameters
$AzureADLog = Invoke-Expression -Command $AzureADScript

# Check if the new user script failed
if ($AzureADLog -contains 'ERROR: '+$Email+' does not exist.' -or $AzureADScript -contains "ERROR: Unable to connect to Azure AD.") {
    Write-Error 'New user has not been created, exiting script...'
    exit
}

# Appends log from MIT-NewUser-Azure.ps1 to the main log, but only if it was successful
$Log += $AzureADLog

Write-Host "Thanks! The account has been created! Please wait a few minutes while a license is added to the account!"

# Wait 300 seconds for licenses (that are applied by security group) to finish applying
for ($i=0; $i -le 300; $i++) {
    $Percent = [math]::Round($i/300*100)
    Write-Progress -Activity "Waiting for license addition" -Status "$Percent% Complete:" -PercentComplete $Percent;
    Start-Sleep -Seconds 1
}

# Build Exchange script parameters
$ExchangeParameters = ' -Email '+$Email+' -Team '+$Team+' -ScriptAdmin '+$ScriptAdmin

# Append parameters to script variable, then run script
$ExchangeScript += $ExchangeParameters
$ExchangeLog = Invoke-Expression $ExchangeScript

# Appends log from MIT-NewUser-Exchange.ps1 to the main log
$Log += $ExchangeLog

# Build Teams script parameters
$TeamsParameters = ' -Email '+$Email+' -Team '+$Team+' -TenantId '+$TenantID+' -ScriptAdmin '+$ScriptAdmin

# Append parameters to script variable, then run script
$TeamsScript += $TeamsParameters
$TeamsLog = Invoke-Expression $TeamsScript

# Appends log from MIT-NewUser-Teams.ps1 to the main log
$Log += $TeamsLog

Write-Host 'The user account should now be finished. Please review the JSON log below for any errors.' -ForegroundColor Yellow

$json = @"
{
"Email": "$Email",`n
"Username": "$Username",`n
"Password": "$Password",`n
"AdminEmail":$AdminEmail,`n
"Log":"$Log"
}
"@

Write-Output $json
Out-File -FilePath '.\'+$DisplayName+'.txt' -InputObject $json