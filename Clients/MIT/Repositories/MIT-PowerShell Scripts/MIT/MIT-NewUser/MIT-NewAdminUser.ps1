<#

Mangano IT - New Admin User Script (Module-Based Version)
Created by: Gabriel Nugent
Version: 1.0

Please make sure you have the included module at this location:
C:\Modules\MIT-NewUserModule\MIT-NewUserModule.ps1

The script will not work without its module!

#>

param(
    [Parameter(Mandatory=$true)][string]$FirstName = '',
    [Parameter(Mandatory=$true)][string]$LastName = '',
    [Parameter(Mandatory=$true)][string]$MainAccount = '',
    [string]$MobilePhone = '',
    [string]$Password = '',
    [string]$Team = ''
)

function Test-ModuleInstalled {
    param (
        [string]$ModuleName
    )
    $module = Get-Module -ListAvailable -Name $ModuleName
    return $module -ne $null
}

## MODULES ##
if (Test-ModuleInstalled -ModuleName "AzureAD") {
    Import-Module AzureAD
    Write-Host "AzureAD module imported successfully."
} elseif (Test-ModuleInstalled -ModuleName "AzureADPreview") {
    Import-Module AzureADPreview
    Write-Host "AzureADPreview module imported successfully."
} else {
    Write-Host "Neither AzureAD nor AzureADPreview module is installed."
}

Import-Module -Name 'C:\Scripts\Repositories\PowerShell Scripts\MIT\MIT-NewUser\MIT-NewUserModule\MIT-NewUserModule' -Force

## VARIABLES ##
$TenantID = "5792a6c1-f4fe-466b-b97c-10eaf4fb3122" #The ManganoIT Tenancy

## DEFINE STANDARD VARIABLES FOR ALL ACCOUNTS ##
# Define groups that every user is added to
$Standard_AD_Admin = @(
    'AdminAgents'
    'SG.License.EMSE5'
    'SG.License.M365E3.NoMail'
    'SG.Role.SSO.ITGlue'
)

$SD_AD_Admin = @(

)

$SDL1_AD_Admin = @(
    'SG.Role.LogicMonitor.Reader'
    'SG.Role.Admin.L1'
)

$SDL2_AD_Admin = @(
    'SG.Role.LogicMonitor.Contributor'
    'SG.Role.Admin.L2'
    'SG.Role.LocalAdmin'
)

$SDL3_AD_Admin = @(
    'SG.Role.LogicMonitor.Admin'
    'SG.Role.Admin.L3'
    'SG.Role.LocalAdmin'
)

$Sales_AD_Admin = @(
    'SG.Role.LogicMonitor.Reader'
)

$Projects_AD_Admin = @(
    'SG.Role.LogicMonitor.Admin'
    'SG.Role.Admin.L3'
    'SG.Role.LocalAdmin'
)

$Leadership_AD_Admin = @(
    'SG.Role.LogicMonitor.Contributor'
)

## Program details - please remember to update the version number when making changes
Write-Host "`nMangano IT - Internal New Admin Account" -ForegroundColor Yellow
Write-Host "Version: " -ForegroundColor yellow -NoNewLine; Write-Host "1.0"
Write-Host "Created by: " -ForegroundColor yellow -NoNewLine; Write-Host "Gabriel Nugent"
Write-Host "`nPlease remember to copy the folder " -NoNewLine; Write-Host "MIT-NewUserModule " -ForegroundColor yellow -NoNewline
Write-Host "to " -NoNewLine; Write-Host "C:\Scripts\Repositories\PowerShell Scripts\MIT\MIT-NewUser\MIT-NewUserModule\" -ForegroundColor yellow

# Confirms the phone number starts with a 0
if ($MobilePhone[0] -eq "4") { $MobilePhone="0"+$MobilePhone }

$Log="The parameters that have been provided are:`n"
$Log+="First Name - $FirstName`n"
$Log+="Last Name - $LastName`n"
$Log+="Main Account - $MainAccount`n"
$Log+="Mobile Phone - $MobilePhone`n"

Write-Host "`nThe parameters that have been provided are:" -ForegroundColor yellow
Write-Host "First Name - $FirstName"
Write-Host "Last Name - $LastName"
Write-Host "Main Account - $MainAccount"
Write-Host "Mobile Phone - $MobilePhone"

# Grab email address of person running script
$ScriptAdmin = $(Write-Host "`nAdmin login: " -ForegroundColor yellow -NoNewLine; Read-Host)
if ($ScriptAdmin -eq "gabe") { $ScriptAdmin = "adm.gabriel.nugent@manganoit.com.au" }

# Generates a password if one isn't already provided
if ($Password -eq '') {
    $Log+="`nGenerating password through DinoPass`n"
    # Loops until the password is at least 12 characters long
    while ($Password.Length -lt 12) { $Password = Invoke-restmethod -uri "https://www.dinopass.com/password/strong" }
}

# Create the admin user's details
#Admin Firstname Surname (No firstname/surname)
Write-Host "`nConstructing the username and email for the user`n"
$AdminUsername = "adm."+$FirstName+'.'+ $LastName
$AdminEmail = $AdminUsername+"@manganoit.com.au"
$AdminDisplayName = "Admin "+$FirstName+' '+ $LastName

# Give a list of teams for the new user to be in - asks for input
if ($Team -eq '') {
    Write-Host "`n1. Service Team Level 1 (including CSCs and trainees)"
    Write-Host "2. Service Team Level 2"
    Write-Host "3. Service Team Level 3"
    Write-Host "4. Projects Team"
    Write-Host "5. Sales and Admin Team"
    Write-Host "6. Leadership Team"
    $TeamSelect = $(Write-Host "`nSelect the new employee's team (by number): " -ForegroundColor yellow -NoNewLine; Read-Host)
    switch ($TeamSelect) {
        "1" { $Team = "Service Team Level 1" }
        "2" { $Team = "Service Team Level 2" }
        "3" { $Team = "Service Team Level 3" }
        "4" { $Team = "Projects Team" }
        "5" { $Team = "Sales Team" }
        "6" { $Team = "Leadership Team" }
    }
}

# Attempt to connect to Azure AD
Write-Host "`nConnecting to Azure AD..."
try { Connect-AzureAD -TenantId $TenantID -AccountId $ScriptAdmin }
catch {
    Write-Error "ERROR: Unable to connect to Azure AD.`n" 
    exit
}

# Creates admin account
Write-Host "Creating the admin account..."
$Log += Add-UserAccountAAD -DisplayName $AdminDisplayName -Email $AdminEmail -UsageLocation 'AU' `
-Password $Password -MailNickname $AdminUsername

# Check if the new user script failed
if ($Log -contains 'ERROR: '+$AdminEmail+' does not exist.') {
    Write-Error 'ERROR: Admin user has not been created.'
    exit
}

# Add user to standard AD groups
$Log += Add-UserToSecGroupsAAD -Email $AdminEmail -Groups $Standard_AD_Admin
$Log += Add-UserToSecGroupsAAD -Email $AdminEmail -Groups $Standard_AD_InitSetup

# Add user to AD groups based on team
switch ($Team) {
    "Service Team Level 1" {
        Write-Host "Adding L1 groups"
        $Log += Add-UserToSecGroupsAAD -Email $AdminEmail -Groups $SD_AD_Admin
        $Log += Add-UserToSecGroupsAAD -Email $AdminEmail -Groups $SDL1_AD_Admin
        break
    }
    "Service Team Level 2" {
        $Log += Add-UserToSecGroupsAAD -Email $AdminEmail -Groups $SD_AD_Admin
        $Log += Add-UserToSecGroupsAAD -Email $AdminEmail -Groups $SDL2_AD_Admin
        break
    }
    "Service Team Level 3" {
        $Log += Add-UserToSecGroupsAAD -Email $AdminEmail -Groups $SD_AD_Admin
        $Log += Add-UserToSecGroupsAAD -Email $AdminEmail -Groups $SDL3_AD_Admin
        break
    }
    "Projects Team" {
        $Log += Add-UserToSecGroupsAAD -Email $AdminEmail -Groups $Projects_AD_Admin
        break
    }
    "Sales Team" {
        $Log += Add-UserToSecGroupsAAD -Email $AdminEmail -Groups $Sales_AD_Admin
        break
    }
    "Leadership Team" {
        $Log += Add-UserToSecGroupsAAD -Email $AdminEmail -Groups $Leadership_AD_Admin
        break
    }
}

## END OF SCRIPT ##
Write-Host "`nDisconnecting from Azure AD..."
Disconnect-AzureAD -Confirm:$false

# Spits out a JSON log for the user to copy
$json = @"
{
"AdminEmail":$AdminEmail,
"Password": "$Password",
"Log":"$Log"
}
"@

# Writes the JSON log to the PowerShell window
Write-Host $json