<#

Mangano IT - New User Script (Module-Based Version)
Created by: Gabriel Nugent
Version: 1.7.1

Please make sure you have the included module at this location:
C:\Modules\MIT-NewUserModule\MIT-NewUserModule.ps1

The script will not work without its module!

#>

param(
    [Parameter(Mandatory=$true)][string]$FirstName = '',
    [Parameter(Mandatory=$true)][string]$LastName = '',
    [string]$OfficePhone = '',
    [string]$MobilePhone = '',
    [Parameter(Mandatory=$true)][string]$Title = '',
    [Parameter(Mandatory=$true)][string]$Manager = '',
    [switch]$AdminRequired,
    [string]$Password = ''
)

## MODULES ##
Import-Module AzureAD
Import-Module ExchangeOnlineManagement
Import-Module MicrosoftTeams
Import-Module ConnectWiseManageAPI
Import-Module -Name 'C:\Scripts\Repositories\PowerShell Scripts\MIT\MIT-NewUser\MIT-NewUserModule\MIT-NewUserModule' -Force

## DEFINE STANDARD VARIABLES FOR ALL ACCOUNTS ##
# Define groups that every user is added to
$Standard_AD = @(
    'SG.Role.SSO.CWManage'
    'SG.LastPass.SynchronizedUsers'
    'SG.Device.Windows.WiFi'
    'SG.License.M365E5.StandardUser'
    'SG.Role.AllStaff'
    'SG.Role.PaloAlto.VPN'
    'Team Mangano'  # Teams channel
    'Tech Team'  # Teams channel
)

# Define temporary groups for setting up the user before they arrive
$Standard_AD_InitSetup = @(
    'SG.Policy.AzureCA.DisableMFA'
    'SG.Policy.AzureCA.BlockNonManganoIP'
)

$Standard_AD_Admin = @(
    'AdminAgents'
    'SG.License.M365E5.Admin'
    'SG.Role.SSO.ITGlue'
)

$Standard_Distribution = @(
    '!All Team'
    'Email Signature Group'
)

$Standard_SharedMailboxes = @(
    'sms@manganoit.com.au'
)

# Define groups for the service delivery team
$SD_AD = @(
    'SG.LastPass.ServiceDesk'
)

$SD_AD_Admin = @(

)

$SD_Distribution = @(
    '!All Techs'
    'Service Delivery Team'
)

$SD_SharedMailboxes = @(
    
)

# Define groups for the L1 members of the service delivery team
$SDL1_AD = @(
    'SG.Role.ServiceDesk.L1'
    'SG.Role.Meraki.Read'
    'SG.Officevibe.Service.L1'
)

$SDL1_AD_Admin = @(
    'SG.Role.LogicMonitor.Reader'
    'SG.Role.Admin.L1'
)

$SDL1_Distribution = @(
    'Service Desk Level 1'
)

$SDL1_SharedMailboxes = @(
    
)

# Define groups for the L2 members of the service delivery team
$SDL2_AD = @(
    'SG.Device.Windows.EnableUSBData'
    'SG.Role.ServiceDesk.L2'
    'SG.Role.Meraki.Write'
    'SG.Officevibe.Service.L2'
)

$SDL2_AD_Admin = @(
    'SG.Role.LogicMonitor.Contributor'
    'SG.Role.Admin.L2'
    'SG.Role.LocalAdmin'
)

$SDL2_Distribution = @(
    'Service Desk Level 2'
)

$SDL2_SharedMailboxes = @(
    
)

# Define groups for the L3 members of the service delivery team
$SDL3_AD = @(
    'SG.Device.Windows.EnableUSBData'
    'SG.License.Project.Plan5'
    'SG.License.Visio.Plan2'
    'SG.App.Microsoft365.Project'
    'SG.App.Microsoft365.Visio'
    'SG.Role.ServiceDesk.L3'
    'SG.Role.Meraki.Write'
    'SG.Officevibe.Service.L3'
)

$SDL3_AD_Admin = @(
    'SG.Role.LogicMonitor.Admin'
    'SG.Role.Admin.L3'
    'SG.Role.LocalAdmin'
)

$SDL3_Distribution = @(
    'Service Desk Level 3'
)

$SDL3_SharedMailboxes = @(
    
)

# Define groups for the sales team
$Sales_AD = @(
    'SG.App.DuoSecurity'
    'SG.App.Bullphish'
    'SG.Role.SSO.CWSell'
    'SG.Role.Sales'
    'SG.License.PowerApp'
    'SG.License.Visio.Plan1'
    'SG.Role.AutomationAccount.AutomationOperator'
    '!All Sales'        # Microsoft 365 group
    'SG.Role.Meraki.Read'
    'SG.Officevibe.Sales'
    'Marketing Team'  # Teams channel
    'Sales Team'  # Teams channel
)

$Sales_AD_Admin = @(
    'SG.Role.LogicMonitor.Reader'
)

$Sales_Distribution = @(
    
)

$Sales_SharedMailboxes = @(
    'Mangano IT - Sales'
    'Tender Opps'
    'agmt.balance'
    'Mangano IT - Complex Data'
)

# Define groups for the project team
$Projects_AD = @(
    'SG.LastPass.Projects'
    'SG.License.PowerApp'
    'SG.License.Project.Plan5'
    'SG.License.Visio.Plan2'
    'SG.App.Microsoft365.Project'
    'SG.App.Microsoft365.Visio'
    'SG.Role.AutomationAccount.AutomationOperator'
    'SG.Role.Projects'
    'SG.Role.Meraki.Write'
    'SG.Officevibe.Projects'
    'Project Delivery'  # Teams channel
)

$Projects_AD_Admin = @(
    'SG.Role.LogicMonitor.Admin'
    'SG.Role.Admin.L3'
    'SG.Role.LocalAdmin'
)

$Projects_Distribution = @(
    '!All Techs'
)

$Projects_SharedMailboxes = @(
    'projectupdates'
)

# Define groups for the leadership team
$Leadership_AD = @(
    'SG.Role.Leadership'
    'SG.LastPass.Leadership'
    'SG.Role.Meraki.Read'
    'Internal Systems'  # Teams channel
    'Leadership Team'  # Teams channel
    'Marketing Team'  # Teams channel
    'Project Delivery'  # Teams channel
    'Recruitment'  # Teams channel
    'Sales Team'  # Teams channel
)

$Leadership_AD_Admin = @(
    'SG.Role.LogicMonitor.Contributor'
)

$Leadership_Distribution = @(
    
)

$Leadership_SharedMailboxes = @(
    
)

## Program details - please remember to update the version number when making changes
Write-Host "`nMangano IT - Internal New User" -ForegroundColor Yellow
Write-Host "Version: " -ForegroundColor yellow -NoNewLine; Write-Host "1.6"
Write-Host "Created by: " -ForegroundColor yellow -NoNewLine; Write-Host "Gabriel Nugent"
Write-Host "`nPlease remember to copy the folder " -NoNewLine; Write-Host "MIT-NewUserModule " -ForegroundColor yellow -NoNewline
Write-Host "to " -NoNewLine; Write-Host "C:\Scripts\Repositories\PowerShell Scripts\MIT\MIT-NewUser\MIT-NewUserModule\MIT-NewUserModule.psm1" -ForegroundColor yellow

# Confirms the phone number starts with a 0
if ($MobilePhone[0] -eq "4") { $MobilePhone = "0$($MobilePhone)" }

# Add generic phone number if not provided
if ($OfficePhone -eq '') { $OfficePhone = '07 3151 9000' }

$Log="The parameters that have been provided are:`n"
$Log += "First Name - $FirstName`n"
$Log += "Last Name - $LastName`n"
$Log += "Title - $Title`n"
$Log += "Mobile Phone - $MobilePhone`n"

Write-Host "`nThe parameters that have been provided are:" -ForegroundColor yellow
Write-Host "First Name - $FirstName"
Write-Host "Last Name - $LastName"
Write-Host "Title - $Title"
Write-Host "Mobile Phone - $MobilePhone"

#Company name in CWM
#Email domain (for CWM and users)
$Domain = '@manganoit.com.au'
$TenantID = "5792a6c1-f4fe-466b-b97c-10eaf4fb3122" #The ManganoIT Tenancy

# Grab email address of person running script
$ScriptAdmin = $(Write-Host "`nAdmin login: " -ForegroundColor yellow -NoNewLine; Read-Host)
if ($ScriptAdmin -eq "gabe") { $ScriptAdmin = "adm.gabriel.nugent@manganoit.com.au" }

# Generates a password if one isn't already provided
if ($Password -eq '') {
    $Log += "`nGenerating password through DinoPass`n"
    # Loops until the password is at least 12 characters long
    while ($Password.Length -lt 12) { 
        $Password += Invoke-restmethod -uri "https://www.dinopass.com/password/strong"
        if ($Password -like "*+*") { $Password = '' }
    }
}

# Create the user's details
$Username = $FirstName+'.'+$LastName
$Email = $Username+$Domain
$DisplayName = $FirstName+' '+ $LastName

<# Redundant right now - will leave here, but commented out

$Log += "`nConstructing the username and email for the user`n"
$Log += "Username: $Username"
$Log += "Email: $Email" #>

# Give a list of teams for the new user to be in - asks for input
$Team = ''

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

# Create the admin user's details
if ($AdminRequired) {
    #Admin Firstname Surname (No firstname/surname)
    Write-Host "`nConstructing the username and email for the user`n"
    $AdminUsername = "adm."+$FirstName+'.'+ $LastName
    $AdminEmail = $AdminUsername+$Domain
    $AdminDisplayName = "Admin "+$FirstName+' '+ $LastName

    <# Redundant right now - will leave here, but commented out
    
    $Log += "`nConstructing the username and email for the user"
    $Log += "`nUsername: $AdminUsername"
    $Log += "`nEmail: $AdminEmail"
    $Log += "`nDisplay name: $AdminDisplayName" #>
}

# Attempt to connect to Azure AD
Write-Host "`nConnecting to Azure AD..."
try { Connect-AzureAD -TenantId $TenantID -AccountId $ScriptAdmin }
catch {
    Write-Error "ERROR: Unable to connect to Azure AD.`n" 
    exit
}

Write-Host "Creating the user account..."
$Log += Add-UserAccountAAD -FirstName $FirstName -LastName $LastName -DisplayName $DisplayName -JobTitle $Title `
-Email $Email -Mobile $MobilePhone -TelephoneNumber $OfficePhone -City 'Hamilton' -StreetAddress '9 Hercules Street' `
-PostalCode '4007' -State 'Queensland' -Country 'AU' -UsageLocation 'AU' -Password $Password -MailNickname $Username `
-Manager $Manager

# Check if the new user script failed
if ($Log -contains 'ERROR: '+$Email+' does not exist.') {
    Write-Error 'ERROR: New user has not been created, exiting script...'
    exit
}

# Add user to standard AD groups
$Log += Add-UserToSecGroupsAAD -Email $Email -Groups $Standard_AD
$Log += Add-UserToSecGroupsAAD -Email $Email -Groups $Standard_AD_InitSetup

# Add user to AD groups based on team
switch ($Team) {
    "Service Team Level 1" {
        $Log += Add-UserToSecGroupsAAD -Email $Email -Groups $SD_AD
        $Log += Add-UserToSecGroupsAAD -Email $Email -Groups $SDL1_AD
        break
    }
    "Service Team Level 2" {
        $Log += Add-UserToSecGroupsAAD -Email $Email -Groups $SD_AD
        $Log += Add-UserToSecGroupsAAD -Email $Email -Groups $SDL2_AD
        break
    }
    "Service Team Level 3" {
        $Log += Add-UserToSecGroupsAAD -Email $Email -Groups $SD_AD
        $Log += Add-UserToSecGroupsAAD -Email $Email -Groups $SDL3_AD
        break
    }
    "Projects Team" {
        $Log += Add-UserToSecGroupsAAD -Email $Email -Groups $Projects_AD
        break
    }
    "Sales Team" {
        $Log += Add-UserToSecGroupsAAD -Email $Email -Groups $Sales_AD
        break
    }
    "Leadership Team" {
        $Log += Add-UserToSecGroupsAAD -Email $Email -Groups $Leadership_AD
        break
    }
}

if ($AdminRequired) {
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
}

Write-Host "Thanks! The account has been created! Please wait a few minutes while a license is added to the account!"

# Wait 300 seconds for licenses (that are applied by security group) to finish applying
for ($i=0; $i -le 300; $i++) {
    $Percent = [math]::Round($i/300*100)
    Write-Progress -Activity "Waiting for license addition" -Status "$Percent% Complete:" -PercentComplete $Percent;
    Start-Sleep -Seconds 1
}

# Connect to Exchange Online
Write-Host "`nConnecting to Exchange Online..."
Connect-ExchangeOnline -UserPrincipalName $ScriptAdmin -ShowBanner:$false -ShowProgress $true

# Checks to make sure the user is licensed before continuing
while ($null -eq (Get-Mailbox $Email -ErrorAction SilentlyContinue).Name) {
    Write-Host "User not found in Exchange Online. Please check that license is assigned and the user exists `in Exchange Online"
    pause
}

Enable-Mailbox -Identity $Email -Archive
$Log += Optimize-CalendarPermissions -Email $Email

# Disabling for now - doesn't work on non-resource users
# $Log += Set-AutoAcceptCalendarInvites -Email $Email

# Connect to Microsoft Teams
# Disabled for now, no channels in use
# Write-Host "`nConnecting to Microsoft Teams..."
# Connect-MicrosoftTeams -TenantId $TenantId -AccountId $ScriptAdmin

# Adds user to standard distribution groups and shared mailboxes
$Log += Add-UserToDistGroups -Email $Email -Groups $Standard_Distribution
$Log += Add-UserToSharedMailboxes -Email $Email -Groups $Standard_SharedMailboxes

# Adds user to standard channels
# Disabled for now, no channels in use
# $Log+=Add-UserToChannels -Email $Email -Channels $Standard_Channels

# Adds user to standard distribution groups, shared mailboxes, and channels
switch ($Team) {
    "Service Team Level 1" {
        $Log += Add-UserToDistGroups -Email $Email -Groups $SD_Distribution
        $Log += Add-UserToSharedMailboxes -Email $Email -Groups $SD_SharedMailboxes
        # $Log += Add-UserToChannels -Email $Email -Channels $SD_Channels
        $Log += Add-UserToDistGroups -Email $Email -Groups $SDL1_Distribution
        $Log += Add-UserToSharedMailboxes -Email $Email -Groups $SDL1_SharedMailboxes
        # $Log += Add-UserToChannels -Email $Email -Channels $SDL1_Channels
        break
    }
    "Service Team Level 2" {
        $Log += Add-UserToDistGroups -Email $Email -Groups $SD_Distribution
        $Log += Add-UserToSharedMailboxes -Email $Email -Groups $SD_SharedMailboxes
        # $Log += Add-UserToChannels -Email $Email -Channels $SD_Channels
        $Log += Add-UserToDistGroups -Email $Email -Groups $SDL2_Distribution
        $Log += Add-UserToSharedMailboxes -Email $Email -Groups $SDL2_SharedMailboxes
        # $Log += Add-UserToChannels -Email $Email -Channels $SDL2_Channels
        break
    }
    "Service Team Level 3" {
        $Log += Add-UserToDistGroups -Email $Email -Groups $SD_Distribution
        $Log += Add-UserToSharedMailboxes -Email $Email -Groups $SD_SharedMailboxes
        # $Log += Add-UserToChannels -Email $Email -Channels $SD_Channels
        $Log += Add-UserToDistGroups -Email $Email -Groups $SDL3_Distribution
        $Log += Add-UserToSharedMailboxes -Email $Email -Groups $SDL3_SharedMailboxes
        # $Log += Add-UserToChannels -Email $Email -Channels $SDL3_Channels
        break
    }
    "Projects Team" {
        $Log += Add-UserToDistGroups -Email $Email -Groups $Projects_Distribution
        $Log += Add-UserToSharedMailboxes -Email $Email -Groups $Projects_SharedMailboxes
        # $Log += Add-UserToChannels -Email $Email -Channels $Projects_Channels
        break
    }
    "Sales Team" {
        $Log += Add-UserToDistGroups -Email $Email -Groups $Sales_Distribution
        $Log += Add-UserToSharedMailboxes -Email $Email -Groups $Sales_SharedMailboxes
        # $Log += Add-UserToChannels -Email $Email -Channels $Sales_Channels
        break
    }
    "Leadership Team" {
        $Log += Add-UserToDistGroups -Email $Email -Groups $Leadership_Distribution
        $Log += Add-UserToSharedMailboxes -Email $Email -Groups $Leadership_SharedMailboxes
        # $Log += Add-UserToChannels -Email $Email -Channels $Leadership_Channels
        break
    }
}

$Log += Add-CWMContact -FirstName $FirstName -LastName $LastName -Title $Title -Email $Email `
-Domain '@manganoit.com.au' -MobilePhone $MobilePhone -OfficePhone $OfficePhone -CompanyName 'Mangano IT'
-SiteId 3529

Write-Host 'The user account should now be finished. Please review the JSON log below for any errors.' -ForegroundColor Yellow

## END OF SCRIPT ##
Write-Host "`nDisconnecting from Exchange Online..."
Disconnect-ExchangeOnline -Confirm:$false
# Write-Host "`nDisconnecting from Microsoft Teams..."
# Disconnect-MicrosoftTeams -Confirm:$false
Write-Host "`nDisconnecting from Azure AD..."
Disconnect-AzureAD -Confirm:$false

# Spits out a JSON log for the user to copy
$json = @"
{
"Email": "$Email",
"Username": "$Username",
"Password": "$Password",
"AdminEmail":$AdminEmail,
"Log":"$Log"
}
"@

# Writes the JSON log to the PowerShell window
Write-Host $json
#Out-File -FilePath '.\'+$DisplayName+'.txt' -InputObject $json