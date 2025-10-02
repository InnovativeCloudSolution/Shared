param(
    [string]$FirstName = '',
    [string]$LastName = '',
    [string]$DisplayName = '',
    [string]$Email = '',
    [string]$MobilePhone = '',
    [string]$Title = '',
    [string]$Team = '',
    [string]$AdminRequired = 'False',
    [string]$AdminEmail = '',
    [string]$AdminDisplayName = '',
    [string]$Password = '',
    [string]$ScriptAdmin = '',
    [string]$TenantId = ''
)

# Function for adding user to a group
function Add-UserToSecGroup {
    param (
        $UserObjectId = '',
        $GroupName = ''
    )

    $GroupObjectId = Get-AzureADGroup -SearchString $GroupName
    try {
        Add-AzureADGroupMember -ObjectId $GroupObjectId -RefObjectId $UserObjectId
        Write-Output "$Email has been added to $GroupName.`r`n"
    }
    catch { 
        Write-Output "ERROR: $Email has not been added to $GroupName.`r`n" 
    }
}

Import-Module AzureAD
Write-Host "`nConnecting to Azure AD..."
try { Connect-AzureAD -TenantId $TenantId -AccountId $ScriptAdmin }
catch {
    Write-Output "ERROR: Unable to connect to Azure AD.`r`n" 
    exit
}

## STANDARD ACCOUNT GROUPS ##

# Define groups that every user is added to
$StandardGroups = @(
    'App_ConnectwiseManage'
    'lp.synchronizedusers'
    'SG.Device.Win10.WiFi'
    'SG.License.M365E5.StandardUser'
)

# Define temporary groups for setting up the user before they arrive
$InitialSetupGroups = @(
    'MFA Disable'
    'SG.Azure.BlockNonManganoIP'
)

# Define groups for the service delivery team
$ServiceDeliveryGroups = @(
    'lp.servicedesk'
)

# Define groups for the L1 members of the service delivery team
$ServiceDeliveryL1Groups = @(
    'SG.Role.MIT.L1'
)

# Define groups for the L2 members of the service delivery team
$ServiceDeliveryL2Groups = @(
    'SG.Device.Win10.EnableUSBData'
    'SG.Role.MIT.L2'
)

# Define groups for the L3 members of the service delivery team
$ServiceDeliveryL3Groups = @(
    'SG.Device.Win10.EnableUSBData'
    'SG.License.Project.Plan5'
    'SG.License.Visio.Plan2'
    'SG.Role.MIT.L3'
)

# Define groups for the sales team
$SalesGroups = @(
    'DuoSecurity'
    'SG.App.Bullphish'
    'SG.App.ConnectWiseSell'
    'SG.License.PowerApp'
    'SG.License.Visio.Plan1'
    'SG.Role.AutomationAccount.AutomationOperator'
    'SG.Role.PaloAlto.VPN'
)

# Define groups for the project team
$ProjectsGroups = @(
    'lp.projects'
    'SG.Device.Win10.EnableUSBData'
    'SG.License.PowerApp'
    'SG.License.Project.Plan5'
    'SG.License.Visio.Plan2'
    'SG.Role.AutomationAccount.AutomationOperator'
)

# Define groups for the leadership team
$LeadershipGroups = @(
    'LeadershipGroup'
    'lp.directors'
    'SG.Device.Win10.EnableUSBData'
)

## ADMIN ACCOUNT GROUPS ##

# Define groups that every admin account is added to
$AdminGroups = @(
    'AdminAgents'
    'SG.License.M365E5.Admin'
    'SG.Role.ITGlueSSO'
)

# Define groups for the service delivery team
$ServiceDeliveryGroups_Admin = @(
    
)

# Define groups for the L1 members of the service delivery team
$ServiceDeliveryL1Groups_Admin = @(
    'SG.Role.LogicMonitor.Reader'
)

# Define groups for the L2 members of the service delivery team
$ServiceDeliveryL2Groups_Admin = @(
    'SG.Role.LogicMonitor.Contributor'
)

# Define groups for the L3 members of the service delivery team
$ServiceDeliveryL3Groups_Admin = @(
    'SG.Role.LogicMonitor.Admin'
)

# Define groups for the sales team
$SalesGroups_Admin = @(
    'SG.Role.LogicMonitor.Reader'
)

# Define groups for the project team
$ProjectsGroups_Admin = @(
    'SG.Role.LogicMonitor.Admin'
)

# Define groups for the leadership team
$LeadershipGroups_Admin = @(
    'SG.Role.LogicMonitor.Contributor'
)

# Makes the password profile - required for Azure AD module
$PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
$PasswordProfile.Password = $Password
$PasswordProfile.ForceChangePasswordNextLogin = $false

# Creates the mail nickname and the admin mail nickname - new user cannot be created without this
$MailNickname = $FirstName + $LastName
$AdminMailNickname = "adm.$MailNickname"

# Creates the new user account
try {
    Write-Output "`r`nCreating the new user`r`n"
    New-AzureADUser -DisplayName $DisplayName -GivenName $FirstName -Surname $LastName -PasswordProfile $PasswordProfile `
    -UserPrincipalName $Email -AccountEnabled $true -TelephoneNumber "07 3151 9000" -Mobile $MobilePhone -City "Hamilton" `
    -StreetAddress "9 Hercules Street" -PostalCode "4007" -State "Queensland" -Country "AU" -MailNickname $MailNickname`
    -Title $Title -UsageLocation "AU"
    Write-Output "New user account created with username: $Email `r`n"
}
catch { Write-Output "ERROR: Unable to create $Email.`r`n" }

# Grabs the user from Azure AD to get the object ID
try { $UserAccount = Get-AzureADUser -ObjectId $Email }
catch { 
    Write-Output "ERROR: $Email does not exist." 
    exit
}

# Sets new user variables based on parameters given

# Adds user to standard and temporary groups
foreach ($GroupName in $StandardGroups) { Add-UserToSecGroup -UserObjectId $UserAccount.ObjectId -GroupName $GroupName }
foreach ($GroupName in $InitialSetupGroups) { Add-UserToSecGroup -UserObjectId $UserAccount.ObjectId -GroupName $GroupName }

# Switch statement for adding user account to groups based on team
switch ($Team) {
    "Service Team Level 1" {
        foreach ($GroupName in $ServiceDeliveryGroups) { Add-UserToSecGroup -UserObjectId $UserAccount.ObjectId -GroupName $GroupName }
        foreach ($GroupName in $ServiceDeliveryL1Groups) { Add-UserToSecGroup -UserObjectId $UserAccount.ObjectId -GroupName $GroupName }
        break
    }
    "Service Team Level 2" {
        foreach ($GroupName in $ServiceDeliveryGroups) { Add-UserToSecGroup -UserObjectId $UserAccount.ObjectId -GroupName $GroupName }
        foreach ($GroupName in $ServiceDeliveryL2Groups) { Add-UserToSecGroup -UserObjectId $UserAccount.ObjectId -GroupName $GroupName }
        break
    }
    "Service Team Level 3" {
        foreach ($GroupName in $ServiceDeliveryGroups) { Add-UserToSecGroup -UserObjectId $UserAccount.ObjectId -GroupName $GroupName }
        foreach ($GroupName in $ServiceDeliveryL3Groups) { Add-UserToSecGroup -UserObjectId $UserAccount.ObjectId -GroupName $GroupName }
        break
    }
    "Projects Team" {
        foreach ($GroupName in $ProjectsGroups) { Add-UserToSecGroup -UserObjectId $UserAccount.ObjectId -GroupName $GroupName }
        break
    }
    "Sales Team" {
        foreach ($GroupName in $SalesGroups) { Add-UserToSecGroup -UserObjectId $UserAccount.ObjectId -GroupName $GroupName }
        break
    }
    "Leadership Team" {
        foreach ($GroupName in $LeadershipGroups) { Add-UserToSecGroup -UserObjectId $UserAccount.ObjectId -GroupName $GroupName }
        break
    }
}

# Makes the admin account (if required)
if ($AdminRequired -eq 'True') {
    try {
        Write-Output "`r`nCreating the new user`r`n"
        New-AzureADUser -DisplayName $AdminDisplayName -PasswordProfile $PasswordProfile -UserPrincipalName $AdminEmail `
        -AccountEnabled $true -Mobile $MobilePhone -UsageLocation "AU" -MailNickName $AdminMailNickname
        Write-Output "AD Account created with Username: $AdminEmail `r`n"
    }
    catch { Write-Output "Error: $AdminEmail already exists`r`n" }

    # Grabs the user from Azure AD to get the object ID
    try { $AdminAccount = Get-AzureADUser -ObjectId $AdminEmail }
    catch { 
        Write-Output "ERROR: $AdminEmail does not exist." 
        exit
    }

    # Adds admin account to admin and temporary groups
    foreach ($GroupName in $AdminGroups) { Add-UserToSecGroup -UserObjectId $AdminAccount.ObjectId -GroupName $GroupName }
    foreach ($GroupName in $InitialSetupGroups) { Add-UserToSecGroup -UserObjectId $AdminAccount.ObjectId -GroupName $GroupName }

    # Switch statement for adding admin account to groups based on team
    switch ($Team) {
        "Service Team Level 1" {
            foreach ($GroupName in $ServiceDeliveryGroups_Admin) { Add-UserToSecGroup -UserObjectId $AdminAccount.ObjectId `-GroupName $GroupName }
            foreach ($GroupName in $ServiceDeliveryL1Groups_Admin) { Add-UserToSecGroup -UserObjectId $AdminAccount.ObjectId -GroupName $GroupName }
            break
        }
        "Service Team Level 2" {
            foreach ($GroupName in $ServiceDeliveryGroups_Admin) { Add-UserToSecGroup -UserObjectId $AdminAccount.ObjectId -GroupName $GroupName }
            foreach ($GroupName in $ServiceDeliveryL2Groups_Admin) { Add-UserToSecGroup -UserObjectId $AdminAccount.ObjectId -GroupName $GroupName }
            break
        }
        "Service Team Level 3" {
            foreach ($GroupName in $ServiceDeliveryGroups_Admin) { Add-UserToSecGroup -UserObjectId $AdminAccount.ObjectId -GroupName $GroupName }
            foreach ($GroupName in $ServiceDeliveryL3Groups_Admin) { Add-UserToSecGroup -UserObjectId $AdminAccount.ObjectId -GroupName $GroupName }
            break
        }
        "Projects Team" {
            foreach ($GroupName in $ProjectsGroups_Admin) { Add-UserToSecGroup -UserObjectId $AdminAccount.ObjectId -GroupName $GroupName }
            break
        }
        "Sales Team" {
            foreach ($GroupName in $SalesGroups_Admin) { Add-UserToSecGroup -UserObjectId $AdminAccount.ObjectId -GroupName $GroupName }
            break
        }
        "Leadership Team" {
            foreach ($GroupName in $LeadershipGroups_Admin) { Add-UserToSecGroup -UserObjectId $AdminAccount.ObjectId -GroupName $GroupName }
            break
        }
    }
}