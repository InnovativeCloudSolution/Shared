param(
    [string]$Email = '',
    [string]$Team = '',
    [string]$ScriptAdmin = ''
)

Import-Module ExchangeOnlineManagement
Write-Host "`nConnecting to Exchange Online..."
Connect-ExchangeOnline -UserPrincipalName $ScriptAdmin -ShowBanner:$false -ShowProgress $true

# Function to add a user to a distribution group
function Add-UserToDistGroup {
    param (
        $Email = '',
        $GroupName = ''
    )

    try {
        Add-DistributionGroupMember -Identity $GroupName -Member $Email
        Write-Output "The user $Email has been added to the $GroupName group.`r`n"
    }
    catch { Write-Output "ERROR: $Email has not been added to $GroupName.`r`n" }
}

# Function to grant a user permissions to access a shared mailbox
function Add-UserToSharedMailbox {
    param (
        $Email = '',
        $GroupName = ''
    )

    try {
        Add-MailboxPermission -Identity $GroupName -User $Email -AccessRights FullAccess
        Write-Output "The user $Email has been added to the $GroupName group.`r`n"
    }
    catch { Write-Output "ERROR: $Email has not been added to $GroupName.`r`n" }
}

# Define groups that every user is added to
$StandardGroups = @(
    '!All Team'
    'Email Signature Group'
)

# Define groups for the service delivery team
$ServiceDeliveryGroups = @(
    '!All Techs'
)

# Define groups for the L1 members of the service delivery team
$ServiceDeliveryL1Groups = @(
    'Service Desk Level 1'
)

# Define groups for the L2 members of the service delivery team
$ServiceDeliveryL2Groups = @(
    'Service Desk Level 2'
)

# Define groups for the L3 members of the service delivery team
$ServiceDeliveryL3Groups = @(
    'Service Desk Level 3'
)

# Define groups for the sales team
$SalesGroups = @(
    '!All Sales'
)

# Define groups for the project team
$ProjectsGroups = @(
    
)

# Define groups for the leadership team
$LeadershipGroups = @(
    
)

$StandardSharedMailboxes = @(
    'sms@manganoit.com.au'
)

# Define groups for the service delivery team
$ServiceDeliverySharedMailboxes = @(
    
)

# Define SharedMailboxes for the L1 members of the service delivery team
$ServiceDeliveryL1SharedMailboxes = @(
    
)

# Define SharedMailboxes for the L2 members of the service delivery team
$ServiceDeliveryL2SharedMailboxes = @(
    
)

# Define SharedMailboxes for the L3 members of the service delivery team
$ServiceDeliveryL3SharedMailboxes = @(
    
)

# Define SharedMailboxes for the sales team
$SalesSharedMailboxes = @(
    'Mangano IT - Sales'
    'Tender Opps'
    'agmt.balance'
    'Mangano IT - Complex Data'
)

# Define SharedMailboxes for the project team
$ProjectsSharedMailboxes = @(
    'projectupdates'
)

# Define SharedMailboxes for the leadership team
$LeadershipSharedMailboxes = @(
    
)

# Checks to make sure the user is licensed before continuing
while ($null -eq (Get-Mailbox $Email -ErrorAction SilentlyContinue).Name) {
    Write-Host "User not found in Exchange Online. Please check that license is assigned and the user exists `in Exchange Online"
    pause
}

Enable-Mailbox -Identity $Email -Archive

# Adds user to standard distribution groups and shared mailboxes
foreach ($GroupName in $StandardGroups) { Add-UserToDistGroup -Email $Email -GroupName $GroupName }
foreach ($GroupName in $StandardSharedMailboxes) { Add-UserToSharedMailbox -Email $Email -GroupName $GroupName }

# Adds user to standard distribution groups and will also have shared mailboxes later (only one now so meh)
switch ($Team) {
    "Service Team Level 1" {
        foreach ($GroupName in $ServiceDeliveryGroups) { Add-UserToDistGroup -Email $Email -GroupName $GroupName }
        foreach ($GroupName in $ServiceDeliveryL1Groups) { Add-UserToDistGroup -Email $Email -GroupName $GroupName }
        foreach ($GroupName in $StandardSharedMailboxes) { Add-UserToSharedMailbox -Email $Email -GroupName $GroupName }
        foreach ($GroupName in $ServiceDeliveryL1SharedMailboxes) { Add-UserToSharedMailbox -Email $Email -GroupName $GroupName }
        break
    }
    "Service Team Level 2" {
        foreach ($GroupName in $ServiceDeliveryGroups) { Add-UserToDistGroup -Email $Email -GroupName $GroupName }
        foreach ($GroupName in $ServiceDeliveryL2Groups) { Add-UserToDistGroup -Email $Email -GroupName $GroupName }
        foreach ($GroupName in $StandardSharedMailboxes) { Add-UserToSharedMailbox -Email $Email -GroupName $GroupName }
        foreach ($GroupName in $ServiceDeliveryL2SharedMailboxes) { Add-UserToSharedMailbox -Email $Email -GroupName $GroupName }
        break
    }
    "Service Team Level 3" {
        foreach ($GroupName in $ServiceDeliveryGroups) { Add-UserToDistGroup -Email $Email -GroupName $GroupName }
        foreach ($GroupName in $ServiceDeliveryL3Groups) { Add-UserToDistGroup -Email $Email -GroupName $GroupName }
        foreach ($GroupName in $ServiceDeliverySharedMailboxes) { Add-UserToSharedMailbox -Email $Email -GroupName $GroupName }
        foreach ($GroupName in $ServiceDeliveryL3SharedMailboxes) { Add-UserToSharedMailbox -Email $Email -GroupName $GroupName }
        break
    }
    "Projects Team" {
        foreach ($GroupName in $ProjectsGroups) { Add-UserToDistGroup -Email $Email -GroupName $GroupName }
        foreach ($GroupName in $ProjectsSharedMailboxes) { Add-UserToSharedMailbox -Email $Email -GroupName $GroupName }
        break
    }
    "Sales Team" {
        foreach ($GroupName in $SalesGroups) { Add-UserToDistGroup -Email $Email -GroupName $GroupName }
        foreach ($GroupName in $SalesSharedMailboxes) { Add-UserToSharedMailbox -Email $Email -GroupName $GroupName }
        break
    }
    "Leadership Team" {
        foreach ($GroupName in $LeadershipGroups) { Add-UserToDistGroup -Email $Email -GroupName $GroupName }
        foreach ($GroupName in $LeadershipSharedMailboxes) { Add-UserToSharedMailbox -Email $Email -GroupName $GroupName }
        break
    }
}

# Builds link to calendar
$CalendarIdentity = "$Email`:\calendar"

# Get calendar permissions
$CalendarPermissions = Get-MailboxFolderPermission $CalendarIdentity

# Fix up calendar permissions? Not sure why we need this, was in the old script
foreach ($permission in $calendarPermissions)
{
    if ($permission.User.DisplayName -ne "Default") { continue }
    if ($permission.AccessRights -notcontains 'LimitedDetails') { Set-MailboxFolderPermission -User "Default" `
    -AccessRights 'LimitedDetails' -Identity $CalendarIdentity }
    break
}