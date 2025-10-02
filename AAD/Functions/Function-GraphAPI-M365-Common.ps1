
# Exchange Online: Mail, Calendar, Contacts
function Convert-GraphUserMailboxToShared {
    param (
        [string]$userPrincipalName
    )
    # Assuming this operation uses Exchange Online cmdlets
    Set-Mailbox -Identity $userPrincipalName -Type Shared
}

function Hide-GraphUserFromGAL {
    param (
        [string]$userPrincipalName
    )
    $body = @{
        showInAddressList = $false
    } | ConvertTo-Json

    Invoke-GraphRequest -method PATCH -url "https://graph.microsoft.com/v1.0/users/$userPrincipalName" -body $body
}

function Remove-GraphUserFromSharedMailbox {
    param (
        [string]$userPrincipalName,
        [string]$sharedMailbox
    )
    Remove-MailboxPermission -Identity $sharedMailbox -User $userPrincipalName -AccessRights FullAccess -InheritanceType All
}

function Remove-GraphUserFromDistributionList {
    param (
        [string]$userPrincipalName,
        [string]$distributionList
    )
    Remove-DistributionGroupMember -Identity $distributionList -Member $userPrincipalName
}

# SharePoint: Sites, Lists, and List Items
# OneDrive: Files
# Teams: Teams, Channels, and Messages
# Planner: Tasks and Plans
# Intune: Devices and Policies

# These services would have similar functions like:
# Create, read, update, delete operations
# Managing memberships, permissions, etc.
