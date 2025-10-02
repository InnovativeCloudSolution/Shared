<#

Mangano IT - Exit User Script
Created by: Gabriel Nugent
Version: 2.1

#>

param(
    [string]$Account = '',
    [string]$ResendAddress = ''
)

## MODULES ##

Import-Module AzureAD
Import-Module ExchangeOnlineManagement

## SCRIPT VARIABLES ##

# Mangano IT tenancy details
$TenantID = "5792a6c1-f4fe-466b-b97c-10eaf4fb3122"

# Create variable for the user account
$User = $null
$AdminUser = $null

# Build user details to output later
$Output = [ordered]@{
    User = ''
    UserId = ''
    SignInRevoked = $false
    AccountDisabled = $false
    ManagerRemoved = $false
    ConvertedToSharedMailbox = $false
    ManagerGrantedMailboxAccess = $false
    AutoReplyEnabled = $false
    Removed = [ordered]@{
        Groups = @()
        SharedMailboxes = @()
        SharedMailboxesSendAs = @()
        DistributionGroups = @()
        Licenses = @()
    }
    Remaining = [ordered]@{
        Groups = @()
        SharedMailboxes = @()
        SharedMailboxesSendAs = @()
        DistributionGroups = @()
    }
}

$AdminOutput = [ordered]@{
    User = ''
    UserId = ''
    SignInRevoked = $false
    AccountDisabled = $false
    Removed = [ordered]@{
        Groups = @()
        Licenses = @()
    }
    Remaining = [ordered]@{
        Groups = @()
    }
}

## FUNCTIONS ##

function LocateUser {
    param (
        [Parameter(Mandatory)][string]$UserPrincipalName
    )
    
    try {
        $User = Get-AzureADUser -ObjectId $UserPrincipalName
        Write-Host "$UserPrincipalName has been found successfully."
    }
    catch {
        Write-Warning "$UserPrincipalName does not exist."
        $User = $null
    }

    return $User
}

function RevokeSignIn {
    param (
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][string]$UserPrincipalName
    )

    try {
        Revoke-AzureADUserAllRefreshToken -ObjectId $UserId
        Write-Host "$UserPrincipalName's sessions have been revoked."
        return $true
    } catch {
        Write-Error "Unable to revoke sign in for $UserPrincipalName"
        Write-Error $_
        return $false
    }
}

function DisableAccount {
    param (
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][string]$UserPrincipalName
    )

    try {
        Write-Host "Blocking sign in for $UserPrincipalName..."
        Set-AzureADUser -ObjectID $UserPrincipalName -AccountEnabled $false | Out-Null
        Write-Host "Sign in for $UserPrincipalName has been blocked."
        return $true
    }
    catch {
        Write-Error "$UserPrincipalName may still be able to sign in."
        Write-Error $_
        return $false
    }
}

function GetAndRemoveManager {
    param (
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][string]$UserPrincipalName
    )

    try {
        Write-Host "Fetching the manager for $UserPrincipalName..."
        $Manager = Get-AzureADUserManager -ObjectId $UserId
        if ($null -ne $Manager) {
            Write-Host "Manager located for $UserPrincipalName."
        } else {
            Write-Host "$UserPrincipalName does not have a manager."
        }
    } catch {
        Write-Error "Unable to fetch manager for $UserPrincipalName"
        Write-Error $_
        $Manager = $null
    }

    if ($null -ne $Manager) {
        try {
            Write-Host "Removing the manager for $UserPrincipalName..."
            Remove-AzureADUserManager -ObjectID $Account
            Write-Host "Manager for $UserPrincipalName has been removed."
        } catch {
            Write-Error "$UserPrincipalName has not had its manager removed."
            Write-Error $_
        }
    }

    return $Manager
}

function ExitFromExchange {
    param(
        [Parameter(Mandatory)][string]$EmailAddress
    )

    # Get the mailbox from Exchange
    $Mailbox = Get-Mailbox -Identity $EmailAddress

    # Get the DN of the mailbox 
    $DistinguishedName = $Mailbox.DistinguishedName

    # Create var for output
    $Result = @{
        Mailbox = $Mailbox
        Removed = @{
            SharedMailboxes = @()
            SharedMailboxesSendAs = @()
            DistributionGroups = @()
        }
        Remaining = @{
            SharedMailboxes = @()
            SharedMailboxesSendAs = @()
            DistributionGroups = @()
        }
    }

    # Set filter 
    $Filter = "Members -like ""$DistinguishedName"""

    # Get DLs with that user
    Write-Host "Searching for distribution groups..."
    $DistributionGroupsList = Get-DistributionGroup -ResultSize Unlimited -Filter $Filter

    # Get Shared Mailboxes the user has access to (previously had -RecipientTypeDetails SharedMailbox
    Write-Host "Searching for shared mailboxes..."
#    $SharedMailboxList = Get-Mailbox -ResultSize unlimited | Get-MailboxPermission | Where-Object {($_.User -like $EmailAddress)}
    $SharedMailboxList = Get-Mailbox -ResultSize Unlimited | ForEach-Object {
        Get-MailboxPermission -Identity $_.DistinguishedName  | Where-Object {($_.User -like $EmailAddress)}
    }


    # Get Shared Mailboxes with SendAs permission
    Write-Host "Searching for shared mailboxes with send as permissions..."
#    $SharedSendAsMailboxList = Get-Mailbox -ResultSize unlimited | Get-RecipientPermission | Where-Object {($_.Trustee -like $EmailAddress)}

    $SharedSendAsMailboxList = Get-Mailbox -ResultSize Unlimited | ForEach-Object {
        Get-RecipientPermission -Identity $_.DistinguishedName  | Where-Object {($_.Trustee -like $EmailAddress)}
    }


    # Removes user from all distribution groups
    foreach ($Item in $DistributionGroupsList) {
        $GroupIdentity = $Item.DisplayName
        try {
            Remove-DistributionGroupMember -Identity $Item.PrimarySmtpAddress -Member $DistinguishedName -BypassSecurityGroupManagerCheck -Confirm:$false
            Write-Host "Distribution Group: Removed $EmailAddress from $GroupIdentity"
            $Result.Removed.DistributionGroups += $GroupIdentity
        } catch {
            Write-Error "Distribution Group: Unable to remove $EmailAddress from $GroupIdentity"
            Write-Error $_
            $Result.Remaining.DistributionGroups += $GroupIdentity
        }
    }

    # Removes user from all shared mailboxes
    foreach ($Item in $SharedMailboxList) {
        $GroupIdentity = $Item.Identity
        try {
            Remove-MailboxPermission -Identity $Item.Identity -User $EmailAddress -AccessRights $Item.AccessRights -InheritanceType All -Confirm:$false
            Write-Host "Shared Mailbox: Removed $EmailAddress from $GroupIdentity"
            $Result.Removed.SharedMailboxes += $GroupIdentity
        } catch {
            Write-Error "Shared Mailbox: Unable to remove $EmailAddress from $GroupIdentity"
            Write-Error $_
            $Result.Remaining.SharedMailboxes += $GroupIdentity
        }
    }

    # Removes user from all shared mailboxes that are attached via send as permissions
    foreach ($Item in $SharedSendAsMailboxList) {
        $GroupIdentity = $Item.Identity
        try {
            Remove-RecipientPermission -Identity $Item.Identity -Trustee $EmailAddress -AccessRights $Item.AccessRights -Confirm:$false
            Write-Host "Shared Mailbox (Send As): Removed $EmailAddress from $GroupIdentity"
            $Result.Removed.SharedMailboxesSendAs += $GroupIdentity
        } catch {
            Write-Error "Shared Mailbox (Send As): Unable to remove $EmailAddress from $GroupIdentity"
            Write-Error $_
            $Result.Remaining.SharedMailboxesSendAs += $GroupIdentity
        }
    }

    return $Result
}

function HideFromGalAndConvertToShared {
    param(
        [Parameter(Mandatory)][string]$EmailAddress
    )
    
    try {
        Set-Mailbox -Identity $EmailAddress -HiddenFromAddressListsEnabled $true -Type Shared
        Write-Host "$EmailAddress has been hidden from the GAL and converted to a shared mailbox."
        return $true
    }
    catch {
        Write-Error "$EmailAddress has not been hidden from the GAL and converted to a shared mailbox."
        Write-Error $_
        return $false
    }
}

function GrantManagerMailboxAccess {
    param(
        [Parameter(Mandatory)][string]$UserPrincipalName,
        [Parameter(Mandatory)][string]$ManagerUserPrincipalName,
        [Parameter(Mandatory)][string]$ManagerName
    )

    try {
        Add-MailboxPermission -Identity $ManagerUserPrincipalName -User $UserPrincipalName -AccessRights FullAccess
        Write-Host "$ManagerName now has access to the shared mailbox for $UserPrincipalName"
        return $true
    }
    catch {
        Write-Error "$ManagerName does not have access to the shared mailbox for $UserPrincipalName"
        Write-Error $_
        return $false
    }
}

function SetOutOfOffice {
    param (
        [Parameter(Mandatory)][string]$DistinguishedName,
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)][string]$ResendAddress
    )
    
    # Create message
    $Message = "<html><body>Hello,<br><br>Thank you for reaching out. Please be advised that $DisplayName is no longer with Mangano IT."
    $Message += " This mailbox is no longer being monitored and we regret any inconvenience this may cause.<br><br>"
    $Message += "Please resend your email to $ResendAddress, and someone from our team will get back to you as soon as possible.<br><br>"
    $Message += "If you require immediate assistance, please call our team on 07 3151 9000.<br><br>"
    $Message += "Regards,<br>Mangano IT</body></html>"

    # Format time
    $StartTime = Get-Date
    $EndTime = $StartTime.AddDays(90)

    # Enable out of office
    $Parameters = @{
        Identity = $DistinguishedName
        AutoReplyState = 'Scheduled'
        InternalMessage = $Message
        ExternalMessage = $Message
        StartTime = $StartTime
        EndTime = $EndTime
    }

    try {
        Set-MailboxAutoReplyConfiguration @Parameters
        Write-Host "Auto reply has been enabled for $EmailAddress"
        return $true
    } catch {
        Write-Error "Auto reply has not been enabled for $EmailAddress"
        Write-Error $_
        return $false
    }
}

function RemoveFromGroups {
    param (
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][string]$UserPrincipalName
    )

    $UserPrincipalNameGroups = Get-AzureADUserMembership -ObjectId $UserId

    # Create var for output
    $Result = @{
        Removed = @{
            Groups = @()
        }
        Remaining = @{
            Groups = @()
        }
    }

    foreach ($Group in $UserPrincipalNameGroups) {
        $GroupName = $Group.DisplayName
        try {
            Remove-AzureADGroupMember -ObjectId $Group.ObjectId -MemberId $UserId
            Write-Host "M365 Group: $UserPrincipalName has been removed from $GroupName."
            $Result.Removed.Groups += $GroupName
        }
        catch {
            Write-Error "M365 Group: $UserPrincipalName has not been removed from $GroupName."
            Write-Error $_
            $Result.Remaining.Groups += $GroupName
        }
    }

    # Waits for security groups to be properly cleared out
    # Some licenses are attached to security groups
    for ($i=0; $i -le 30; $i++) {
        $Percent = [math]::Round($i/30*100)
        Write-Progress -Activity "Waiting for security groups to finish being removed" -Status "$Percent% Complete:" -PercentComplete $Percent;
        Start-Sleep -Seconds 1
    }

    return $Result
}

function RemoveLicenses {
    param (
        [Parameter(Mandatory)]$User
    )

    # Define var for removed licenses
    $RemovedLicenses = @()
    
    # Grab all of the user's licenses and remove them
    $Skus = $User | Select-Object -ExpandProperty AssignedLicenses | Select-Object SkuID, SkuPartNumber
    if ($User.Count -ne 0) {
        if ($Skus -is [array]) {
            $Licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
            for ($i=0; $i -lt $Skus.Count; $i++) {
                $RemovedLicenses += $Skus[$i].SkuPartNumber
                $Licenses.RemoveLicenses += (Get-AzureADSubscribedSku | Where-Object -Property SkuID -Value $Skus[$i].SkuId -EQ).SkuID   
            }
            Set-AzureADUserLicense -ObjectId $Account -AssignedLicenses $Licenses
            Write-Host "Licenses have been removed for $Account, minus those applied by a security group."
        } else {
            $Licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
            $RemovedLicenses += $Skus.SkuPartNumber
            $Licenses.RemoveLicenses = (Get-AzureADSubscribedSku | Where-Object -Property SkuID -Value $Skus.SkuId -EQ).SkuID
            Set-AzureADUserLicense -ObjectId $Account -AssignedLicenses $Licenses
            Write-Host "Licenses have been removed for $Account, minus those applied by a security group."
        }
    }

    return $RemovedLicenses
}

## SCRIPT ##

# Program details - please remember to update the version number when making changes
Write-Host "`nMangano IT - Exit Internal User" -ForegroundColor Yellow
Write-Host "Version: " -ForegroundColor yellow -NoNewLine; Write-Host "1.4"
Write-Host "Created by: " -ForegroundColor yellow -NoNewLine; Write-Host "Gabriel Nugent"

# Get details of user to exit if not provided
if ($Account -eq '') {
    $Account = $(Write-Host "`nEmail address of account to exit: " -ForegroundColor yellow -NoNewLine; Read-Host)
}

# Get details of resend address if not provided
if ($ResendAddress -eq '') {
    $ResendAddress = $(Write-Host "`nEmail address that the auto-reply will redirect to: " -ForegroundColor yellow -NoNewLine; Read-Host)
}

# Confirm details of user to exit
Write-Host "User to exit: " -ForegroundColor yellow -NoNewLine; Write-Host $Account
$Output.User = $Account

$AdminAccount = "adm."+$Account
$AdminOutput.User = $AdminAccount
$AdminExists = $true

# Connects to Azure AD
$ScriptAdmin = $(Write-Host "`nAdmin login: " -ForegroundColor yellow -NoNewLine; Read-Host)
if ($ScriptAdmin -eq "alex") { $ScriptAdmin = "adm.alex.williams@manganoit.com.au" }
Write-Host "`nConnecting to Azure AD..."
try { Connect-AzureAD -TenantId $TenantID -AccountId $ScriptAdmin }
catch {
    Write-Error "ERROR: Unable to connect to Azure AD."
    Write-Error $_ 
    exit
}

# Checks to see if the user exists before signing them out of active sessions and disabling their account
$User = LocateUser -UserPrincipalName $Account
if ($null -eq $User) {
    Write-Error "Incorrect user details, or you do not have the correct permissions for this operation.`nProgram will now terminate."
    Disconnect-AzureAD -Confirm:$false
    Start-Sleep -Seconds 2
    exit
} else {
    $Output.UserId = $User.ObjectId
    $Output.SignInRevoked = RevokeSignIn -UserId $User.ObjectId -UserPrincipalName $Account
    $Output.AccountDisabled = DisableAccount -UserId $User.ObjectId -UserPrincipalName $Account
}

# Checks to see if the admin account exists before signing them out of active sessions and disabling their account
$AdminUser = LocateUser -UserPrincipalName $AdminAccount
if ($null -eq $AdminUser) {
    $AdminExists = $false
} else {
    $AdminOutput.UserId = $AdminUser.ObjectId
    $AdminOutput.SignInRevoked = RevokeSignIn -UserId $AdminUser.ObjectId -UserPrincipalName $AdminAccount
    $AdminOutput.AccountDisabled = DisableAccount -UserId $AdminUser.ObjectId -UserPrincipalName $AdminAccount
}

# Gets the manager for assigning as the Exchange delegate
$Manager = GetAndRemoveManager -UserId $User.ObjectId -UserPrincipalName $Account
if ($null -ne $Manager) {
    $Output.ManagerRemoved = $true
}

# Connects to Exchange Online and Microsoft Teams with the given creds
Write-Host "`nConnecting to Exchange Online..."
Connect-ExchangeOnline -UserPrincipalName $ScriptAdmin -ShowBanner:$false -ShowProgress $true

# Remove from Exchange
$ExchangeExit = ExitFromExchange -EmailAddress $Account

# Grab groups for export, hides mailbox from GAL, converts to a shared mailbox
if ($null -ne $ExchangeExit.Mailbox) {
    $Output.Removed.SharedMailboxes = $ExchangeExit.Removed.SharedMailboxes
    $Output.Removed.SharedMailboxesSendAs = $ExchangeExit.Removed.SharedMailboxesSendAs
    $Output.Removed.DistributionGroups = $ExchangeExit.Removed.DistributionGroups
    $Output.Remaining.SharedMailboxes = $ExchangeExit.Remaining.SharedMailboxes
    $Output.Remaining.SharedMailboxesSendAs = $ExchangeExit.Remaining.SharedMailboxesSendAs
    $Output.Remaining.DistributionGroups = $ExchangeExit.Remaining.DistributionGroups
    HideFromGalAndConvertToShared -EmailAddress $Account
    SetOutOfOffice -DistinguishedName $DistinguishedName -ResendAddress $ResendAddress
}

# Grants manager full access to the mailbox
if ($null -ne $Manager) {
    $Output.GrantManagerMailboxAccess = GrantManagerMailboxAccess -UserPrincipalName $Account -ManagerUserPrincipalName $Manager.userPrincipalName`
    -ManagerName $Manager.displayName
}

# Removes security group membership
$GroupExit = RemoveFromGroups -UserId $User.ObjectId -UserPrincipalName $Account
$Output.Removed.Groups = $GroupExit.Removed.Groups
$Output.Remaining.Groups = $GroupExit.Remaining.Groups

# Removes licenses
$Output.Removed.Licenses = RemoveLicenses -User $User

# Removes licenses and security group membership for admin account (if it exists)
if ($true -eq $AdminExists) {
    # Remove admin account from all groups it's in
    $AdminGroupExit = RemoveFromGroups -UserId $AdminUser.ObjectId -UserPrincipalName $AdminAccount
    $AdminOutput.Removed.Groups = $AdminGroupExit.Removed.Groups
    $AdminOutput.Remaining.Groups = $AdminGroupExit.Remaining.Groups

    # Removes licenses
    $AdminOutput.Removed.Licenses = RemoveLicenses -User $AdminUser
}

## END OF SCRIPT ##
Write-Host "`nDisconnecting from Exchange Online..."
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnecting from Azure AD...`n`n"
Disconnect-AzureAD -Confirm:$false

# Spits out a JSON log for the user to copy
Write-Output $Output | ConvertTo-Json -Depth 100

# If an admin account existed, write output
if ($AdminExists) { Write-Output $AdminOutput | ConvertTo-Json -Depth 100 }