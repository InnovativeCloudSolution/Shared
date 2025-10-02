## AZURE AD ##

# Function for adding user to a group
function Add-UserToSecGroupAAD {
    param (
        [string]$UserObjectId = '',
        [string]$Email = '',
        [Parameter(Mandatory=$true)][string]$GroupName = ''
    )

    # Fetches the user object ID if it's not provided
    if ($Email -ne '' -and $UserObjectId -eq '') { 
        $User = Get-AzureADUser -ObjectId $Email
        $UserObjectId = $User.ObjectId 
    }

    # Fetches the user's email if it's not provided
    if ($Email -eq '') {
        $User = Get-AzureADUser -ObjectId $UserObjectId
        $Email = $User.userPrincipalName
    }

    $Group = Get-AzureADGroup -SearchString $GroupName
    try {
        Add-AzureADGroupMember -ObjectId $Group.ObjectId -RefObjectId $UserObjectId
        $Log += "`n$Email has been added to $GroupName."
    }
    catch { 
        $Log += "`nERROR: $Email has not been added to $GroupName." 
    }

    Write-Output $Log
}

# Add user to multiple groups
function Add-UserToSecGroupsAAD {
    param (
        [string]$UserObjectId = '',
        [string]$Email = '',
        $Groups = @()
    )

    $Log = "`n"

    # Fetches the user object ID if it's not provided
    if ($Email -ne '') { $User = Get-AzureADUser -ObjectId $Email }

    if ($null -eq $Groups) { $Log += "`nNo security groups to add in the given list.`n"}
    else { foreach ($GroupName in $Groups) { $Log += Add-UserToSecGroupAAD -UserObjectId $User.ObjectId -GroupName $GroupName -Email $Email }}

    Write-Output $Log
}

# Create account
function Add-UserAccountAAD {
    param (
        [string]$FirstName = '',
        [string]$LastName = '',
        [Parameter(Mandatory=$true)][string]$DisplayName = '',
        [Parameter(Mandatory=$true)][string]$MailNickname = '',
        [string]$JobTitle = '',
        [Parameter(Mandatory=$true)][string]$Email = '',
        [string]$Mobile = '',
        [string]$TelephoneNumber = '',
        [string]$City = '',
        [string]$StreetAddress = '',
        [string]$PostalCode = '',
        [string]$State = '',
        [string]$Country = '',
        [string]$UsageLocation = '',
        [string]$Manager = '',
        [Parameter(Mandatory=$true)][string]$Password = ''
    )

    $Log = "`n"

    # Makes the password profile - required for Azure AD module
    $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
    $PasswordProfile.Password = $Password
    $PasswordProfile.ForceChangePasswordNextLogin = $false

    try {
        $Log += "`nCreating the new user"
        New-AzureADUser -DisplayName $DisplayName -PasswordProfile $PasswordProfile -UserPrincipalName $Email `
        -AccountEnabled $true -MailNickname $MailNickname
        $Log += "`nNew user account created with username: $Email `n"
    }
    catch { $Log += "`nERROR: Unable to create $Email.`n" }

    # Grabs the user from Azure AD to get the object ID and to check that it exists
    try { 
        $UserAccount = Get-AzureADUser -ObjectId $Email

        # Sets new user variables based on parameters given
        if ($FirstName -ne '') { Set-AzureADUser -ObjectId $UserAccount.ObjectId -GivenName $FirstName }
        if ($LastName -ne '') { Set-AzureADUser -ObjectId $UserAccount.ObjectId -Surname $LastName }
        if ($JobTitle -ne '') { Set-AzureADUser -ObjectId $UserAccount.ObjectId -JobTitle $JobTitle }
        if ($Mobile -ne '') { Set-AzureADUser -ObjectId $UserAccount.ObjectId -Mobile $Mobile }
        if ($TelephoneNumber -ne '') { Set-AzureADUser -ObjectId $UserAccount.ObjectId -TelephoneNumber $TelephoneNumber }
        if ($City -ne '') { Set-AzureADUser -ObjectId $UserAccount.ObjectId -City $City }
        if ($StreetAddress -ne '') { Set-AzureADUser -ObjectId $UserAccount.ObjectId -StreetAddress $StreetAddress }
        if ($PostalCode -ne '') { Set-AzureADUser -ObjectId $UserAccount.ObjectId -PostalCode $PostalCode }
        if ($State -ne '') { Set-AzureADUser -ObjectId $UserAccount.ObjectId -State $State }
        if ($Country -ne '') { Set-AzureADUser -ObjectId $UserAccount.ObjectId -Country $Country }
        if ($UsageLocation -ne '') { Set-AzureADUser -ObjectId $UserAccount.ObjectId -UsageLocation $UsageLocation }
        if ($Manager -ne '') { 
            $ManagerObject = Get-AzureADUser -SearchString $Manager
            Set-AzureADUserManager -ObjectId $UserAccount.ObjectId -RefObjectId $ManagerObject.ObjectId
        }
    }
    catch { 
        $Log += "`nERROR: $Email does not exist.`n" 
    }

    Write-Output $Log
}

## EXCHANGE ##

# Function to add a user to a distribution group
function Add-UserToDistGroup {
    param (
        [string]$Email = '',
        [string]$GroupName = ''
    )

    $Log = ''

    $DistGroup = Get-DistributionGroup -Identity $GroupName
    $Mailbox = Get-Mailbox -Identity $Email

    try {
        Add-DistributionGroupMember -Identity $DistGroup.DistinguishedName -Member $Mailbox.DistinguishedName
        $Log += "`nThe user $Email has been added to the $GroupName group."
    }
    catch { $Log += "`nERROR: $Email has not been added to $GroupName." }

    Write-Output $Log
}

# Function to add user to multiple distribution groups contained in an array
function Add-UserToDistGroups {
    param (
        [string]$Email = '',
        $Groups = @()
    )

    $Log = "`n"

    if ($null -eq $Groups) { $Log += "`nNo distribution groups to add in the given list.`n"}
    else { foreach ($GroupName in $Groups) { $Log += Add-UserToDistGroup -Email $Email -GroupName $GroupName }}

    Write-Output $Log
}

# Function to grant a user permissions to access a shared mailbox
function Add-UserToSharedMailbox {
    param (
        [string]$Email = '',
        [string]$GroupName = ''
    )

    $Log = ''

    try {
        Add-MailboxPermission -Identity $GroupName -User $Email -AccessRights FullAccess
        $Log += "`nThe user $Email has been added to the $GroupName group."
    }
    catch { $Log += "`nERROR: $Email has not been added to $GroupName." }

    Write-Output $Log
}

# Function to add user to multiple shared mailboxes contained in an array
function Add-UserToSharedMailboxes {
    param (
        [string]$Email = '',
        $Groups = @()
    )

    $Log = "`n"

    if ($null -eq $Groups) { $Log += "`nNo shared mailboxes to add in the given list.`n"}
    else { foreach ($GroupName in $Groups) { $Log += Add-UserToSharedMailbox -Email $Email -GroupName $GroupName }}

    Write-Output $Log
}

function Optimize-CalendarPermissions {
    param (
        [string]$Email = ''
    )

    $Log = "`n"

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

    $Log += "`nCalendar permissions have been updated for $Email`n"

    Write-Output $Log
}

function Set-AutoAcceptCalendarInvites {
    param (
        [string]$Email = ''
    )

    $Log = "`n"

    # Sets the calendar permissions to automatically accept invites
    try {
        Set-CalendarProcessing -Identity $Email -AutomateProcessing AutoAccept -AllowConflicts $false
        $Log += "`n$Email will now automatically accept calendar invites.`n"
    }
    catch {
        $Log += "`nERROR: $Email will not automatically accept calendar invites.`n"
    }

    Write-Output $Log
}

## TEAMS ##

# Function to add a user to a Teams channel
function Add-UserToChannel {
    param (
        [string]$Email = '',
        [string]$ChannelName = ''
    )

    $Log = ''

    try {
        $Channel = Get-Team -DisplayName $ChannelName
        Add-TeamUser -GroupId $Channel.GroupId -User $Email
        $Log += "`nThe user $Email has been added to $ChannelName."
    }
    catch { $Log += "`nERROR: $Email has not been added to $ChannelName." }

    Write-Output $Log
}

# Function to add user to all channels in an array
function Add-UserToChannels {
    param (
        [string]$Email = '',
        $Channels = @()
    )

    $Log = "`n"

    if ($null -eq $Channels) { $Log += "`nNo Teams channels to add in the given list.`n"}
    else { foreach ($ChannelName in $Channels) { $Log += Add-UserToChannel -Email $Email -ChannelName $ChannelName } }

    Write-Output $Log
}

## ConnectWise Manage ##
function Connect-CWManage {
    # Set CRM variables, connect to server
    $Server = 'https://api-aus.myconnectwise.net'
    $Company = 'mit'
    $pubkey = 'v8BGY94cBIbYUnPe '
    $privatekey = 'JBY3UBAzPtUgj0Af'
    $clientId = '1208536d-40b8-4fc0-8bf3-b4955dd9d3b7'

    # Create a credential object
    $Connection = @{
        Server = $Server
        Company = $Company
        pubkey = $pubkey
        privatekey = $privatekey
        clientId = $clientId
    }

    # Connect to Manage server
    Connect-CWM @Connection
}

function Add-CWMContact {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FirstName,
        [Parameter(Mandatory=$true)]
        [string]$LastName,
        [string]$Title,
        [Parameter(Mandatory=$true)]
        [string]$Email,
        [Parameter(Mandatory=$true)]
        [string]$Domain,
        [string]$MobilePhone,
        [string]$OfficePhone,
        [Parameter(Mandatory=$true)]
        [string]$CompanyName,
        [string]$SiteId
    )

    $Log=""

    #Phone Number String Validation
    if(($MobilePhone[0] -eq "4") -and ($MobilePhone.Length -eq 9)){$MobilePhone="0"+$MobilePhone}
    if(($OfficePhone[0] -eq "7") -and ($OfficePhone.Length -eq 9)){$OfficePhone="0"+$OfficePhone}

    #### CREATING CONTACT IN CONNECTWISE MANAGE ####
    Connect-CWManage
    
    # Get the company (id)
    $Company = Get-CWMCompany -Condition "name LIKE '$CompanyName' AND deletedFlag=false"

    $comArray = @()
    $EmailComm = @{
        type=@{id=1;name='Email'}
        value=$Email
        domain=$Domain
        defaultFlag=$True;
        communicationType='Email'
    }
    $comArray += $EmailComm

    #If we are given an office phone number, add that to the ConnectWise contact
    if(($OfficePhone -ne "")-and($null -ne $OfficePhone)-and($OfficePhone -ne "null")-and($OfficePhone -ne 0)){
        $Direct = @{
            type=@{id=2;name='Direct'}
            value=$OfficePhone
            defaultFlag=$False;
            communicationType='Phone'
        }
        $comArray += $Direct
    }

    #If we are given a mobile phone number, add that to the ConnectWise contact
    if(($MobilePhone -ne "")-and($null -ne $MobilePhone)-and($MobilePhone -ne "null")-and($MobilePhone -ne 0)){
        $Mobile = @{
            type=@{id=4;name='Mobile'}
            value=$MobilePhone
            defaultFlag=$True;
            communicationType='Phone'
        }
        $comArray += $Mobile      
    }

    $CompanyId = $Company.id

    #Set the users title to the one provided. If none is given, set it to " " (space) as the CWM script does not allow empty strings to be given
    $Log += "`nCreating CWM Contact for $FirstName $LastName - $Title"
    if(($Title -eq $null) -or ($Title -eq "")){$Title=" "}
    try{
        New-CWMContact -firstName $FirstName -lastName $LastName -title $Title -company @{id=$CompanyId} -site @{id=$SiteID} -communicationItems $comArray
        $Log += "`nCreated CWM Contact for $FirstName $LastName`n"
    }
    catch{
        $Log += "`nFailed to create CWM Contact for $FirstName $LastName`n"
    }
    Disconnect-CWM

    Write-Output $Log
}

function Add-CWMMember {
    [Parameter(Mandatory=$true)]
        [string]$FirstName,
        [Parameter(Mandatory=$true)]
        [string]$LastName,
        [string]$Title,
        [Parameter(Mandatory=$true)]
        [string]$Email,
        [Parameter(Mandatory=$true)]
        [string]$Domain,
        [string]$MobilePhone,
        [string]$OfficePhone

        # Initialise log variable
        $Log = ''

        # Initialise variables for new member
        $MemberUsername = $FirstName + $LastName

        # Generates a password that's at least 30 characters long - not meant to be logged
        $Log+="`nGenerating junk password through DinoPass`n"
        $Password = ''
        while ($Password.Length -lt 30) { $Password += Invoke-restmethod -uri "https://www.dinopass.com/password/strong" }

        $NewMember = @{
            identifier = $MemberUsername
            firstName = $FirstName
            lastName = $LastName
            password = $Password
            title = $Title
            licenseClass = 'F'
            agreementInvoicingDisplayOptions = 'RemainOnInvoicingScreen'
            companyActivityTabFormat = 'SummaryList'
            defaultEmail = 'Office'
            defaultPhone = 'Office'
            invoiceScreenDefaultTabFormat = 'ShowInvoicingTab'
            invoiceTimeTabFormat = 'SummaryList'
            invoicingDisplayOptions = 'RemainOnInvoicingScreen'
            hireDate = $(ConvertTo-CWMTime (Get-Date).AddDays(-1) -Raw)
            # additional required fields
            timeZone = @{id=1}
            securityRole = @{id=63}
            structureLevel = @{id=1}
            securityLocation = @{id=38}
            defaultLocation = @{id=2}
            defaultDepartment = @{id=10}
            workRole = @{id=21}
            timeApprover = @{id=186}
            expenseApprover = @{id=186}
            salesDefaultLocation = @{id=38}
            officeEmail = $Email
            officePhone = $OfficePhone
        }

        $Log += "`nCreating CWM Member for $FirstName $LastName - $Title"
        if(($Title -eq $null) -or ($Title -eq "")){$Title=" "}
        try{
            New-CWMMember @NewMember
            $Log += "`nCreated CWM Member for $FirstName $LastName`n"
        }
        catch{
            $Log += "`nFailed to create CWM Contact for $FirstName $LastName`n"
        }
}


## MISC ##
# Exports all of the functions in the module
Export-ModuleMember -Function *