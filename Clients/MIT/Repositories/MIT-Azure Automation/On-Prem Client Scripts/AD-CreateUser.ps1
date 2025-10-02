<#

Mangano IT - Active Directory - Create AD User
Created by: Gabriel Nugent
Version: 1.12.8

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [int]$TicketId,
    [string]$CompanyDomain,
    [bool]$ForceADSync,
    [string]$AzureADServer,
    [Parameter(Mandatory=$true)][string]$OrganizationalUnitPath,
	[Parameter(Mandatory=$true)][string]$GivenName,
    [string]$MiddleInitial,
	[Parameter(Mandatory=$true)][string]$Surname,
    [ValidateSet('Number', 'Initial', 'Fname.Lname', 'fname.lname', 'FnameLname', 'fnamelname', 'F.Lname', 'f.lname', 'FLname', 'flname', 'N/A')]
    [string]$Suffix = 'N/A',
    [ValidateSet('Fname.Lname', 'fname.lname', 'FnameLname', 'fnamelname', 'F.Lname', 'f.lname', 'FLname', 'flname')]
    [string]$UsernameFormat = 'Fname.Lname', # Single character means initial
    [string]$UserPrincipalName,
    [string]$Username,
    [string]$DisplayName,
    [string]$MailAddress,
    [string]$MailNickname,
    [string]$Password,
    [int]$PasswordLength = 12,
    [int]$DinopassCalls = 1,
    [string]$MobilePhone,
    [string]$OfficePhone,
    [string]$FaxNumber,
    [string]$Title,
    [string]$Department,
    [string]$Company,
    [string]$StreetAddress,
    [string]$City,
    [string]$State,
    [string]$PostalCode,
    [string]$ManagerDisplayName,
    [string]$TemplateUser,
    [string]$Country = "AU",
    [boolean]$PreviouslyEmployed = $false,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

# Title of the task related to this script - only for when a ticket is involved
$TaskNotes = "Create New User"

# Output the status variable content to a file
$Date = Get-Date -Format "dd-MM-yyyy HHmm"
$FilePath = "C:\Scripts\Logs\AD-CreateUser"
$FileName = "$GivenName $Surname-$Date.txt"

$Result = $false
$UserAlreadyExists = $false

# Get task to see if user has been created by this automation already
if ($TicketId -ne 0 -and $null -ne $ApiSecrets) {
    $NewUserTask = .\CWM-FindTicketTasks.ps1 -TicketId $TicketId -TaskNotes $TaskNotes -ApiSecrets $ApiSecrets | ConvertFrom-Json
    $TaskClosed = $NewUserTask.closedFlag
} else {
    $TaskClosed = $false # Provide default answer if not attached to ticket
}

## FUNCTIONS ##

function CreateUsernames {
    param (
        [Parameter(Mandatory)][string]$GivenName,
        [string]$MiddleInitial,
        [Parameter(Mandatory)][string]$Surname,
        [Parameter(Mandatory)][string]$CompanyDomain,
        [ValidateSet('Number', 'Initial', 'N/A')][string]$Suffix = 'N/A',
        [ValidateSet('Fname.Lname', 'fname.lname', 'FnameLname', 'fnamelname', 'F.Lname', 'f.lname', 'FLname', 'flname')]
        [string]$UsernameFormat = 'Firstname.Lastname',
        [int]$SuffixNumber,
        [string]$DisplayName,
        [string]$Username,
        [string]$UserPrincipalName,
        [string]$MailAddress,
        [string]$MailNickname
    )
    
    # If DisplayName has not been provided, create one
    if ($DisplayName -eq '') {
        $DisplayName = "$GivenName $Surname"
        $Log += "INFO: Display name has been generated.`nDisplay name: $DisplayName.`n`n"
    }

    # Remove special characters from the first and last names
    $GivenName = $GivenName -replace '\W', ''
    $Surname = $Surname -replace '\W', ''

    # If Username has not been provided, create one, then format username to meet size requirements
    if ($Username -eq '') {
        $GivenInitial = $GivenName.Substring(0, 1)
        switch ($Suffix) {
            'Initial' { $Username = "$GivenName.$MiddleInitial.$Surname" }
            Default {
                switch -CaseSensitive ($UsernameFormat) {
                    'Fname.Lname' { $Username = "$GivenName.$Surname" }
                    'fname.lname' {
                        $Username = "$GivenName.$Surname"
                        $Username = $Username.ToLower();
                    }
                    'FnameLname' { $Username = "$GivenName$Surname" }
                    'fnamelname' {
                        $Username = "$GivenName$Surname"
                        $Username = $Username.ToLower();
                    }
                    'F.Lname' { $Username = "$GivenInitial.$Surname" }
                    'f.lname' {
                        $Username = "$GivenInitial.$Surname"
                        $Username = $Username.ToLower();
                    }
                    'FLname' { $Username = "$GivenInitial$Surname" }
                    'flname' {
                        $Username = "$GivenInitial$Surname"
                        $Username = $Username.ToLower();
                    }
                }
            }
        }
        if ($Suffix -eq 'Number') {
            $Username = "$Username$SuffixNumber"
        }

        # Remove unwanted characters
        $Username = $Username -replace '\s', ''
        $Username = $Username -replace "'", ''
        $Username = $Username -replace '-', ''
        $Log += "INFO: Username has been generated.`nUsername: $Username.`n"
        Write-Warning "Username has been generated: $Username"
    }

    # If UserPrincipalName has not been provided, create one
    if ($UserPrincipalName -eq '') {
        if ($CompanyDomain -like '@*') { $UserPrincipalName = "$Username$CompanyDomain" }
        else { $UserPrincipalName = "$Username@$CompanyDomain" }
        $Log += "INFO: User principal name has been generated.`nUser principal name: $UserPrincipalName.`n`n"
    }

    if ($Username.Length -gt 18) {
        if ($Suffix -eq 'Number') {
            $SubstringLength = 18 - $SuffixNumber
            $SamAccountName = $Username.Substring(0,$SubstringLength)
            $SamAccountName = "$SamAccountName$SuffixNumber"
        } else { $SamAccountName = $Username.Substring(0,18) }
        $Log += "INFO: SAM account name has been formatted to meet size requirements.`nSAM account name: $SamAccountName.`n`n"
    } else {
        $SamAccountName = $Username
        $Log += "INFO: SAM account name has been set.`nSAM account name: $SamAccountName.`n`n"
    }

    # If MailAddress has not been provided, use UPN
    if ($MailAddress -eq '') {
        $MailAddress = $UserPrincipalName
        $Log += "INFO: Mail address has been generated.`nMail address: $MailAddress.`n`n"
    }

    # If MailNickname requested to be MailAddress, set
    if ($MailNickname -eq 'Email Address') {
        $MailNickname = $MailAddress
    }

    # Create output object
    $Output = @{
        DisplayName = $DisplayName
        Username = $Username
        UserPrincipalName = $UserPrincipalName
        SamAccountName = $SamAccountName
        MailAddress = $MailAddress
        MailNickname = $MailNickname
    }

    return $Output
}

## CHECKS AND BALANCES FOR USER DETAILS ##

# Create username etc.
$UsernameParams = @{
    GivenName = $GivenName
    Surname = $Surname
    Suffix = 'N/A'
    CompanyDomain = $CompanyDomain
    DisplayName = $DisplayName
    Username = $Username
    UsernameFormat = $UsernameFormat
    UserPrincipalName = $UserPrincipalName
    MailAddress = $MailAddress
    MailNickname = $MailNickname
}

$UserDetails = CreateUsernames @UsernameParams
$DisplayName = $UserDetails.DisplayName
$Username = $UserDetails.Username
$UserPrincipalName = $UserDetails.UserPrincipalName
$SamAccountName = $UserDetails.SamAccountName
$MailAddress = $UserDetails.MailAddress
$MailNickname = $UserDetails.MailNickname

# If Password has not been provided, create one with a minimum of 12 characters and no pluses
if ($Password -eq '') {
    $Password = .\GenerateSecurePassword.ps1 -Length $PasswordLength -DinopassCalls $DinopassCalls
    $Log += "INFO: Password has been generated, and will be only provided via script output.`n`n"
}

# Convert Password to secure string
$SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force

## LOCATE MANAGER ##

if ($ManagerDisplayName -ne '') {
    $Log += "Attempting to locate manager $ManagerDisplayName...`n"
    $Manager = .\AD-CheckUserExists.ps1 -DisplayName $ManagerDisplayName
    if ($Manager.Result) {
        $Log += "SUCCESS: Manager $ManagerDisplayName located.`n`n"
        $ManagerDistinguishedName = $Manager.User.DistinguishedName
    } else {
        $Log += "ERROR: Unable to locate manager $ManagerDisplayName. No manager will be set.`n`n"
        Write-Error "Unable to locate manager $ManagerDisplayName. No manager will be set : $_"
        $ManagerDistinguishedName = $null
    }
} else {
    $Log += "INFO: No manager details provided.`n`n"
    $ManagerDistinguishedName = $null
}

## CHECK IF USER ALREADY EXISTS ##

$GetUserResponse = .\AD-CheckUserExists.ps1 -GivenName $GivenName -Surname $Surname
$UserAlreadyExists = $GetUserResponse.Result

# Reformat username if user already exists
if ($UserAlreadyExists -and !$PreviouslyEmployed -and $Suffix -ne 'N/A' -and !$TaskClosed) {
    # Create number to increment until account found
    $SuffixNumber = 0
    do {
        $SuffixNumber += 1

        # Recreate username etc.
        $NewUsernameParams = @{
            GivenName = $GivenName
            MiddleInitial = $MiddleInitial
            Surname = $Surname
            Suffix = $Suffix
            CompanyDomain = $CompanyDomain
            SuffixNumber = $SuffixNumber
            UsernameFormat = $UsernameFormat
            MailNickname = "$MailNickname$SuffixNumber"
        }

        # Replace suffix details if alt username format suggested
        if ($Suffix -like '*name') {
            $NewUsernameParams.UsernameFormat = $Suffix
        }

        $UserDetails = CreateUsernames @NewUsernameParams
        $DisplayName = $UserDetails.DisplayName
        $Username = $UserDetails.Username
        $UserPrincipalName = $UserDetails.UserPrincipalName
        $SamAccountName = $UserDetails.SamAccountName
        $MailAddress = $UserDetails.MailAddress
        $MailNickname = $UserDetails.MailNickname

        $GetUserResponse = .\AD-CheckUserExists.ps1 -SamAccountName $SamAccountName
        $UserAlreadyExists = $GetUserResponse.Result

        # Stop loop until user no longer found, or until account with matching middle initial found
        if (!$UserAlreadyExists -or ($UserAlreadyExists -and ($Suffix -eq 'Initial' -or $Suffix -like '*name'))) {
            $BreakCondition = $true
        }
    } until ($BreakCondition)
}

## CREATE USER PARAMETERS ##

$UserProperties = @{
    UserPrincipalName = $UserPrincipalName
    SAMAccountName = $SamAccountName
    Name = $DisplayName
    DisplayName = $DisplayName
    EmailAddress = $UserPrincipalName
    GivenName = $GivenName
    Surname = $Surname
    AccountPassword = $SecurePassword
    Path = $OrganizationalUnitPath
    Enabled = $true
}

# For each optional variable that isn't accounted for, add to ApiBody if it isn't null
if ($MailNickname -ne '') { $UserProperties.Add('OtherAttributes', @{'mailNickname' = $MailNickname}) }
if ($MobilePhone -ne '') { $UserProperties.Add('MobilePhone', $MobilePhone) }
if ($OfficePhone -ne '') { $UserProperties.Add('OfficePhone', $OfficePhone) }
if ($FaxNumber -ne '') { $UserProperties.Add('Fax', $FaxNumber) }
if ($Title -ne '') { $UserProperties.Add('Title', $Title) }
if ($Department -ne '') { $UserProperties.Add('Department', $Department) }
if ($Company -ne '') { $UserProperties.Add('Company', $Company) }
if ($StreetAddress -ne '') { $UserProperties.Add('StreetAddress', $StreetAddress) }
if ($City -ne '') { $UserProperties.Add('City', $City) }
if ($State -ne '') { $UserProperties.Add('State', $State) }
if ($PostalCode -ne '') { $UserProperties.Add('PostalCode', $PostalCode) }
if ($Country -ne '' -and $Country.Length -eq 2) { $UserProperties.Add('Country', $Country) }
if ($null -ne $ManagerDistinguishedName) { $UserProperties.Add('Manager', $ManagerDistinguishedName) }
if ($TicketId -ne 0) {
    $Ticket = .\CWM-FindTicketDetails.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets | ConvertFrom-Json
    $UserProperties.Add('Description', ('#' + $Ticket.id + ' - ' + $Ticket.summary))
}

## CREATE USER ##

# Skips user creation if already done
if ($TaskClosed) {
    $Log += "`nSkipping user creation, task already closed."
    Write-Warning "Skipping user creation, task already closed."
    $Result = $true
}

# Creates the user if not located
elseif (!$UserAlreadyExists) {
    try {
        $Log += "`nAttempting to create $UserPrincipalName...`n"
        New-ADUser @UserProperties | Out-Null
        $Log += "SUCCESS: $UserPrincipalName has been created."
        Write-Warning "SUCCESS: $UserPrincipalName has been created."
        $Result = $true
    } catch {
        $Log += "ERROR: $UserPrincipalName not created.`nERROR DETAILS: " + $_
        Write-Error "$UserPrincipalName not created : $_"
    }
}

# Updates user details if user was previously employed
elseif ($UserAlreadyExists -and $PreviouslyEmployed) {
    # Reconfigure user properties
    $UserProperties = @{
        Identity = $SamAccountName
        UserPrincipalName = $UserPrincipalName
        SAMAccountName = $SamAccountName
        DisplayName = $DisplayName
        EmailAddress = $UserPrincipalName
        GivenName = $GivenName
        Surname = $Surname
        Enabled = $true
    }

    # For each optional variable that isn't accounted for, add to ApiBody if it isn't null
    if ($MobilePhone -ne '') { $UserProperties.Add('MobilePhone', $MobilePhone) }
    if ($OfficePhone -ne '') { $UserProperties.Add('OfficePhone', $OfficePhone) }
    if ($FaxNumber -ne '') { $UserProperties.Add('Fax', $FaxNumber) }
    if ($Title -ne '') { $UserProperties.Add('Title', $Title) }
    if ($Department -ne '') { $UserProperties.Add('Department', $Department) }
    if ($Company -ne '') { $UserProperties.Add('Company', $Company) }
    if ($StreetAddress -ne '') { $UserProperties.Add('StreetAddress', $StreetAddress) }
    if ($City -ne '') { $UserProperties.Add('City', $City) }
    if ($State -ne '') { $UserProperties.Add('State', $State) }
    if ($PostalCode -ne '') { $UserProperties.Add('PostalCode', $PostalCode) }
    if ($Country -ne '' -and $Country.Length -eq 2) { $UserProperties.Add('Country', $Country) }
    if ($null -ne $ManagerDistinguishedName) { $UserProperties.Add('Manager', $ManagerDistinguishedName) }
    if ($TicketId -ne 0) {
        $Ticket = .\CWM-FindTicketDetails.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets | ConvertFrom-Json
        $UserProperties.Add('Description', ($GetUserResponse.User.description + "`n`n#" + $Ticket.id + ' - ' + $Ticket.summary))
    }

    try {
        $Log += "`nAttempting to update user details for $UserPrincipalName...`n"
        $User = Set-ADUser @UserProperties -PassThru
        $Log += "SUCCESS: $UserPrincipalName has been updated."
        Write-Warning "SUCCESS: $UserPrincipalName has been updated."
        $User | Move-ADObject -TargetPath $OrganizationalUnitPath
        $Log += "SUCCESS: $UserPrincipalName has been moved to $OrganizationalUnitPath."
        Write-Warning "SUCCESS: $UserPrincipalName has been moved to $OrganizationalUnitPath."
        $User | Set-ADAccountPassword -NewPassword $SecurePassword
        $Log += "SUCCESS: $UserPrincipalName has had their password updated."
        Write-Warning "SUCCESS: $UserPrincipalName has had their password updated."
        $Result = $true
    } catch {
        $Log += "ERROR: $UserPrincipalName not updated.`nERROR DETAILS: " + $_
        Write-Error "$UserPrincipalName not updated : $_"
    }
}

## ADD USER TO TEMPLATE GROUPS ##

if ($TemplateUser -ne '' -and $Result) {
    try {
        $Log += "`nAttempting to add $UserPrincipalName to template groups...`n"
        $TemplateStandard = Get-ADUser $TemplateUser -Properties memberof
        $TemplateStandard.memberof | Add-ADGroupMember -Members $SamAccountName
        $Log += "SUCCESS: $UserPrincipalName has been added to template groups."
        Write-Warning "SUCCESS: $UserPrincipalName has been added to template groups."
    }
    catch {
        $Log += "ERROR: $UserPrincipalName not added to template groups.`nERROR DETAILS: " + $_
        Write-Error "$UserPrincipalName not added to template groups : $_"
    }
}

## RUN AZURE AD SYNC IF REQUIRED ##

if ($ForceADSync -and $Result) {
    $SyncResult = .\AD-RunADSync.ps1 -AzureADServer $AzureADServer
    $Log += "`n`nINFO: ADSync $SyncResult"
    Write-Warning "AD sync result: $SyncResult"
} else { $SyncResult = $false }

## UPDATE RELATED TASK IF TICKET ID PROVIDED ##

if ($TicketId -ne 0) {
    # Collate info for task resolution
    $TaskOutput = @{
        Result = $Result
        UserAlreadyExists = $UserAlreadyExists
        SamAccountName = $SamAccountName
        UserPrincipalName = $UserPrincipalName
        DisplayName = $DisplayName
        MailAddress = $MailAddress
        SyncResult = $SyncResult
        LogFile = "$FilePath\$FileName"
    }

    if ($Result) {
        # Convert info to JSON
        [string]$TaskOutputJson = $TaskOutput | ConvertTo-Json

        # Update task with info
        $Task = .\CWM-UpdateTask.ps1 -TicketId $TicketId -Note 'Create New User' -Resolution $TaskOutputJson -ClosedStatus $true -ApiSecrets $ApiSecrets
        $Log += "`n`n" + $Task.Log

        # Set note text for when user creation works
        $NoteText = "Create New User [automated task]`n`nNew user $GivenName $Surname has been created.`n`n"
        $NoteText += "User principal name: $UserPrincipalName`nDisplay name: $DisplayName`nMail address: $MailAddress`n`nLog file: $FilePath\$FileName"

        # Add sync results if performed
        if ($ForceADSync) {
            $NoteText += "`nAD sync result: $SyncResult"
            if (!$SyncResult) { $NoteText += " - the automation will likely stop, and you will need to manually re-run the AAD sync for the account to appear." }
        }
    } else {
        # Set note text for when user creation fails
        $NoteText = "Create New User [automated task]`n`nNew user $GivenName $Surname has not been created.`nUser already exists: $UserAlreadyExists`n`n"
        $NoteText += "User principal name: $UserPrincipalName`nDisplay name: $DisplayName`nMail address: $MailAddress`n`nLog file: $FilePath\$FileName"
    }

    # Add note to ticket
    $Note = .\CWM-AddTicketNote.ps1 -TicketId $TicketId -Text $NoteText -ResolutionFlag $true -ApiSecrets $ApiSecrets
    $Log += "`n`n" + $Note.Log
}

## SEND DETAILS BACK TO FLOW ##

$Output = @{
    Result = $Result
    UserAlreadyExists = $UserAlreadyExists
    SamAccountName = $SamAccountName
    UserPrincipalName = $UserPrincipalName
    DisplayName = $DisplayName
    MailAddress = $MailAddress
    Password = $Password
    SyncResult = $SyncResult
    Log = $Log
    LogFile = "$FilePath\$FileName"
}

Write-Output $Output | ConvertTo-Json

# Makes folder for logs and outputs logs
.\CreateLogFile.ps1 -Log $Log -FilePath $FilePath -FileName $FileName | Out-Null