<#

Mangano IT - Azure Active Directory - Create AAD User
Created by: Gabriel Nugent
Version: 1.19.5

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [int]$TicketId,
    [string]$CompanyDomain,
	[string]$BearerToken,
    [string]$TenantUrl,
	[Parameter(Mandatory=$true)][string]$GivenName,
	[Parameter(Mandatory=$true)][string]$Surname,
    [string]$UserPrincipalName,
    [string]$DisplayName,
    [string]$MailAddress,
    [string]$MailNickname,
    [ValidateSet('Fname.Lname', 'fname.lname', 'FnameLname', 'fnamelname', 'F.Lname', 'f.lname', 'FLname', 'flname')]
    [string]$MailNicknameFormat = 'Fname.Lname', # Single character means initial
    [string]$Password,
    [int]$PasswordLength = 12,
    [int]$DinopassCalls = 1,
    [string]$MobilePhone,
    [string]$OfficePhone,
    [string]$JobTitle,
    [string]$Department,
    [string]$Company,
    [string]$OfficeLocation,
    [string]$StreetAddress,
    [string]$City,
    [string]$State,
    [string]$PostalCode,
    [string]$ManagerUserPrincipalName,
    [string]$Country = "Australia",
    [string]$UsageLocation = "AU",
    [boolean]$PreviouslyEmployed = $false,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

# Title of the task related to this script - only for when a ticket is involved
$TaskNotes = "Create New User"

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    $Log += "Bearer token not supplied. Getting bearer token for $TenantUrl...`n`n"
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

# Fetch CWM credentials if not provided
if ($null -eq $ApiSecrets -and $TicketId -ne 0) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

# Arguments for user creation
$ApiArguments = @{
    Uri = "https://graph.microsoft.com/v1.0/users"
    Method = 'POST'
    Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
    Body = $null
    UseBasicParsing = $true
}

$Result = $false
$UserAlreadyExists = $false

# Form for submitting replacement values
$ProvideInfoForm = 'https://forms.office.com/r/GEKBcczPT3'

# Get task to see if user has been created by this automation already
if ($TicketId -ne 0) {
    $NewUserTask = .\CWM-FindTicketTasks.ps1 -TicketId $TicketId -TaskNotes $TaskNotes -ApiSecrets $ApiSecrets | ConvertFrom-Json
    $TaskClosed = $NewUserTask.closedFlag
} else {
    $TaskClosed = $false # Provide default answer if not attached to ticket
}

## CHECKS AND BALANCES FOR USER DETAILS ##

# Remove spaces from MailAddress, MailNickname and UserPrincipalName 
$MailAddress = $MailAddress -replace '\s', ''
$MailNickname = $MailNickname -replace '\s', ''
$UserPrincipalName = $UserPrincipalName -replace '\s', ''

# Print mobile number in Azure Automation console
if ($MobilePhone -ne '') {
    Write-Warning "Mobile number: $($MobilePhone)"
}

# If DisplayName has not been provided, create one
if ($DisplayName -eq '') {
    $DisplayName = "$GivenName $Surname"
    $Log += "INFO: Display name has been generated.`nDisplay name: $DisplayName.`n`n"
}

# If MailNickname has not been provided, create one
if ($MailNickname -eq '') {
    # Remove unwanted characters
    $GivenNameFixed = $GivenName -replace "\W", ""
    $SurnameFixed = $Surname -replace "\W", ""
    $GivenInitial = $GivenNameFixed.Substring(0, 1)

    # Create mail nickname based on requested format
    switch -CaseSensitive ($MailNicknameFormat) {
        'Fname.Lname' { $MailNickname = "$GivenNameFixed.$SurnameFixed" }
        'fname.lname' {
            $MailNickname = "$GivenNameFixed.$SurnameFixed"
            $MailNickname = $MailNickname.ToLower();
        }
        'FnameLname' { $MailNickname = "$GivenNameFixed$SurnameFixed" }
        'fnamelname' {
            $MailNickname = "$GivenName$SurnameFixed"
            $MailNickname = $MailNickname.ToLower();
        }
        'F.Lname' { $MailNickname = "$GivenInitial.$SurnameFixed" }
        'f.lname' {
            $MailNickname = "$GivenInitial.$SurnameFixed"
            $MailNickname = $MailNickname.ToLower();
        }
        'FLname' { $MailNickname = "$GivenInitial$SurnameFixed" }
        'flname' {
            $MailNickname = "$GivenInitial$SurnameFixed"
            $MailNickname = $MailNickname.ToLower();
        }
        default { $MailNickname = "$GivenNameFixed.$SurnameFixed" }
    }

    $Log += "INFO: Mail nickname has been generated.`nMail nickname: $MailNickname.`n`n"
}

# If UserPrincipalName has not been provided, create one
if ($UserPrincipalName -eq '') {
    if ($CompanyDomain -like '@*') { $UserPrincipalName = $MailNickname + $CompanyDomain }
    else { $UserPrincipalName = "$MailNickname@$CompanyDomain" }
    $Log += "INFO: User principal name has been generated.`nUser principal name: $UserPrincipalName.`n`n"
}

# If MailAddress has not been provided, use UPN
if ($MailAddress -eq '') {
    $MailAddress = $UserPrincipalName
    $Log += "INFO: Mail address has been generated.`nMail address: $MailAddress.`n`n"
}

# If Password has not been provided, create one with a minimum of 12 characters and no pluses
if ($Password -eq '') {
    $Password = .\GenerateSecurePassword.ps1 -Length $PasswordLength -DinopassCalls $DinopassCalls
    $Log += "INFO: Password has been generated, and will be only provided via script output.`n`n"
    Write-Warning "Password: $Password"
}

## CREATE REQUEST BODY ##

$ApiBody = @{
    userPrincipalName = $UserPrincipalName
    displayName = $DisplayName
    mail = $MailAddress
    mailNickname = $MailNickname
    givenName = $GivenName
    surname = $Surname
    usageLocation = $UsageLocation
    passwordProfile = @{
    	forceChangePasswordNextSignIn = $true
    	forceChangePasswordNextSignInWithMfa = $true
    	password = $Password
    }
    accountEnabled = $true
}

# For each optional variable that isn't accounted for, add to ApiBody if it isn't null
if ($MobilePhone -ne '') { $ApiBody.Add('mobilePhone', $MobilePhone) }
if ($OfficePhone -ne '') { $ApiBody.Add('officePhone', $OfficePhone) }
if ($JobTitle -ne '') { $ApiBody.Add('jobTitle', $JobTitle) }
if ($Department -ne '') { $ApiBody.Add('department', $Department) }
if ($Company -ne '') { $ApiBody.Add('companyName', $Company) }
if ($OfficeLocation -ne '') { $ApiBody.Add('officeLocation', $OfficeLocation) }
if ($StreetAddress -ne '') { $ApiBody.Add('streetAddress', $StreetAddress) }
if ($City -ne '') { $ApiBody.Add('city', $City) }
if ($State -ne '') { $ApiBody.Add('state', $State) }
if ($PostalCode -ne '') { $ApiBody.Add('postalCode', $PostalCode) }
if ($Country -ne '') { $ApiBody.Add('country', $Country) }

## CHECK IF USER ALREADY EXISTS ##

$GetUserArguments = @{
    Uri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName" + '?$select=id,displayName,userPrincipalName,mail,mailNickname,accountEnabled'
    Method = 'GET'
    Headers = @{ 'Content-Type'="application/json";'Authorization'="Bearer $BearerToken" }
    UseBasicParsing = $true
}

try {
    $Log += "Checking to see if $UserPrincipalName already exists...`n"
    $GetUserResponse = Invoke-WebRequest @GetUserArguments | ConvertFrom-Json
    $UserAlreadyExists = $true

    # Stash vars that the flow provided
    $UserPrincipalName_Old = $UserPrincipalName
    $MailAddress_Old = $MailAddress
    $MailNickname_Old = $MailNickname
    $DisplayName_Old = $DisplayName

    # Match vars to what exists on the account
    $UserId = $GetUserResponse.id
    $UserPrincipalName = $GetUserResponse.userPrincipalName
    $DisplayName = $GetUserResponse.displayName
    $MailAddress = $GetUserResponse.mail
    $MailNickname = $GetUserResponse.mailNickname
    $AccountEnabled = $GetUserResponse.accountEnabled
    $Log += "INFO: $UserPrincipalName has been located.`n"
    if ($PreviouslyEmployed) {
        $ApiArguments.Method = 'PATCH'
        $ApiArguments.Uri += "/$UserPrincipalName"

        # Remove password from API body - PATCH cannot set password
        $ApiBody.Remove('passwordProfile')
        $Password = ''

        $Log += "The user has been previously employed, so their details will be updated.`n"
        Write-Warning "INFO: The user has been previously employed, so their details will be updated."
    }
    else {
        if ($TaskClosed) {
            $Log += "Automation for this ticket has already created this user!`n"
            Write-Warning "Automation for this ticket has already created this user!"
            $Result = $true
        } else {
            $Log += "The user has not been marked as previously employed, so their account will not be created.`nPlease supply a new user principal name.`n"
            Write-Error "The user has not been marked as previously employed, so their account will not be created.`nPlease supply a new user principal name."
        }
    }
} catch {
    $GetUserResponse = $null
    $Log += "SUCCESS: $UserPrincipalName has not been located.`n"
    Write-Warning "SUCCESS: $UserPrincipalName has not been located."
}

# Convert ApiBody to JSON
$ApiArguments.Body = $ApiBody | ConvertTo-Json -Depth 100

## CREATE USER ##

# Contacts Graph API to create the new user account - basic parsing required for Azure Runbook
if ($TaskClosed) {
    $Log += "`nSkipping user creation, task already closed."
    Write-Warning "Skipping user creation, task already closed."
    $Result = $true
    $AccountEnabled = $true
}
elseif ($null -eq $GetUserResponse -or $PreviouslyEmployed) {
    try {
        $Log += "`nAttempting to create/update $UserPrincipalName...`n"
        $ApiResponse = Invoke-WebRequest @ApiArguments | ConvertFrom-Json
        $Log += "SUCCESS: $UserPrincipalName has been created/updated."
        Write-Warning "SUCCESS: $UserPrincipalName has been created/updated."
        $Result = $true
        $AccountEnabled = $true
        if (!$UserAlreadyExists) {
            $UserId = $ApiResponse.id
            $Log += "`nINFO: User ID - $UserId"
        }
    } catch {
        $Log += "ERROR: $UserPrincipalName not created.`nERROR DETAILS: " + $_
        Write-Error "$UserPrincipalName not created : $_"
        $ApiResponse.id = $null
        $UserId = 'null'
        $AccountEnabled = $false
    }
}

## UPDATE MANAGER ##

if ($null -ne $UserId -and $ManagerUserPrincipalName -ne '') {
    # Get manager details
    $GetManagerArguments = @{
        Uri = "https://graph.microsoft.com/v1.0/users/$($ManagerUserPrincipalName)"
        Method = 'GET'
        Headers = @{ 'Content-Type'="application/json";'Authorization'="Bearer $BearerToken" }
        UseBasicParsing = $true
    }

    $Manager = Invoke-WebRequest @GetManagerArguments

    # Update manager if found
    if ($null -ne $Manager) {
        $SetManagerArguments = @{
            Uri = "https://graph.microsoft.com/v1.0/users/$($UserId)"
            Method = 'PUT'
            Headers = @{ 'Content-Type'="application/json";'Authorization'="Bearer $BearerToken" }
            UseBasicParsing = $true
            Body = @{
                '@odata.id' = "https://graph.microsoft.com/v1.0/users/$($Manager.id)"
            } | ConvertTo-Json -Depth 100
        }

        try {
            Invoke-WebRequest @SetManagerArguments | Out-Null
            $Log += "`n`nSUCCESS: $($ManagerUserPrincipalName) has been set as the manager for $($UserPrincipalName)."
            Write-Warning "SUCCESS: $($ManagerUserPrincipalName) has been set as the manager for $($UserPrincipalName)."
        } catch {
            $Log += "ERROR: $($ManagerUserPrincipalName) has not been set as the manager for $($UserPrincipalName).`nERROR DETAILS: $($_)"
            Write-Error "$($ManagerUserPrincipalName) has not been set as the manager for $($UserPrincipalName) : $($_)"
        }
    }
}

## UPDATE RELATED TASK IF TICKET ID PROVIDED ##

if ($TicketId -ne 0 -and $UserId -ne 'null') {
    # Collate info for task resolution
    $TaskOutput = @{
        Result = $Result
        UserAlreadyExists = $UserAlreadyExists
        UserId = $UserId
        UserPrincipalName = $UserPrincipalName
        DisplayName = $DisplayName
        MailAddress = $MailAddress
        MailNickname = $MailNickname
    }

    # Convert info to JSON
    [string]$TaskOutputJson = $TaskOutput | ConvertTo-Json

    if ($Result) {
        # Update task only if it isn't already closed
        if (!$TaskClosed) {
            # Convert info to JSON
            [string]$TaskOutputJson = $TaskOutput | ConvertTo-Json

            # Update task with info
            $Task = .\CWM-UpdateTask.ps1 -TicketId $TicketId -Note 'Create New User' -Resolution $TaskOutputJson -ClosedStatus $true `
            -ApiSecrets $ApiSecrets | ConvertFrom-Json
            $Log += "`n`n" + $Task.Log
        }

        # Set note text for when user creation works
        $NoteText = "$TaskNotes [automated task]`n`nNew user $GivenName $Surname has been created.`n`nAccount enabled: $AccountEnabled`n"
        $NoteText += "User principal name: $UserPrincipalName`nUser ID: $UserId`nDisplay name: $DisplayName`nMail address: $MailAddress"
    } else {
        # Set note text for when user creation fails
        $NoteText = "$TaskNotes [automated task]`n`nNew user $GivenName $Surname has not been created.`nUser already exists: $UserAlreadyExists`n`n"
        $NoteText += "Account enabled: $AccountEnabled`nUser principal name: $UserPrincipalName`nUser ID: $UserId`n"
        $NoteText += "Display name: $DisplayName`nMail address: $MailAddress"

        # Append request for info if user exists
        if ($UserAlreadyExists -and $DisplayName -ne $DisplayName_Old) {
            $NoteText += "`n`nAs the user already exists, you will need to supply replacement values for the following attributes:`n"
            if ($UserPrincipalName -eq $UserPrincipalName_Old) { $NoteText += "`n- User principal name" }
            if ($MailAddress -eq $MailAddress_Old) { $NoteText += "`n- Mail address" }
            if ($MailNickname -eq $MailNickname_Old) { $NoteText += "`n- Mail nickname" }
            $NoteText += "`n`nPlease reach out to the client to agree upon attributes for the new user.`n`n"
            $NoteText += "Once this is done, please submit the new attributes to this form: $ProvideInfoForm"
        }
        
        # Append request for info if user exists, is disabled, and has a matching display name
        elseif (!$AccountEnabled -and $DisplayName -eq $DisplayName_Old) {
            $NoteText += "`n`nBased on the display name of the existing user, it's likely that the current request is for a "
            $NoteText += "person who was previously employed by Seasons Living."
            $NoteText += "`n`nPlease reach out to the client to confirm that this is the case.`n`n"
            $NoteText += "Once this is done, please submit the new attributes to this form: $ProvideInfoForm"
        }
    }

    # Add note to ticket
    $Note = .\CWM-AddTicketNote.ps1 -TicketId $TicketId -Text $NoteText -ResolutionFlag $true -ApiSecrets $ApiSecrets
    $Log += "`n`n" + $Note.Log
} 

## SEND DETAILS BACK TO FLOW ##

$Output = @{
    Result = $Result
    UserAlreadyExists = $UserAlreadyExists
    AccountEnabled = $AccountEnabled
    UserId = $UserId
    UserPrincipalName = $UserPrincipalName
    DisplayName = $DisplayName
    MailAddress = $MailAddress
    MailNickname = $MailNickname
    Password = $Password
    Log = $Log
}

Write-Output $Output | ConvertTo-Json