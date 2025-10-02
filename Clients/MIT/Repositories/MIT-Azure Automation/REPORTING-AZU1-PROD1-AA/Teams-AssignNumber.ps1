<#

Mangano IT - Teams - Set Up Calling
Created by: Liam Adair, Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory)][string]$NumberLookup,
    [Parameter(Mandatory)][string]$UserPrincipalName,
    [Parameter(Mandatory)][string]$NumberType,
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][int]$TicketId,
    [string]$BearerToken,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$MatchingPhoneNumbers = @()

## GET CREDENTIALS ##

# Get CW Manage credentials if not provided

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

# Fetch tenant slug
$TenantSlug = .\CWM-FindCompanySlug.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets

# Grab connection variables
$ApplicationId = .\KeyVault-GetSecret.ps1 -SecretName "$TenantSlug-TMS-ApplicationID"
$CertificateThumbprint = .\KeyVault-GetSecret.ps1 -SecretName "$TenantSlug-TMS-CertificateThumbprint"

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantId..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantId
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

## CONNECT TO TEAMS ##
    
try {
    Connect-MicrosoftTeams -CertificateThumbprint $CertificateThumbprint -ApplicationId $ApplicationId -TenantId $TenantId | Out-Null
    Write-Warning "SUCCESS: Connected to Teams Online."
} catch {
    Write-Error "Failed to connect to Teams Online : $($_)"
}

## GRAB USER DETAILS ##

# Build request variables
$GetUserArguments = @{
    Uri = "https://graph.microsoft.com/v1.0/users/$($UserPrincipalName)"
    Method = 'GET'
    Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
    UseBasicParsing = $true
}

# Get user from AAD
try {
    $User = Invoke-WebRequest @GetUserArguments | ConvertFrom-Json
    Write-Warning "SUCCESS: Grabbed user details for $($UserPrincipalName)."
} catch {
    Write-Error "Unable to grab user details for $($UserPrincipalName) : $($_)"
    exit
}

## SEARCH FOR AND ASSIGN NUMBERS ##

# Lookup user phone numbers
Write-Warning "Looking up '$($NumberType)' numbers like $($NumberLookup) for $($UserPrincipalName)..."
$PhoneNumbers = Get-CsPhoneNumberAssignment -ActivationState Activated -CapabilitiesContain UserAssignment 

# Check if user already has a Teams number assigned
Write-Warning "Checking to see if $($UserPrincipalName) already has any numbers assigned..."
$AssignedPhoneNumbers = $PhoneNumbers | Sort-Object TelephoneNumber | Where-Object { $_.AssignedPstnTargetId -ne $null }
foreach ($Number in $AssignedPhoneNumbers) {
    if ($Number.AssignedPstnTargetId -eq $User.id) { 
        Write-Warning "$($UserPrincipalName) has already been assigned the number $($Number.TelephoneNumber). Script will now exit."

        ## ADD A TICKET NOTE HERE ##
        $Text = "$($UserPrincipalName) has already been assigned the number $($Number.TelephoneNumber)."
        .\CWM-AddTicketNote.ps1 -TicketId $TicketId -Text $Text -ResolutionFlag $true -ApiSecrets $ApiSecrets | Out-Null

        Exit
    }
}

# Sort through available phone numbers
Write-Warning "Checking all available phone numbers..."
$AvailablePhoneNumbers = $PhoneNumbers | Sort-Object TelephoneNumber | Where-Object { $_.AssignedPstnTargetId -eq $null }
if ($AvailablePhoneNumbers.Count -eq 0) {
    Write-Warning "No available phone numbers. Script will now exit."

    ## ADD A TICKET NOTE HERE ##

    Exit
} else {
    Write-Warning "$($AvailablePhoneNumbers.Count) available phone number/s have been found."
}

# Find a phone number that matches the provided prefix
foreach ($AvailablePhoneNumber in $AvailablePhoneNumbers){
    if ($AvailablePhoneNumber.TelephoneNumber -like $NumberLookup)
    {
        $MatchingPhoneNumbers += $AvailablePhoneNumber.TelephoneNumber
    }
}

# Grab the first phone number from the list (if there are any)
$NumberToAssign = $MatchingPhoneNumbers[0]
if (($MatchingPhoneNumbers.Count -le 0) -or ($null -eq $NumberToAssign)) {
    Write-Warning "There are no numbers available that match the required prefix. Script will now exit."

    ## ADD A TICKET NOTE HERE ##

    Exit 
} else {
    Write-Warning "$($MatchingPhoneNumbers.Count) matching phone number/s found."
    Write-Warning "Next available phone number: $($NumberToAssign)"
}

## ASSIGN PHONE NUMBER TO USER ##

Write-Warning "Attempting to assign phone number $($NumberToAssign) to $($UserPrincipalName)..."
try {
    Set-CsPhoneNumberAssignment -Identity $UserPrincipalName -PhoneNumber $NumberToAssign -PhoneNumberType $NumberType -ErrorAction Stop | Out-Null
} catch {
    Write-Error "Unable to assign phone number $($NumberToAssign) to $($UserPrincipalName) : $($_)"

    ## ADD A TICKET NOTE HERE ##

    Exit
}

# Ensure number is assigned
$ValidateAssignment = Get-CsPhoneNumberAssignment -TelephoneNumber $NumberToAssign

# Grab user details again
$ValidateUserArguments = @{
    Uri = "https://graph.microsoft.com/v1.0/users/$($ValidateAssignment.AssignedPstnTargetId)"
    Method = 'GET'
    Headers = @{'Content-Type'="application/json";'Authorization'="Bearer $BearerToken"}
    UseBasicParsing = $true
}

# Get user from AAD
try {
    $ValidateUser = Invoke-WebRequest @ValidateUserArguments | ConvertFrom-Json
    Write-Warning "SUCCESS: Grabbed user details for $($ValidateAssignment.AssignedPstnTargetId)."
} catch {
    Write-Error "Unable to grab user details for $($ValidateAssignment.AssignedPstnTargetId) : $($_)"
    exit
}

# Check to see if fetched user matches requested user
if ($UserPrincipalName -eq $ValidateUser.UserPrincipalName){
    Write-Warning "SUCCESS: $($NumberToAssign) has been assigned to $($UserPrincipalName)."

    ## ADD TICKET NOTE HERE ##

} else {
    Write-Error "$($NumberToAssign) has not been assigned to $($UserPrincipalName)."

    ## ADD TICKET NOTE HERE ##
}