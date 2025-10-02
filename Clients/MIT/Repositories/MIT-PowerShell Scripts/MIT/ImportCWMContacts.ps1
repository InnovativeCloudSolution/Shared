param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath
)
Clear-Host
Import-Module 'ConnectWiseManageAPI'

## Program details - please remember to update the version number when making changes
Write-Host "`nMangano IT - Bulk Import Contacts to ConnectWise Manage" -ForegroundColor Yellow
Write-Host "Version: " -ForegroundColor yellow -NoNewLine; Write-Host "1.0"
Write-Host "Created by: " -ForegroundColor yellow -NoNewLine; Write-Host "Gabriel Nugent"

# Declare default CSV headers
$CsvHeaders = 'FirstName', 'LastName', 'Email', 'JobTitle', 'DirectPhone', 'MobilePhone', 'SiteId', 'CompanyName'

#### CREATING USER IN CONNECTWISE MANAGE ####
# Set CRM variables, connect to server
$CWMServer = 'https://api-aus.myconnectwise.net'
$CWMCompany = 'mit'
$CWMPublicKey = 'mBemoCno7IwHElgT'
$CWMPrivateKey = 'BmC0AR9dN7eJFouT'
$CWMClientId = '1208536d-40b8-4fc0-8bf3-b4955dd9d3b7'

# Create a credential object
$Connection = @{
    Server = $CWMServer
    Company = $CWMCompany
    pubkey = $CWMPublicKey
    privatekey = $CWMPrivateKey
    clientId = $CWMClientId
}

# Connect to Manage server
try {
    Write-Host "`nConnecting to the ConnectWise Manage server..."
    Connect-CWM @Connection
    Write-Host "Connection successful."
}
catch {
    Write-Host "Unable to connect to ConnectWise Manage. Stopping script." -ForegroundColor red
    Exit-PSSession
}

# Import CSV from specified path
try {  
    Write-Host "`nAttempting to import the CSV from " -NoNewLine; Write-Host "$CsvPath " -ForegroundColor blue -NoNewline; Write-Host "now, please wait..."
    $ContactsCsv = Import-Csv $CsvPath -Delimiter ',' -Header $CsvHeaders
    Write-Host "CSV imported successfully."
} catch {
    Write-Host "CSV was not able to be imported. Stopping script." -ForegroundColor red
    Exit-PSSession
}

# Count for end reporting
$ContactsCount = 0

# Create contact for each row in CSV
foreach ($Contact in $ContactsCsv) {
    if ($Contact.FirstName -ne 'FirstName') {
        # Save variables in a way that strings can read
        $FirstName = $Contact.FirstName
        $LastName = $Contact.LastName
        $Email = $Contact.Email
        $Title = $Contact.JobTitle
        $CompanyName = $Contact.CompanyName
        $MobilePhone = $Contact.MobilePhone
        $DirectPhone = $Contact.DirectPhone
        $SiteId = $Contact.SiteId

        # Get the company (id)
        $Company = Get-CWMCompany -Condition "name LIKE '$CompanyName' AND deletedFlag=false"

        # Remove non-numeric characters
        $RegexPattern = '[^0-9]'
        $MobilePhone = $MobilePhone -replace $RegexPattern, ''
        $DirectPhone = $DirectPhone -replace $RegexPattern, ''

        # Checks to make sure the phone number provided is valid
        if (($MobilePhone[0] -ne "0") -and ($MobilePhone.Length -eq 9)) { $MobilePhone="0"+$MobilePhone }
        if (($DirectPhone[0] -ne "0") -and ($DirectPhone.Length -eq 9)) { $DirectPhone="0"+$DirectPhone }

        $CommunicationsArray = @()
        $EmailComm = @{
            type=@{id=1;name='Email'}
            value=$Email
            domain=$Domain
            defaultFlag=$True;
            communicationType='Email'
        }
        $CommunicationsArray += $EmailComm

        # If we are given an office phone number, add that to the ConnectWise contact
        if(($DirectPhone -ne "")-and($null -ne $DirectPhone)-and($DirectPhone -ne "null")-and($DirectPhone -ne 0)){
            $Direct = @{
                type=@{id=2;name='Direct'}
                value=$DirectPhone
                defaultFlag=$False;
                communicationType='Phone'
            }
            $CommunicationsArray += $Direct
        }

        # If we are given a mobile phone number, add that to the ConnectWise contact
        if(($MobilePhone -ne "")-and($null -ne $MobilePhone)-and($MobilePhone -ne "null")-and($MobilePhone -ne 0)){
            $Mobile = @{
                type=@{id=4;name='Mobile'}
                value=$MobilePhone
                defaultFlag=$True;
                communicationType='Phone'
            }
            $CommunicationsArray += $Mobile      
        }

        $CompanyId = $Company.id

        # Set the users title to the one provided. If none is given, set it to " " (space) as the CWM script does not allow empty strings to be given
        if (($null -eq $Title) -or ($Title -eq "")) { $Title=" " }

        # Checks to see if the contact already exists
        Write-Host "`nChecking to see if a contact for " -NoNewLine
        Write-Host "$FirstName $LastName ($Title) " -ForegroundColor blue -NoNewline
        Write-Host "already exists..."
        
        $FirstNameTemp = "'"+$FirstName+"'" # Enclose in quotes so that the contact check accepts the string
        $LastNameTemp = "'"+$LastName+"'"
        $TitleTemp = "'"+$Title+"'"
        $ContactCheck = Get-CWMCompanyContact -Condition "firstName = $FirstNameTemp AND lastName = $LastNameTemp AND title = $TitleTemp AND company/id = $CompanyId AND site/id = $SiteId" -all

        if ($null -eq $ContactCheck) {
            Write-Host "Duplicate contact does not exist. Creating contact for " -NoNewline
            Write-Host "$FirstName $LastName ($Title)" -ForegroundColor blue -NoNewline
            Write-Host "..."
            try {
                $NewContact = New-CWMCompanyContact -firstName $FirstName -lastName $LastName -title $Title -company @{id=$CompanyId} -site @{id=$SiteId} -communicationItems $CommunicationsArray
                if ($null -ne $NewContact) {
                    Write-Host "Created contact for " -NoNewline
                    Write-Host "$FirstName $LastName ($Title)" -ForegroundColor blue -NoNewline
                    Write-Host "."
                    $ContactsCount += 1
                } else { 
                    Write-Host "Failed to create contact for " -ForegroundColor red -NoNewline
                    Write-Host "$FirstName $LastName ($Title)" -ForegroundColor blue -NoNewline
                    Write-Host "." -ForegroundColor red }
            }
            catch {
                Write-Host "Failed to create contact for " -ForegroundColor red -NoNewline
                Write-Host "$FirstName $LastName ($Title)" -ForegroundColor blue -NoNewline
                Write-Host "." -ForegroundColor red
            }
        }
        else {
            # Update contact to be active
            Write-Host "A contact for " -NoNewline; Write-Host "$FirstName $LastName ($Title) " -ForegroundColor blue -NoNewline; Write-Host "already exists. Marking as active..."
            try {
                $UpdateContact = Update-CWMCompanyContact -id $ContactCheck.id -Operation replace -Path 'inactiveFlag' -Value $false
                if ($null -ne $UpdateContact) {
                    Write-Host "Updated contact for " -NoNewline
                    Write-Host "$FirstName $LastName ($Title) " -ForegroundColor blue -NoNewline
                    Write-Host "."
                    $ContactsCount += 1
                } else { 
                    Write-Host "Failed to update contact for " -ForegroundColor red -NoNewline
                    Write-Host "$FirstName $LastName ($Title)" -ForegroundColor blue -NoNewline
                    Write-Host "." -ForegroundColor red }
            }
            catch {
                Write-Host "Failed to update contact for " -ForegroundColor red -NoNewline
                Write-Host "$FirstName $LastName ($Title)" -ForegroundColor blue -NoNewline
                Write-Host "." -ForegroundColor red
            }
        }
    } else { Write-Host "`nSkipping first line, as it's just the CSV headers..." }
}

Disconnect-CWM

Write-Host "`nContacts imported. Number of contacts created/updated: " -NoNewline; Write-Host "$ContactsCount" -ForegroundColor blue