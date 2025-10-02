param (
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
    [string]$SiteId,
	[bool]$NonSupportUser = $false
)

$AzKeyVaultName = Get-AutomationVariable -Name 'AzKeyVaultName'
$Log=""

#Phone Number String Validation
if(($MobilePhone[0] -eq "4") -and ($MobilePhone.Length -eq 9)){$MobilePhone="0"+$MobilePhone}
if(($OfficePhone[0] -eq "7") -and ($OfficePhone.Length -eq 9)){$OfficePhone="0"+$OfficePhone}

## CONNECT TO AZURE KEY VAULT ##

$AzKeyVaultConnectionName = 'AzureRunAsConnection'
try {
    # Get the connection properties
    $ServicePrincipalConnection = Get-AutomationConnection -Name $AzKeyVaultConnectionName
    $null = Connect-AzAccount `
        -ServicePrincipal `
        -TenantId $ServicePrincipalConnection.TenantId `
        -ApplicationId $ServicePrincipalConnection.ApplicationId `
        -CertificateThumbprint $ServicePrincipalConnection.CertificateThumbprint 
} catch {
    if (!$ServicePrincipalConnection) {
        # Azure Run As Account has not been enabled
        $ErrorMessage = "Connection $AzKeyVaultConnectionName not found."
        throw $ErrorMessage
    } else {
        # Something else went wrong
        Write-Error -Message $_.Exception.Message
        throw $_.Exception
    }
}

#### CREATING USER IN CONNECTWISE MANAGE ####

# Keys from the Azure key vault
$CWMApiServer = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWMApi-Server' -AsPlainText
$CWMApiCompany = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWMApi-CompanyId' -AsPlainText
$CWMApiPublicKey = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWMApi-PublicKey' -AsPlainText
$CWMApiPrivateKey = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWMApi-PrivateKey' -AsPlainText
$CWMApiClientId = Get-AzKeyVaultSecret -VaultName $AzKeyVaultName -Name 'MIT-CWMApi-ClientId' -AsPlainText

# Data validation
$CWMApiServer = $CWMApiServer.replace(" ", "")
$CWMApiPublicKey = $CWMApiPublicKey.replace("mit+", "")

# Create a credential object
$Connection = @{
    Server = $CWMApiServer
    Company = $CWMApiCompany
    pubkey = $CWMApiPublicKey
    privatekey = $CWMApiPrivateKey
    clientId = $CWMApiClientId
}

# Connect to Manage server
Connect-CWM @Connection

# Get the company (id)
$Company = Get-CWMCompany -Condition "name LIKE '$CompanyName' AND deletedFlag=false"
$CompanyId = $Company.id

$CommunicationsArray = @()
$EmailComm = @{
    type=@{id=1;name='Email'}
    value=$Email
    domain=$Domain
    defaultFlag=$True;
    communicationType='Email'
}
$CommunicationsArray += $EmailComm

#If we are given an office phone number, add that to the ConnectWise contact
if(($OfficePhone -ne "")-and($OfficePhone -ne $null)-and($OfficePhone -ne "null")-and($OfficePhone -ne 0)){
    $Direct = @{
        type=@{id=2;name='Direct'}
        value=$OfficePhone
        defaultFlag=$False;
        communicationType='Phone'
    }
    $CommunicationsArray += $Direct
}

# If we are given a mobile phone number, add that to the ConnectWise contact
if(($MobilePhone -ne "")-and($MobilePhone -ne $null)-and($MobilePhone -ne "null")-and($MobilePhone -ne 0)){
    $Mobile = @{
        type=@{id=4;name='Mobile'}
        value=$MobilePhone
        defaultFlag=$True;
        communicationType='Phone'
    }
    $CommunicationsArray += $Mobile      
}

# Sets type of contact based on provided bool
if ($NonSupportUser) { $ContactType = @{ id = 27 } # Non-Support User
} else { $ContactType = @{ id = 10 } } # End User
$ContactTypes = @($ContactType)

# Checks to see if the contact already exists
$Log += "Checking to see if CWM Contact for $FirstName $LastName - $Title exists<br>"
if (($Title -eq $null) -or ($Title -eq "")) { $Title=" " }
# Enclose in quotes so that the contact check accepts the string
$FirstNameTemp = "'"+$FirstName+"'"
$LastNameTemp = "'"+$LastName+"'"
$TitleTemp = "'"+$Title+"'"
try { $ContactCheck = Get-CWMCompanyContact -Condition "firstName = $FirstNameTemp AND lastName = $LastNameTemp AND title = $TitleTemp AND company/id = $CompanyId AND site/id = $SiteID" -all }
catch {}

if ($null -eq $ContactCheck) {
	# Set the users title to the one provided. If none is given, set it to " " (space) as the CWM script does not allow empty strings to be given
	$Log += "Duplicate contact does not exist. Creating CWM Contact for $FirstName $LastName - $Title<br>"
	try {
		$Contact = New-CWMContact -firstName $FirstName -lastName $LastName -title $Title -company @{id=$CompanyId} -site @{id=$SiteID} -communicationItems $CommunicationsArray -types $ContactTypes
		$Log += "Created CWM Contact for $FirstName $LastName<br>"
	}
	catch {
		$Log += "Failed to create CWM Contact for $FirstName $LastName<br>"
	}
}
else {
	# Update contact to be active
	$Log += "CWM Contact for $FirstName $LastName - $Title exists, marking as active.<br>"
	try {
		$Contact = Update-CWMCompanyContact -id $ContactCheck.id -Operation replace -Path 'inactiveFlag' -Value $false
		$Log += "Updated CWM Contact for $FirstName $LastName<br>"
	}
	catch {
		$Log += "Failed to update CWM Contact for $FirstName $LastName<br>"
	}
}

Disconnect-CWM

$Output = @{	
	CompanyId = $CompanyId
	ContactId = $Contact.id
	Log = $Log
}

Write-Output ($Output | ConvertTo-Json)