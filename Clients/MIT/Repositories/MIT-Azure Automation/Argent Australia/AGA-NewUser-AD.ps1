param(
    #13
    [string]$FirstName='',
    [string]$LastName='',
    [string]$Title='',
    [string]$MobilePhone='',
    [string]$OfficePhone='',
    [string]$Site='Fortitude Valley',
    [string]$Domain='@argentaust.com.au',
    [string]$Manager,
    [string]$Department,
    [string]$StreetAddress = "Level 1, 22 Arthur Street",
    [string]$City = "Fortitude Valley",
    [string]$State = 'QLD',
    [string]$Postcode = "4006",
    [string]$OU='OU=FVAL1,OU=Users,OU=Argent,DC=internal,DC=argentaust,DC=com,DC=au',
    #20 DGs
    [string]$AllArgentStaff,
    [string]$AllFortitudeGroup,
    [string]$AllSalesStaff,
    [string]$StockAdvice,
    [string]$AllMastercardHolders,
    [string]$AllFortitudeValley,
    [string]$AllMobilePhones,
    [string]$AllProjectSales,
    [string]$AllSalesReps,
    [string]$AllRetailSales,
    [string]$AllNationalOffice,
    [string]$AllKAM,
    [string]$AllSydneyProjectGroup,
    [string]$AllSydneyOffice,
    [string]$SSC,
    [string]$AllLogisticsGroup,
    [string]$AllMelbourneOffice,
    [string]$AllCustomerCentral,
    [string]$AllBrisbaneSalesOffice,
    [string]$AllPerthOffice,
    #10 folders
    [string]$CustomerCentralFolderAccess,
    [string]$CorporateFolder,
    [string]$ExecutiveFolder,
    [string]$FinanceFolder,
    [string]$ForecastFolder,
    [string]$HRFolder,
    [string]$ITFolder,
    [string]$LTIWorkingsFolder,
    [string]$MarketingFolder,
    [string]$MarketingFolderReadOnly,
    # RemoteDesktop
    [string]$WVDAccess,
    [string]$AdobeLicensing
)

# Mobile number string validation
$MobilePhone = .\SMS-FormatValidPhoneNumber.ps1 -PhoneNumber $MobilePhone -KeepSpaces $true

# Office number string validation
$OfficePhone = .\SMS-FormatValidPhoneNumber.ps1 -PhoneNumber $OfficePhone -KeepSpaces $true -IsMobilePhone $false

$Date = Get-Date -Format "dd-MM-yyyy"
$FaxNumber = "1300 364 748"
$FileName = "$FirstName-$LastName-$Date"

"The parameters that have been provided are:<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
"First Name - $FirstName<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
"Last Name - $LastName<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
"Title - $Title<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
"Mobile Phone - $MobilePhone<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
"Site - $Site<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 

$Errors = ""

# Generate a password, convert it to a secure string
"<br>Generating password through DinoPass<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
# Set the password being safe to false
$SafePass = $false

Function Test-PasswordForDomain {
    Param (
        [Parameter(Mandatory=$true)][string]$Password,
        [Parameter(Mandatory=$false)][string]$AccountSamAccountName = "",
        [Parameter(Mandatory=$false)][string]$AccountDisplayName,
        $PasswordPolicy = (Get-ADDefaultDomainPasswordPolicy -ErrorAction SilentlyContinue)
    )

    If ($Password.Length -lt $PasswordPolicy.MinPasswordLength) {
        "Password Doesnt meet password length"  | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
        return $false
    }

    if (($AccountSamAccountName) -and ($Password -match "$AccountSamAccountName")) {
        "Password Matches SAMAccountName"  | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
        return $false
    }

    if ($AccountDisplayName) {
        $tokens = $AccountDisplayName.Split(",.-,_ #`t")
        foreach ($token in $tokens) {
            if (($token) -and ($Password -match "$token")) {
                "Password Matches token $token" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
                return $false
            }
        }
    }

    if ($PasswordPolicy.ComplexityEnabled -eq $true) {
        If (
           ($Password -cmatch "[A-Z\p{Lu}\s]") `
           -and ($Password -cmatch "[a-z\p{Ll}\s]") `
           -and ($Password -match "[\d]") `
           -and ($Password -match "[^\w]")  
        ) { 
            return $true
        }
        else{
            "Doesnt meet complexity" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
            return $false
        }
    } else {
        return $true
    }
}

#Set the 'Password' to be something that the domain will accept 
Do {
    $Password = Invoke-restmethod -uri "https://www.dinopass.com/password/strong"
    $Password | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 

    $SafePass = Test-PasswordForDomain $Password
} While ($SafePass -eq $False)

# Convert new password string to be secure
$SecurePassword=ConvertTo-SecureString $Password -AsPlainText -Force

# Construct the username & email for the user
$Username = ($FirstName+'.'+ $LastName) -replace '[^\.^\-^\w]', ''
if ($Username.Length -gt 18) { $Username = $Username.Substring(0,18) }
$UsernameLowerCase = $Username.ToLower()
$Email = $UsernameLowerCase+$Domain
"<br>Constructing the username and email for the user<br>"   | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
"Username: $Username<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
"Email: $Email<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 

# Retreive DN of Manager
"<br>Confirming the DN of the provided manager ($Manager).<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
$ManagerName = $Manager.Split("<").Trim()[0]
$ManagerDN = Get-ADUser -Filter "Name -like '$ManagerName'" | % { $_.DistinguishedName}

# Construct the user details. Before the = sign is the AD account field.
$UserProperties=@{
	UserPrincipalName = $Email
	SAMAccountName = $Username
	Name = $FirstName + ' ' + $LastName
	GivenName = $FirstName
	Surname = $LastName
	Title = $Title
	Department = $Department
	StreetAddress = $StreetAddress 
	City = $City   
	PostalCode = $Postcode
	State = $State
	Fax = $FaxNumber
	DisplayName = $FirstName + ' ' + $LastName
	EmailAddress = $Email
	AccountPassword = $SecurePassword
	Path = $OU
	Country = "AU"
	Enabled = $True
}

# Add mobile number if provided and valid
if ($MobilePhone -ne '') {
    $UserProperties += @{ MobilePhone = $MobilePhone }
}

# Add office phone if provided and valid
if ($OfficePhone -ne '') {
    $UserProperties += @{ TelephoneNumber = $OfficePhone }
}

if (!$ManagerDN) {
    "The manager that was provided could not be found.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
    $ManagerStatus = "false"
} else {
    "DN of manager is $ManagerDN<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
    $ManagerStatus = "true"

    # Add manager details to UserProperties
    $UserProperties += @{ Manager = $ManagerDN }
}

# Creation of AD User
try {
    # Try make the user
    "<br>Creating the AD User<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
    $UserProperties | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
    New-ADUser @UserProperties
    "AD Account created with Username: $Username <br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
    $CreatedUser = $true
    $Errors = "false"
    try {
        $TemplateStandard=get-aduser _Template.Standard -properties memberof
        $TemplateStandard.memberof | add-adgroupmember -members $Username
    }
    catch {
        "Failed to add $Username to Template Groups<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
    }
}
catch {
    # The user has failed to be made
    "Error: $Username already exists<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
    $CreatedUser = $false
    $Errors = "true"
}

#############AD Distribution Groups##############################
"AllArgentStaff = $AllArgentStaff<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $AllArgentStaff -eq 'True') {
    Add-ADGroupMember "All Argent Staff" $Username | Out-Null
    "The user $Email has been added to AllArgentStaff.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"AllFortitudeGroup = $AllFortitudeGroup<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $AllFortitudeGroup -eq 'True') {
    Add-ADGroupMember "All Fortitude Group" $Username | Out-Null
    "The user $Email has been added to AllFortitudeGroup.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"AllSalesStaff = $AllSalesStaff<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $AllSalesStaff -eq 'True') {
    Add-ADGroupMember "All Sales Staff" $Username | Out-Null
    "The user $Email has been added to AllSalesStaff.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"StockAdvice = $StockAdvice<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $StockAdvice -eq 'True') {
    Add-ADGroupMember "Stock Advice" $Username | Out-Null
    "The user $Email has been added to StockAdvice.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"AllMastercardHolders = $AllMastercardHolders<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $AllMastercardHolders -eq 'True') {
    Add-ADGroupMember "All Mastercard Holders" $Username | Out-Null
    "The user $Email has been added to AllMastercardHolders.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"AllFortitudeValley = $AllFortitudeValley<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $AllFortitudeValley -eq 'True') {
    Add-ADGroupMember "All Fortitude Valley" $Username | Out-Null
    "The user $Email has been added to AllFortitudeValley.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"AllMobilePhones = $AllMobilePhones<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $AllMobilePhones -eq 'True') {
    Add-ADGroupMember "All Mobile Phones" $Username | Out-Null
    "The user $Email has been added to AllMobilePhones.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"AllProjectSales = $AllProjectSales<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $AllProjectSales -eq 'True') {
    Add-ADGroupMember "All Project Sales" $Username | Out-Null
    "The user $Email has been added to AllProjectSales.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"AllSalesReps = $AllSalesReps<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $AllSalesReps -eq 'True') {
    Add-ADGroupMember "All Sales Representatives" $Username | Out-Null
    "The user $Email has been added to AllSalesReps.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"AllRetailSales = $AllRetailSales<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $AllRetailSales -eq 'True') {
    Add-ADGroupMember "All Retail Sales" $Username | Out-Null
    "The user $Email has been added to AllRetailSales.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"AllNationalOffice = $AllNationalOffice<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $AllNationalOffice -eq 'True') {
    Add-ADGroupMember "All National Office" $Username | Out-Null
    "The user $Email has been added to AllNationalOffice.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"AllKAM = $AllKAM<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $AllKAM -eq 'True') {
    Add-ADGroupMember "All KAM" $Username | Out-Null
    "The user $Email has been added to AllKAM.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"AllSydneyProjectGroup = $AllSydneyProjectGroup<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $AllSydneyProjectGroup -eq 'True') {
    Add-ADGroupMember "All Sydney Project Group" $Username | Out-Null
    "The user $Email has been added to AllSydneyProjectGroup.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"AllSydneyOffice = $AllSydneyOffice<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $AllSydneyOffice -eq 'True') {
    Add-ADGroupMember "All Sydney Office" $Username | Out-Null
    "The user $Email has been added to AllSydneyOffice.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"SSC = $SSC<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $SSC -eq 'True') {
    Add-ADGroupMember "SSC" $Username | Out-Null
    "The user $Email has been added to SSC.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"AllLogisticsGroup = $AllLogisticsGroup<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $AllLogisticsGroup -eq 'True') {
    Add-ADGroupMember "All Logistics Group" $Username | Out-Null
    "The user $Email has been added to AllLogisticsGroup.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"AllMelbourneOffice = $AllMelbourneOffice<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $AllMelbourneOffice -eq 'True') {
    Add-ADGroupMember "All Melbourne Office" $Username | Out-Null
    "The user $Email has been added to SAllMelbourneOfficeSC.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"AllCustomerCentral = $AllCustomerCentral<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $AllCustomerCentral -eq 'True') {
    Add-ADGroupMember "All Customer-Central" $Username | Out-Null
    "The user $Email has been added to AllCustomerCentral.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"AllBrisbaneSalesOffice = $AllBrisbaneSalesOffice<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $AllBrisbaneSalesOffice -eq 'True') {
    Add-ADGroupMember "All Brisbane Sales Office" $Username | Out-Null
    "The user $Email has been added to AllBrisbaneSalesOffice.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"AllPerthOffice = $AllPerthOffice<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $AllPerthOffice -eq 'True') {
    Add-ADGroupMember "All Perth Office" $Username | Out-Null
    "The user $Email has been added to SAllPerthOfficeSC.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}


######Security Groups#####

"CustomerCentralFolderAccess = $CustomerCentralFolderAccess<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $CustomerCentralFolderAccess -eq 'True') {
    Add-ADGroupMember "Customer Central Folder Access" $Username | Out-Null
    "The user $Email has been added to CustomerCentralFolderAccess.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"CorporateFolder = $CorporateFolder<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $CorporateFolder -eq 'True') {
    Add-ADGroupMember "Argent Corporate Folder Access" $Username | Out-Null
    "The user $Email has been added to CorporateFolder.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"ExecutiveFolder = $ExecutiveFolder<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $ExecutiveFolder -eq 'True') {
    Add-ADGroupMember "Argent executive users" $Username | Out-Null
    "The user $Email has been added to ExecutiveFolder.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"FinanceFolder = $FinanceFolder<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $FinanceFolder -eq 'True') {
    Add-ADGroupMember "Argent Finance folder access" $Username | Out-Null
    "The user $Email has been added to FinanceFolder.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"ForecastFolder = $ForecastFolder<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $ForecastFolder -eq 'True') {
    Add-ADGroupMember "Argent Forecast Folder Access" $Username | Out-Null
    "The user $Email has been added to ForecastFolder.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"HRFolder = $HRFolder<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $HRFolder -eq 'True') {
    Add-ADGroupMember "Argent HR Folder Access" $Username | Out-Null
    "The user $Email has been added to HRFolder.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"ITFolder = $ITFolder<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $ITFolder -eq 'True') {
    Add-ADGroupMember "Argent IT folder access" $Username | Out-Null
    "The user $Email has been added to ITFolder.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"LTIWorkingsFolder = $LTIWorkingsFolder<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $LTIWorkingsFolder -eq 'True') {
    Add-ADGroupMember "Argent LTI Workings Folder Access" $Username | Out-Null
    "The user $Email has been added to LTIWorkingsFolder.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"MarketingFolder = $MarketingFolder<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $MarketingFolder -eq 'True') {
    Add-ADGroupMember "Argent Marketing Folder Access" $Username | Out-Null
    "The user $Email has been added to MarketingFolder.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

"MarketingFolderReadOnly = $MarketingFolderReadOnly<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
if ($CreatedUser -AND $MarketingFolderReadOnly -eq 'True') {
    Add-ADGroupMember "Argent Marketing ReadOnly Group" $Username | Out-Null
    "The user $Email has been added to MarketingFolderReadOnly.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
}

if ($CreatedUser -AND $AdobeLicensing -eq 'Adobe CC Suite (License required)') {
    Add-ADGroupMember "SG.AVD.FullDesktop" $Username | Out-Null
    Add-ADGroupMember "SG.AVD.AdobeCCUsers" $Username | Out-Null
} elseif ($CreatedUser) {
    "WVDAccess = $WVDAccess<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
    if ($CreatedUser -AND $WVDAccess -eq 'Full Desktop') {
        Add-ADGroupMember "SG.AVD.FullDesktop" $Username | Out-Null
        "The user $Email has been added to SG.AVD.FullDesktop<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 

        if ($AdobeLicensing -eq 'Adobe Acrobat Professional DC (License required)') {
            "The user $Email has been added to SG.AVD.AdobeDCUsers.Full<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
            Add-ADGroupMember "SG.AVD.AdobeDCUsers.Full" $Username | Out-Null
        }
    }
    
    if ($CreatedUser -AND $WVDAccess -eq 'Published Apps') {
        Add-ADGroupMember "SG.AVD.RemoteApp" $Username | Out-Null
        "The user $Email has been added to SG.AVD.RemoteApp.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 

        if ($AdobeLicensing -eq 'Adobe Acrobat Professional DC (License required)') {
            "The user $Email has been added to SG.AVD.AdobeDCUsers<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
            Add-ADGroupMember "SG.AVD.AdobeDCUsers" $Username | Out-Null
        }
    }
}

if ($CreatedUser) {
	# Add user to site groups
	switch ($Site) {
		"Alexandria" {
			Add-ADGroupMember "SG.Citrix.Desktop.AXDR" $Username | Out-Null
			Add-ADGroupMember "SG.Citrix.Printing.AXDR" $Username | Out-Null
		}
		"Fortitude Valley" {
			Add-ADGroupMember "SG.Citrix.Desktop.FVAL" $Username | Out-Null        
			Add-ADGroupMember "SG.Citrix.Printing.FVAL" $Username | Out-Null        
		}
		"Osbourne Park" {
			Add-ADGroupMember "SG.Citrix.Desktop.OSPK" $Username | Out-Null    
			Add-ADGroupMember "SG.Citrix.Printing.OSPK" $Username | Out-Null    
		}
		"Pinkenba" {
			Add-ADGroupMember "SG.Citrix.Desktop.PKBA" $Username | Out-Null    
			Add-ADGroupMember "SG.Citrix.Printing.PKBA" $Username | Out-Null    
		}
		"South Melbourne" {
			Add-ADGroupMember "SG.Citrix.Desktop.SMEL" $Username | Out-Null
			Add-ADGroupMember "SG.Citrix.Printing.SMEL" $Username | Out-Null
		}
	}

	# Add user to CodeTwo signature groups
	switch ($Department) {
		"Admin" {
			Add-ADGroupMember "SG.User.ArgentGeneralSignature" $Username | Out-Null
		}
		"Arthaus" {
			Add-ADGroupMember "SG.User.ArthausSignature" $Username | Out-Null
		}
		"Clearance Centre" {
			Add-ADGroupMember "SG.User.ClearanceCenterSignature" $Username | Out-Null
		}
		"Customer Central" {
			Add-ADGroupMember "SG.User.CustomerCentralSignature" $Username | Out-Null
		}
		"Marketing" {
			Add-ADGroupMember "SG.User.ArgentGeneralSignature" $Username | Out-Null
		}
		"Projects" {
			Add-ADGroupMember "SG.User.ArgentProjectsSignature" $Username | Out-Null
		}
		"Retail" {
			Add-ADGroupMember "SG.User.ArgentRetailSignature" $Username | Out-Null
		}
		"Warehouse" {
			Add-ADGroupMember "SG.User.CustomerCentralSignature" $Username | Out-Null
		}
		Default {
			Add-ADGroupMember "SG.User.ArgentGeneralSignature" $Username | Out-Null
		}
	}

	# Add user to generic groups
    Add-ADGroupMember "SG.AllArgentStaff" $Username | Out-Null
    Add-ADGroupMember "SG.Access.CorpWifi" $Username | Out-Null
    Add-ADGroupMember "SG.Access.DataDriveStandard" $Username | Out-Null
	Add-ADGroupMember "SG.User.CodeTwoLicense" $Username | Out-Null
}

# Sleep for a few seconds for AD to catch up
Start-Sleep -Seconds 30

# Run an Azure ADSync
"<br>Running an Azure ADSync<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
$SyncResult = Invoke-Command -ScriptBlock {
    Start-ADSyncSyncCycle -PolicyType Delta
} | Select-Object result
"ADSync $SyncResult" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 

if ($CreatedUser) { $CreatedUserString="true" }
else { $CreatedUserString="false" }

if ($SyncResult -match "Success") { $SyncResultLog = "Success" }
else { $SyncResultLog = "Failure" }

$Log = Get-Content "C:\Scripts\NewUserCreation\$FileName.txt"

$json = @"
{
"Email": "$Email",
"Username": "$Username",
"Password": "$Password",
"Manager": "$ManagerStatus",
"FileName": "$FileName",
"CreatedUser":$CreatedUserString,
"Errors":$Errors,
"AzureAD": "$SyncResultLog",
"Log": "$Log"
}
"@
Write-Output $json