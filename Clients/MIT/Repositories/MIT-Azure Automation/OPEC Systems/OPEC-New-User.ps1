param(
    [string]$FirstName='',
    [string]$LastName='',
    [string]$MobilePhone='',
    [string]$Title='',
    [string]$OU='OU=BELR1,OU=OPEC Users,OU=MyBusiness,DC=opec,DC=com,DC=au',
    [string]$Domain='@opecsystems.com',
    [string]$Manager,
    [string]$Department,
    [string]$StreetAddress = "48-50/7 Narabang Way",
    [string]$City = "Belrose",
    [string]$Postcode = "2085",
	[string]$Country = "AU",
    [string]$OSCARDefaultAccess = ''
)

#Phone Number String Validation
if($MobilePhone[0] -eq "4"){$MobilePhone="0"+$MobilePhone}

$Date = Get-Date -Format "dd-MM-yyyy"
$FileName = "$FirstName-$LastName-$Date"

"The parameters that have been provided are:<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
"First Name - $FirstName<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
"Last Name - $LastName<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
"Title - $Title<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
"Mobile Phone - $MobilePhone<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
"Site - $Site<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 


$Errors = ""

#Generate a password, convert it to a secure string
"<br>Generating password through DinoPass<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
#Set the password being safe to false
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
    $PasswordA = Invoke-restmethod -uri "https://www.dinopass.com/password/strong"
    $PasswordB = Invoke-restmethod -uri "https://www.dinopass.com/password/strong"
    $Password = $PasswordA+$PasswordB
    $Password | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 

    $SafePass = Test-PasswordForDomain $Password
} While ($SafePass -eq $False)

#Convert new password string to be secure
$SecurePassword=ConvertTo-SecureString $Password -AsPlainText -Force

#Construct the username & email for the user
$Username = (($FirstName[0]+ $LastName) -replace '[^\.^\-^\w]', '').ToLower()

if($Username.Length -gt 18){$Username = $Username.Substring(0,18)}

$Email = $Username+$Domain
"<br>Constructing the username and email for the user<br>"   | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
"Username: $Username<br>"   | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
"Email: $Email<br>"   | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 

#Retreive DN of Manager
"<br>Confirming the DN of the provided manager ($Manager).<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
$ManagerDN =  Get-ADUser -Filter "UserPrincipalName -like '$Manager'" | % { $_.DistinguishedName}

# Define country code

if (!$ManagerDN) {
    "The manager that was provided could not be found.<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
    $ManagerStatus = "false"
    #Construct the user details. Before the = sign is the AD account field. As the manager was not returned do not attempt to fill that attribute.
    $props=@{
        UserPrincipalName = $Email
        SAMAccountName = $Username
	    Name  = $FirstName + ' ' + $LastName
	    GivenName = $FirstName
	    Surname = $LastName
        Title = $Title
        Department = $Department
        StreetAddress = $StreetAddress 
        City = $City   
        PostalCode = $Postcode
        State = $State
        MobilePhone = $MobilePhone
        DisplayName = $FirstName + ' ' + $LastName
    	EmailAddress = $Email
	    AccountPassword=$SecurePassword
	    Path = $OU
        Country = $Country
        Enabled = $True
    }
}else{
    "DN of manager is $ManagerDN<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
    $ManagerStatus = "true"
    #Construct the user details. Before the = sign is the AD account field. As the manager was returned fill that attribute.
    $props=@{
        UserPrincipalName = $Email
        SAMAccountName = $Username
	    Name  = $FirstName + ' ' + $LastName
	    GivenName = $FirstName
	    Surname = $LastName
        Title = $Title
        Manager = $ManagerDN
        Department = $Department
        StreetAddress = $StreetAddress 
        City = $City   
        PostalCode = $Postcode
        State = $State
        OfficePhone = $OfficePhone
        MobilePhone = $MobilePhone
        DisplayName = $FirstName + ' ' + $LastName
    	EmailAddress = $Email
	    AccountPassword=$SecurePassword
	    Path = $OU
        Country = $Country
        Enabled = $True
    }
}

#Creation of AD User
Try{
    #Try make the user
    "<br>Creating the AD User<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
    $props | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
    New-ADUser @props
    "AD Account created with Username: $Username <br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
    $createduser = $true
    $Errors = "false"
    Try{
        $TemplateStandard=get-aduser _Template.Standard -properties memberof
        $TemplateStandard.memberof | add-adgroupmember -members $Username
    }
    Catch{
        "Failed to add $Username to Template Groups<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
    }
}
Catch{
    #The users failed to be made
    "Error: $Username already exists<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
    $createduser = $false
    $Errors = "true"
}

#Sleep for a few seconds for AD to catch up
Start-Sleep -Seconds 30

#Run an Azure ADSync
"<br>Running an Azure ADSync<br>" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 
$SyncResult = Invoke-Command -ComputerName 'OPC-AZU1-DC1' -ScriptBlock {
   
    Start-ADSyncSyncCycle -PolicyType Delta
} | select result
"ADSync $SyncResult" | Out-File -FilePath "C:\Scripts\NewUserCreation\$FileName.txt" -Append 

if($createduser){
    $CreatedUserString="true"
}else{
    $CreatedUserString="false"
}

$Log += "FuelandTank = $FuelandTank<br>"
if ($FuelandTank -eq 'True'){
    $GroupName = "!Fuel & Tank"
    Add-ADGroupMember  $GroupName $Username
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "BLBOperations = $BLBOperations<br>"
if ($BLBOperations -eq 'True'){
    $GroupName = "BLBOperations"
    Add-ADGroupMember  $GroupName $Username
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "DefenceForwarding = $DefenceForwarding<br>"
if ($DefenceForwarding -eq 'True'){
    $GroupName = "Defence Forwarding"
    Add-ADGroupMember  $GroupName $Username
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "Marine = $Marine<br>"
if ($Marine -eq 'True'){
    $GroupName = "marine"
    Add-ADGroupMember  $GroupName $Username
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "Subsea = $Subsea<br>"
if ($Subsea -eq 'True'){
    $GroupName = "Subsea"
    Add-ADGroupMember  $GroupName $Username
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "Training = $Training<br>"
if ($Training -eq 'True'){
    $GroupName = "Training"
    Add-ADGroupMember  $GroupName $Username
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "VesselHire = $VesselHire<br>"
if ($VesselHire -eq 'True'){
    $GroupName = "vesselhire"
    Add-ADGroupMember  $GroupName $Username
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "OSCARDefaultAccess = $OSCARDefaultAccess<br>"
if ($OSCARDefaultAccess -eq 'Yes'){
    $GroupName = "SG.SharePoint.Default"
    Add-ADGroupMember  $GroupName $Username
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}

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
"AzureAD": "$SyncResult",
"Log": "$Log"
}
"@
Write-Output $json