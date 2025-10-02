param(
    [string]$FirstName='',
    [string]$LastName='',
    [string]$MobilePhone='',
    [string]$Title='',
    [string]$Site='',
    [string]$SupportQueue='',
    [string]$EAMQueue='',
    [string]$NationalQueue=''
)
#Phone Number String Validation
if($MobilePhone[0] -eq "4"){$MobilePhone="0"+$MobilePhone}

$Log="The parameters that have been provided are:<br>"
$Log+="First Name - $FirstName<br>"
$Log+="Last Name - $LastName<br>"
$Log+="Title - $Title<br>"
$Log+="Mobile Phone - $MobilePhone<br>"
$Log+="Site - $Site<br>"

$Errors = ""

#Company name in CWM
$CompanyName = 'Elston'
#Email domain (for CWM and users)
$Domain = '@elston.com.au'
$TenantID = '2939c5b6-63d7-430c-a345-aba7b3d6ab1b'

$MITSTenant = "5792a6c1-f4fe-466b-b97c-10eaf4fb3122" #The ManganoIT Tenancy

#Teams Calling Arrays
$Names            = @("Brisbane", "Gold Coast", "Hervey Bay", "Canberra", "Ballina", "Sydney")
$CallerIDs        = @("bris1", "gold1", "herv1", "canb1", "ball1", "sydn1")
$SiteNumber = [array]::indexof($Names, $Site)

#Generate a password, convert it to a secure string
$Log+="<br>Generating password through DinoPass<br>"
#Set the password being safe to false
$SafePass = $false

Function Test-PasswordForDomain {
    Param (
        [Parameter(Mandatory=$true)][string]$Password
    )

    If ($Password.Length -lt 8) {
        return $false
    }

    if (
       ($Password -cmatch "[A-Z\p{Lu}\s]") `
       -and ($Password -cmatch "[a-z\p{Ll}\s]") `
    ) 
    { 
        return $true
    } else {
        return $false
    }
}

#Set the 'Password' to be something that the domain will accept 
Do {
    $Password = Invoke-restmethod -uri "https://www.dinopass.com/password/strong"
    $SafePass = Test-PasswordForDomain $Password
} While ($SafePass -eq $False)
# $Password Generated

############### AZURE AD VIA MSONLINE ###############

$PartnerCenterApp = Get-AutomationPSCredential -Name 'PartnerCenterApp'
$ClientAppId = $PartnerCenterApp.UserName
$ClientSecret = $PartnerCenterApp.Password
$refreshToken = Get-AutomationVariable -Name 'PartnerCenterAppRefreshToken'
$Credential = New-Object System.Management.Automation.PSCredential ($ClientAppId, $ClientSecret)
$aadGraphToken = New-PartnerAccessToken -ApplicationId $ClientAppId -Credential $Credential -RefreshToken $refreshToken -Scopes 'https://graph.windows.net/.default' -ServicePrincipal -Tenant $MITSTenant
$graphToken = New-PartnerAccessToken -ApplicationId $ClientAppId -Credential $Credential -RefreshToken $refreshToken -Scopes 'https://graph.microsoft.com/.default' -ServicePrincipal -Tenant $MITSTenant
$Connection = Connect-MsolService -AdGraphAccessToken $aadGraphToken.AccessToken -MsGraphAccessToken $graphToken.AccessToken

#Create the user
$Username = $FirstName+'.'+ $LastName
$Email = $Username+$Domain
$DisplayName = $FirstName+' '+ $LastName
$Log+="<br>Constructing the username and email for the user<br>"
$Log+="Username: $Username<br>"
$Log+="Email: $Email<br>"


Try{
    $Log+="<br>Creating the AD User<br>"
    $MSOLUser = New-MsolUser -TenantId $TenantID -FirstName $FirstName -LastName $LastName -UserPrincipalName $Email -DisplayName $DisplayName -City $Site -Country "AU" -Title $Title -MobilePhone $MobilePhone -UsageLocation "AU" -Password $Password
    $Log+="AD Account created with Username: $Username <br>"
    $createduser="true"
    $Errors="false"
}
Catch{
    $Log+="Error: $Username already exists<br>"
    $createduser="false"
    $Errors="true"
}


##############GROUPS#########################
$Log += "Support Queue = $SupportQueue<br>"
if ($createduser -AND $SupportQueue -eq 'Include'){
    $GroupName = "teams."+$CallerIDs[$SiteNumber]+".support"
    $GroupID = Get-MsolGroup -TenantId $TenantID | Where-Object{$_.DisplayName -eq $GroupName}
    Add-MsolGroupMember -GroupObjectId $GroupID.ObjectId -GroupMemberObjectId $MSOLUser.ObjectId -TenantId $TenantID
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "National Queue = $NationalQueue<br>"
if ($createduser -AND $NationalQueue -eq 'Include'){
    $GroupName = "teams.national.queue3.static"
    $GroupID = Get-MsolGroup -TenantId $TenantID | Where-Object{$_.DisplayName -eq $GroupName}
    Add-MsolGroupMember -GroupObjectId $GroupID.ObjectId -GroupMemberObjectId $MSOLUser.ObjectId -TenantId $TenantID
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "EAM Queue = $EAMQueue<br>"
if ($createduser -AND $EAMQueue -eq 'Include'){
    $GroupName = "teams.eam.support"
    $GroupID = Get-MsolGroup -TenantId $TenantID | Where-Object{$_.DisplayName -eq $GroupName}
    Add-MsolGroupMember -GroupObjectId $GroupID.ObjectId -GroupMemberObjectId $MSOLUser.ObjectId -TenantId $TenantID
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
if ($createduser -AND $Site -eq 'Sydney'){
    $GroupName = "teams.sydn1.allstaff"
    $GroupID = Get-MsolGroup -TenantId $TenantID | Where-Object{$_.DisplayName -eq $GroupName}
    Add-MsolGroupMember -GroupObjectId $GroupID.ObjectId -GroupMemberObjectId $MSOLUser.ObjectId -TenantId $TenantID
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}



$json = @"
{
"Email": "$Email",
"Username": "$Username",
"Password": "$Password",
"CreatedUser":$createduser,
"Errors":$Errors,
"Log":"$Log"
}
"@
Write-Output $json