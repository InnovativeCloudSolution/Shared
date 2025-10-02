param(
    [string]$FirstName='',
    [string]$LastName='',
    [string]$MobilePhone='',
    [string]$Title='',
    [string]$ServiceDeskLevel,
    [switch]$ServiceDeliveryTeam,
    [switch]$AdminRequired
)
# Phone Number String Validation
if($MobilePhone[0] -eq "4"){$MobilePhone="0"+$MobilePhone}

$Log="The parameters that have been provided are:`r`n"
$Log+="First Name - $FirstName`r`n"
$Log+="Last Name - $LastName`r`n"
$Log+="Title - $Title`r`n"
$Log+="Mobile Phone - $MobilePhone`r`n"

$Errors = ""

# Company name in CWM
$CompanyName = 'Mangano IT'
# Email domain (for CWM and users)
$Domain = '@manganoit.com.au'

$TenantID = "5792a6c1-f4fe-466b-b97c-10eaf4fb3122" #The ManganoIT Tenancy

#Generate a password, convert it to a secure string
$Log+="`r`nGenerating password through DinoPass`r`n"
#Set the password being safe to false

$Password = Invoke-restmethod -uri "https://www.dinopass.com/password/strong"
$Password += Invoke-restmethod -uri "https://www.dinopass.com/password/strong"

#Because we dont want to use partner centre in any way shape or form
Write-Host "Please authenticate using your User Administrator account for Msol"
$MsolConnection = Connect-MsolService

#Test Connection:
Get-MsolDomain -ErrorAction SilentlyContinue
if($?)
{
    Write-Host "Connected to MSOL. Continuing."
}
else
{
    Write-Error "You are not connected to MSOL. Stopping Script"
    exit
}

#Create the user
$Username = $FirstName+'.'+ $LastName
$Email = $Username+$Domain
$DisplayName = $FirstName+' '+ $LastName
$Log+="`r`nConstructing the username and email for the user`r`n"
$Log+="Username: $Username`r`n"
$Log+="Email: $Email`r`n"

Try{
    $Log+="`r`nCreating the AD User`r`n"
    $MSOLUser = New-MsolUser -TenantId $TenantID -FirstName $FirstName -LastName $LastName -PhoneNumber "07 3151 9000" `
    -UserPrincipalName $Email -DisplayName $DisplayName -City "Hamilton" -StreetAddress "9 Hercules Street" `
    -PostalCode "4007" -State "Queensland" -Country "AU" -Title $Title -UsageLocation "AU" -Password $Password `
    -ForceChangePassword $false
    $Log+="AD Account created with Username: $Username `r`n"
    $createduser="true"
    $Errors="false"
}
Catch{
    $Log+="Error: $Username already exists`r`n"
    $createduser="false"
    $Errors="true"
}

##############GROUPS#########################
#SG.License.StandardUser.M365E5
$GroupName = "SG.License.M365E5.StandardUser"
$GroupID = Get-MsolGroup -TenantId $TenantID | Where-Object{$_.DisplayName -eq $GroupName}
Add-MsolGroupMember -GroupObjectId $GroupID.ObjectId -GroupMemberObjectId $MSOLUser.ObjectId -TenantId $TenantID
$Log += "The user $Email has been added to the $GroupName group.`r`n"

#App_ConnectwiseManage
$GroupName = "App_ConnectwiseManage"
$GroupID = Get-MsolGroup -TenantId $TenantID | Where-Object{$_.DisplayName -eq $GroupName}
Add-MsolGroupMember -GroupObjectId $GroupID.ObjectId -GroupMemberObjectId $MSOLUser.ObjectId -TenantId $TenantID
$Log += "The user $Email has been added to the $GroupName group.`r`n"

#lp.synchronizedusers
$GroupName = "lp.synchronizedusers"
$GroupID = Get-MsolGroup -TenantId $TenantID | Where-Object{$_.DisplayName -eq $GroupName}
Add-MsolGroupMember -GroupObjectId $GroupID.ObjectId -GroupMemberObjectId $MSOLUser.ObjectId -TenantId $TenantID
$Log += "The user $Email has been added to the $GroupName group.`r`n"

#SG.Azure.BlockNonManganoIP
$GroupName = "SG.Azure.BlockNonManganoIP"
$GroupID = Get-MsolGroup -TenantId $TenantID | Where-Object{$_.DisplayName -eq $GroupName}
Add-MsolGroupMember -GroupObjectId $GroupID.ObjectId -GroupMemberObjectId $MSOLUser.ObjectId -TenantId $TenantID
$Log += "The user $Email has been added to the $GroupName group.`r`n"

#MFA Disable
$GroupName = "MFA Disable"
$GroupID = Get-MsolGroup -TenantId $TenantID | Where-Object{$_.DisplayName -eq $GroupName}
Add-MsolGroupMember -GroupObjectId $GroupID.ObjectId -GroupMemberObjectId $MSOLUser.ObjectId -TenantId $TenantID
$Log += "The user $Email has been added to the $GroupName group.`r`nRemember to remove the user from the above two groups.`r`n"

if($AdminRequired){

    #Admin Firstname Surname (No firstname/surname)
    $AdminUsername = "adm."+$FirstName+'.'+ $LastName
    $AdminEmail = $AdminUsername+$Domain
    $AdminDisplayName = "Admin "+$FirstName+' '+ $LastName
    $Log+="`r`nConstructing the username and email for the user`r`n"
    $Log+="Username: $Username`r`n"
    $Log+="Email: $Email`r`n"

    Try{
        $Log+="`r`nCreating the AD Admin User`r`n"
        $AdminMSOLUser = New-MsolUser -TenantId $TenantID -UserPrincipalName $AdminEmail -DisplayName $AdminDisplayName `
        -MobilePhone $MobilePhone -UsageLocation "AU" -Password $Password -ForceChangePassword $false
        $Log+="AD Account created with Username: $Username `r`n"
        $AdminCreatedUser="true"
    }
    Catch{
        $Log+="Error: $Username already exists`r`n"
        $AdminCreatedUser="false"
        $Errors="true"
    }

    #SG.License.AdminUser.M365E5
    $GroupName = "SG.License.M365E5.Admin"
    $GroupID = Get-MsolGroup -TenantId $TenantID | Where-Object{$_.DisplayName -eq $GroupName}
    Add-MsolGroupMember -GroupObjectId $GroupID.ObjectId -GroupMemberObjectId $AdminMSOLUser.ObjectId -TenantId $TenantID
    $Log += "The user $AdminEmail has been added to the $GroupName group.`r`n"

    #AdminAgents
    $GroupName = "AdminAgents"
    $GroupID = Get-MsolGroup -TenantId $TenantID | Where-Object{$_.DisplayName -eq $GroupName}
    Add-MsolGroupMember -GroupObjectId $GroupID.ObjectId -GroupMemberObjectId $AdminMSOLUser.ObjectId -TenantId $TenantID
    $Log += "The user $AdminEmail has been added to the $GroupName group.`r`n"

    #SG.Role.ITGlueSSO
    $GroupName = "SG.Role.ITGlueSSO"
    $GroupID = Get-MsolGroup -TenantId $TenantID | Where-Object{$_.DisplayName -eq $GroupName}
    Add-MsolGroupMember -GroupObjectId $GroupID.ObjectId -GroupMemberObjectId $AdminMSOLUser.ObjectId -TenantId $TenantID
    $Log += "The user $AdminEmail has been added to the $GroupName group.`r`n"

    #SG.Azure.BlockNonManganoIP
    $GroupName = "SG.Azure.BlockNonManganoIP"
    $GroupID = Get-MsolGroup -TenantId $TenantID | Where-Object{$_.DisplayName -eq $GroupName}
    Add-MsolGroupMember -GroupObjectId $GroupID.ObjectId -GroupMemberObjectId $MSOLUser.ObjectId -TenantId $TenantID
    $Log += "The user $Email has been added to the $GroupName group.`r`n"

    #MFA Disable
    $GroupName = "MFA Disable"
    $GroupID = Get-MsolGroup -TenantId $TenantID | Where-Object{$_.DisplayName -eq $GroupName}
    Add-MsolGroupMember -GroupObjectId $GroupID.ObjectId -GroupMemberObjectId $MSOLUser.ObjectId -TenantId $TenantID
    $Log += "The user $Email has been added to the $GroupName group.`r`nRemember to remove the user from the above two groups.`r`n"

    if ($ServiceDeskLevel -eq "2") {
        #SG.Role.LogicMonitor.Contributor
        $GroupName = "SG.Role.LogicMonitor.Contributor"
        $GroupID = Get-MsolGroup -TenantId $TenantID | Where-Object{$_.DisplayName -eq $GroupName}
        Add-MsolGroupMember -GroupObjectId $GroupID.ObjectId -GroupMemberObjectId $AdminMSOLUser.ObjectId -TenantId $TenantID
        $Log += "The user $AdminEmail has been added to the $GroupName group.`r`n"
    }

    if ($ServiceDeskLevel -eq "3") {
        #SG.Role.LogicMonitor.Admin
        $GroupName = "SG.Role.LogicMonitor.Admin"
        $GroupID = Get-MsolGroup -TenantId $TenantID | Where-Object{$_.DisplayName -eq $GroupName}
        Add-MsolGroupMember -GroupObjectId $GroupID.ObjectId -GroupMemberObjectId $AdminMSOLUser.ObjectId -TenantId $TenantID
        $Log += "The user $AdminEmail has been added to the $GroupName group.`r`n"
    }

    else {
        #SG.Role.LogicMonitor.Reader
        $GroupName = "SG.Role.LogicMonitor.Reader"
        $GroupID = Get-MsolGroup -TenantId $TenantID | Where-Object{$_.DisplayName -eq $GroupName}
        Add-MsolGroupMember -GroupObjectId $GroupID.ObjectId -GroupMemberObjectId $AdminMSOLUser.ObjectId -TenantId $TenantID
        $Log += "The user $AdminEmail has been added to the $GroupName group.`r`n"
    }
}

Write-Host "Thanks! The account has been created! Please wait a few minutes while a license is added to the account!"

for ($i=0; $i -le 300; $i++){
    $Percent = [math]::Round($i/300*100)
    Write-Progress -Activity "Waiting for license addition" -Status "$Percent% Complete:" -PercentComplete $Percent;
    Start-Sleep -Seconds 1
}

Write-Host "Please authenticate using your Exchange Administrator account for Exchange"
$ExchangeConnection = Connect-ExchangeOnline

while ($null -eq (Get-Mailbox $Email -ErrorAction SilentlyContinue).Name){
    Write-Host "User not found in Exchange Online. Please check that license is assigned and the user exists in Exchange Online"
    pause
}

Enable-Mailbox -Identity $Email -Archive

#!All Techs
$GroupName = "!All Techs"
Add-DistributionGroupMember -Identity $GroupName -Member $Email
$Log += "The user $Email has been added to the $GroupName group.`r`n"

#Email Signature Group
$GroupName = "Email Signature Group"
Add-DistributionGroupMember -Identity $GroupName -Member $Email
$Log += "The user $Email has been added to the $GroupName group.`r`n"

#!All Team
$GroupName = "!All Team"
Add-DistributionGroupMember -Identity $GroupName -Member $Email
$Log += "The user $Email has been added to the $GroupName group.`r`n"

#Tech Team
$GroupName = "Tech Team"
Add-UnifiedGroupLinks -Identity $GroupName -LinkType "Members" -Links $Email
$Log += "The user $Email has been added to the $GroupName group.`r`n"

#Team Mangano
$GroupName = "Team Mangano"
Add-UnifiedGroupLinks -Identity $GroupName -LinkType "Members" -Links $Email
$Log += "The user $Email has been added to the $GroupName group.`r`n"

if($ServiceDeliveryTeam){
    #Service Delivery Team
    $GroupName = "ServiceDelivery@manganoit.com.au"
    Add-DistributionGroupMember -Identity $GroupName -Member $Email
    $Log += "The user $Email has been added to the $GroupName group.`r`n"
}

if($ServiceDeskLevel){
    #Service Desk Level X
    $GroupName = "Service Desk Level "+$ServiceDeskLevel
    Add-UnifiedGroupLinks -Identity $GroupName -LinkType "Members" -Links $Email
    $Log += "The user $Email has been added to the $GroupName group.`r`n"
}

#SMS Shared Mailbox Delegation
$GroupName = "sms@manganoit.com.au"
Add-MailboxPermission -Identity $GroupName -User $Email -AccessRights FullAccess

$calendarIdentity = "$Email`:\calendar"

# Get Calendar permissions
$calendarPermissions = Get-MailboxFolderPermission $calendarIdentity
foreach ($permission in $calendarPermissions)
{
    if ($permission.User.DisplayName -ne "Default")
    {
        continue
    }
            
    if ($permission.AccessRights -notcontains 'LimitedDetails')
    {
        Set-MailboxFolderPermission -User "Default" -AccessRights 'LimitedDetails' -Identity $calendarIdentity
    }
            
    break
}

$json = @"
{
"Email": "$Email",
"Username": "$Username",
"Password": "$Password",
"CreatedUser":$createduser,
"AdminEmail":$AdminEmail,
"AdminCreatedUser":$AdminCreatedUser,
"Errors":$Errors,
"Log":"$Log"
}
"@
Write-Output $json