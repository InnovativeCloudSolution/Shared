param(
    #2
    [string]$Site='',
    [string]$Email='',

#Email Distribution Lists
    [string]$ElstonPrivateWealth ='',
    [string]$EPWBrisbane ='',
    [string]$Marketing ='',
    [string]$ElstonASOs ='',
    [string]$ElstonAssociateAdvisers ='',
    [string]$EAM = '',


#Mail Enabled Security Groups 6
    [string]$ElstonGlobal='',
    [string]$ElstonOOL='',
    [string]$ElstonBAL='',
    [string]$ElstonCBR='',
    [string]$ElstonHVB='',
    [string]$ElstonBNE='',

#Public Folders 21
    [string]$Accounts ='',
    [string]$AssureComms ='',
    [string]$AssureInfo ='',
    [string]$CFO ='',
    [string]$ClientReviews ='',
    [string]$Comms ='',
    [string]$ContractNotes ='',
    [string]$ElstonReports ='',
    [string]$HR ='',
    [string]$Hub24 ='',
    [string]$IPSTeam ='',
    [string]$ITAssist ='',
    [string]$Payroll ='',
    [string]$Portfolios ='',
    [string]$Spam ='',
    [string]$Succession ='',
    [string]$UBSAdmin ='',
    [string]$Voicemail ='',
    [string]$WebEnquiries ='',
    [string]$XplanAssist ='',

#Shared Mailboxes 9
    [string]$EAMShared ='',
    [string]$ElstonAdmin ='',
    [string]$ElstonEPFS ='',
    [string]$ElstonPhilanthropicServices ='',
    [string]$ElstonWealthPartner ='',
    [string]$FundApplications ='',
    [string]$InvestorAccounts ='',
    [string]$ShareRegistryInfo ='',
    [string]$UnitPrices=''
)
$Log = ""
$Domain = "@elston.com.au"
$TenantID = '2939c5b6-63d7-430c-a345-aba7b3d6ab1b'
$MITSTenant = "5792a6c1-f4fe-466b-b97c-10eaf4fb3122" #The ManganoIT Tenancy
$createduser = $True
$CompanyName = "Elston"
############### TEAMS CALLING ###############

#Teams Calling Arrays
$Names            = @("Brisbane", "Gold Coast", "Hervey Bay", "Canberra", "Ballina", "Sydney")
$PhoneBlockLower  = @(61730023810, 61755573010, 61731513310, 61261534010, 61279032510, 61256462010)
$PhoneBlockHigher = @(61730023899, 61755573099, 61731513399, 61261534099, 61279032599, 61256462099)
$PhoneNumberStart = @(617300238,   617555730  , 617315133  , 612615340  , 612790325  , 612564620)
$CallerIDs        = @("bris1", "gold1", "herv1", "canb1", "ball1", "sydn1")
$DialingPlans     = @("QLDDialOut","QLDDialOut","QLDDialOut","NSWDialOut","NSWDialOut","NSWDialOut")
$SiteNumber = [array]::indexof($Names, $Site)
$Log += "Site Index Selected: $SiteNumber - "+$Names[$SiteNumber]+"<br>"

try{
    $Log += "Connecting to Teams Online<br>"
    $TeamsCredentials = Get-AutomationPSCredential -Name "TeamsAdmin@elstongroup.onmicrosoft.com"
    $TeamsSession = Connect-MicrosoftTeams -Credential $TeamsCredentials
    $Log += "Connected to Teams Online<br>"
}
catch{
    $Log += "Failed to connect to Teams Online<br>"
}

#Assign a phone number. 
$Log += "Picking phone number<br>"
$PhoneNumbers = Get-CsOnlineTelephoneNumber -TelephoneNumberStartsWith $PhoneNumberStart[$SiteNumber] -IsNotAssigned -InventoryType Subscriber
foreach($Number in $PhoneNumbers){
    if($Number.Id -gt $PhoneBlockLower[$SiteNumber]){
        $OfficePhone = $Number.Id
        break
    }
}
$Log += "Found Phone Number: "+$OfficePhone+"<br>"

try{
    Set-CsOnlineVoiceUser -Identity $Email -TelephoneNumber $OfficePhone
    $Log += "Added Office Phone to User in Teams<br>"
}catch{
    $Log += "Failed to add Office Phone to User in Teams<br>"
}

#Assign a dialling plan
$Log += "Assigning Dial Plan<br>"
try{
    Grant-CsTenantDialPlan -Identity $Email -PolicyName $DialingPlans[$SiteNumber]
    $Log += "Assigned Dial Plan to user"
}
catch{
    $Log += "Failed to assign Dial Plan to user"
}

#Assign Calling ID Policy
$Log += "Assigning Calling ID Policy<br>"
try{
    Grant-CsCallingLineIdentity -Identity $Email -PolicyName $CallerIDs[$SiteNumber]
    $Log += "Assigned Calling ID Policy<br>"
}
catch{
    $Log += "Failed to assign Calling ID Policy<br>"
}

$Log += "Teams Calling Additions Complete. Removing PSSession"
Disconnect-MicrosoftTeams

$Log += "Connecting to Exchange Online"
$ExchangeRefreshToken = Get-AutomationVariable -Name 'ExchangeAppRefreshToken'
$upn = "workflows@manganoit.com.au"
$token = New-PartnerAccessToken -Module ExchangeOnline -RefreshToken $ExchangeRefreshToken -Tenant $TenantID
$tokenValue = ConvertTo-SecureString "Bearer $($token.AccessToken)" -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($upn, $tokenValue)
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid?DelegatedOrg=$($TenantID)&BasicAuthToOAuthConversion=true" -Credential $credential -Authentication Basic -AllowRedirection
Import-Module -function Get-RetentionCompliancePolicy,Set-RetentionCompliancePolicy (Import-PSSession -Session $ExchangeSession -DisableNameChecking -AllowClobber ) -Global

###################################Public Folders###################################

$Log += "Accounts = $Accounts<br>"
if ($createduser -AND $Accounts -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\Accounts" -User $Email -AccessRights PublishingEditor | Out-Null
    $Log += "The user $Email has been added to Accounts.<br>"
}

$Log += "Assure Comms = $AssureComms<br>"
if ($createduser -AND $AssureComms -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\Assure Comms" -User $Email -AccessRights Editor | Out-Null
    $Log += "The user $Email has been added to Assure Comms.<br>"
}

$Log += "AssureInfo = $AssureInfo<br>"
if ($createduser -AND $AssureInfo -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\Assure Info" -User $Email -AccessRights Editor | Out-Null
    $Log += "The user $Email has been added to AssureInfo.<br>"
}

$Log += "CFO = $CFO<br>"
if ($createduser -AND $CFO -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\CFO" -User $Email -AccessRights Editor | Out-Null
    $Log += "The user $Email has been added to CFO.<br>"
}

$Log += "ClientReviews = $ClientReviews<br>"
if ($createduser -AND $ClientReviews -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\ClientReviews" -User $Email -AccessRights Editor | Out-Null
    $Log += "The user $Email has been added to ClientReviews.<br>"
}

$Log += "Comms = $Comms<br>"
if ($createduser -AND $Comms -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\Comms" -User $Email -AccessRights Reviewer | Out-Null
    $Log += "The user $Email has been added to Comms.<br>"
}

$Log += "ContractNotes = $ContractNotes<br>"
if ($createduser -AND $ContractNotes -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\Contract Notes" -User $Email -AccessRights Reviewer | Out-Null
    $Log += "The user $Email has been added to Contract Notes.<br>"
}

$Log += "ElstonReports = $ElstonReports<br>"
if ($createduser -AND $ElstonReports -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\Elston Reports" -User $Email -AccessRights PublishingEditor | Out-Null
    $Log += "The user $Email has been added to ElstonReports.<br>"
}

$Log += "HR = $HR<br>"
if ($createduser -AND $HR -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\HR" -User $Email -AccessRights Editor | Out-Null
    $Log += "The user $Email has been added to HR.<br>"
}

$Log += "Hub24 = $Hub24<br>"
if ($createduser -AND $Hub24 -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\Hub24" -User $Email -AccessRights PublishingEditor | Out-Null
    $Log += "The user $Email has been added to Hub24.<br>"
}

$Log += "IPSTeam = $IPSTeam<br>"
if ($createduser -AND $IPSTeam -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\IPS Team" -User $Email -AccessRights PublishingEditor | Out-Null
    $Log += "The user $Email has been added to IPSTeam.<br>"
}


$Log += "ITAssist = $ITAssist<br>"
if ($createduser -AND $ITAssist -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\IT Assist" -User $Email -AccessRights Owner | Out-Null
    $Log += "The user $Email has been added to ITAssist.<br>"
}

$Log += "Payroll = $Payroll<br>"
if ($createduser -AND $Payroll -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\Payroll" -User $Email -AccessRights PublishingEditor | Out-Null
    $Log += "The user $Email has been added to Payroll.<br>"
}

$Log += "Portfolios = $Portfolios<br>"
if ($createduser -AND $Portfolios -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\Portfolios" -User $Email -AccessRights Editor | Out-Null
    $Log += "The user $Email has been added to Portfolios.<br>"
}

$Log += "Spam = $Spam<br>"
if ($createduser -AND $Spam -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\Spam" -User $Email -AccessRights PublishingEditor | Out-Null
    $Log += "The user $Email has been added to Spam.<br>"
}


$Log += "Succession = $Succession<br>"
if ($createduser -AND $Succession -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\Succession" -User $Email -AccessRights Editor | Out-Null
    $Log += "The user $Email has been added to Succession.<br>"
}

$Log += "UBSAdmin = $UBSAdmin<br>"
if ($createduser -AND $UBSAdmin -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\UBS Admin" -User $Email -AccessRights Editor | Out-Null
    $Log += "The user $Email has been added to UBS Admin.<br>"
}

$Log += "Voicemail = $Voicemail<br>"
if ($createduser -AND $Voicemail -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\Voicemail" -User $Email -AccessRights Owner | Out-Null
    $Log += "The user $Email has been added to Voicemail.<br>"
}


$Log += "WebEnquiries = $WebEnquiries<br>"
if ($createduser -AND $WebEnquiries -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\Web Enquiries" -User $Email -AccessRights Reviewer | Out-Null
    $Log += "The user $Email has been added to WebEnquiries.<br>"
}

$Log += "XplanAssist = $XplanAssist<br>"
if ($createduser -AND $XplanAssist -eq 'True'){
    Add-PublicFolderClientPermission -Identity "\Xplan Assist" -User $Email -AccessRights Owner | Out-Null
    $Log += "The user $Email has been added to XplanAssist.<br>"
}

#############SHARED MAILBOXES##############################
$Log += "EAMShared = $EAMShared<br>"
if ($createduser -AND $EAMShared -eq 'True'){
    Add-MailboxPermission "EAM.Shared@elston.com.au" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to EAMShared.<br>"
}

$Log += "ElstonAdmin = $ElstonAdmin<br>"
if ($createduser -AND $ElstonAdmin -eq 'True'){
    Add-MailboxPermission "Elston.Admin@elston.com.au" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to ElstonAdmin.<br>"
}

$Log += "ElstonEPFS = $ElstonEPFS<br>"
if ($createduser -AND $ElstonEPFS -eq 'True'){
    Add-MailboxPermission "EPFS@elston.com.au" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to ElstonEPFS.<br>"
}

$Log += "ElstonPhilanthropicServices = $ElstonPhilanthropicServices<br>"
if ($createduser -AND $ElstonPhilanthropicServices -eq 'True'){
    Add-MailboxPermission "philanthropy@elston.com.au" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to ElstonPhilanthropicServices.<br>"
}

$Log += "ElstonWealthPartner = $ElstonWealthPartner<br>"
if ($createduser -AND $ElstonWealthPartner -eq 'True'){
    Add-MailboxPermission "Wealth.Partner@elston.com.au" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to ElstonWealthPartner.<br>"
}

$Log += "FundApplications = $FundApplications<br>"
if ($createduser -AND $FundApplications -eq 'True'){
    Add-MailboxPermission "FundApplications@elston.com.au" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to FundApplications.<br>"
}

$Log += "InvestorAccounts = $InvestorAccounts<br>"
if ($createduser -AND $InvestorAccounts -eq 'True'){
    Add-MailboxPermission "InvestorAccounts@elston.com.au" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to InvestorAccounts.<br>"
}

$Log += "ShareRegistryInfo = $ShareRegistryInfo<br>"
if ($createduser -AND $ShareRegistryInfo -eq 'True'){
    Add-MailboxPermission "ShareRegistryInfo@elston.com.au" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to ShareRegistryInfo.<br>"
}

$Log += "UnitPrices = $UnitPrices<br>"
if ($createduser -AND $UnitPrices -eq 'True'){
    Add-MailboxPermission "Unitprices@elston.com.au" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to UnitPrices.<br>"
}

#####DISTRIBUTION LISTS#############

$Log += "Elston Private Wealth (EPW) = $ElstonPrivateWealth<br>"
if ($createduser -AND $ElstonPrivateWealth -eq 'True'){
    $GroupName = "epw@elston.com.au"
    Add-DistributionGroupMember -Identity $GroupName -Member $Email -BypassSecurityGroupManagerCheck
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "EPW Brisbane = $EPWBrisbane<br>"
if ($createduser -AND $EPWBrisbane -eq 'True'){
    $GroupName = "EPW.Brisbane@elston.com.au"
    Add-DistributionGroupMember -Identity $GroupName -Member $Email -BypassSecurityGroupManagerCheck
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "Marketing = $Marketing<br>"
if ($createduser -AND $Marketing -eq 'True'){
    $GroupName = "Marketing@elston.com.au"
    Add-DistributionGroupMember -Identity $GroupName -Member $Email -BypassSecurityGroupManagerCheck
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "Elston ASO's = $ElstonASOs<br>"
if ($createduser -AND $ElstonASOs -eq 'True'){
    $GroupName = "Elston ASO's"
    Add-DistributionGroupMember -Identity $GroupName -Member $Email -BypassSecurityGroupManagerCheck
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "Elston Associate Advisers = $ElstonAssociateAdvisers<br>"
if ($createduser -AND $ElstonAssociateAdvisers -eq 'True'){
    $GroupName = "AssociateAdvisers@elston.com.au"
    Add-DistributionGroupMember -Identity $GroupName -Member $Email -BypassSecurityGroupManagerCheck
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "EAM = $EAM<br>"
if ($createduser -AND $EAM -eq 'True'){
    $GroupName = "EAM@elston.com.au"
    Add-DistributionGroupMember -Identity $GroupName -Member $Email -BypassSecurityGroupManagerCheck
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}


#####MAIL ENABLED SECURITY GROUPS#####
$Log += "Elston Global = $ElstonGlobal<br>"
if ($createduser -AND $ElstonGlobal -eq 'True'){
    $GroupName = "Elston_Global@elston.com.au"
    Add-DistributionGroupMember -Identity $GroupName -Member $Email -BypassSecurityGroupManagerCheck
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "Elston OOL = $ElstonOOL<br>"
if ($createduser -AND $ElstonOOL -eq 'True'){
    $GroupName = "Elston_OOL@elston.com.au"
    Add-DistributionGroupMember -Identity $GroupName -Member $Email -BypassSecurityGroupManagerCheck
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "Elston BNK = $ElstonBAL<br>"
if ($createduser -AND $ElstonBAL -eq 'True'){
    $GroupName = "Elston_BNK@elston.com.au"
    Add-DistributionGroupMember -Identity $GroupName -Member $Email -BypassSecurityGroupManagerCheck
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "Elston CBR = $ElstonCBR<br>"
if ($createduser -AND $ElstonCBR -eq 'True'){
    $GroupName = "Elston_CBR@elston.com.au"
    Add-DistributionGroupMember -Identity $GroupName -Member $Email -BypassSecurityGroupManagerCheck
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "Elston HVB = $ElstonHVB<br>"
if ($createduser -AND $ElstonHVB -eq 'True'){
    $GroupName = "Elston_HVB@elston.com.au"
    Add-DistributionGroupMember -Identity $GroupName -Member $Email -BypassSecurityGroupManagerCheck
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "Elston BNE = $ElstonBNE<br>"
if ($createduser -AND $ElstonBNE -eq 'True'){
    $GroupName = "Elston_BNE@elston.com.au"
    Add-DistributionGroupMember -Identity $GroupName -Member $Email -BypassSecurityGroupManagerCheck
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
Remove-PSSession $ExchangeSession

$json = @"
{
"OfficePhone": "$OfficePhone",
"Log":"$Log"
}
"@
Write-Output $json
