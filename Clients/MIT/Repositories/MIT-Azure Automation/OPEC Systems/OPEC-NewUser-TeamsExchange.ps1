param(
    #2
    [string]$Site='',
    [string]$Email='',

#Email Distribution Lists 12
    [string]$AllAustralia ='',
    [string]$AllMelbourne ='',
    [string]$AllStaff ='',
    [string]$AllSydney ='',
    [string]$AllBrisbane = '',
    [string]$FuelandTank ='',
    [string]$BLBOperations = '',
    [string]$DefenceForwarding = '',
    [string]$Marine = '',
    [string]$Subsea = '',
    [string]$Training = '',
    [string]$VesselHire = ''
)
$Log = ""

# Get certificate thumbprint and App ID
$CertificateThumbprint = Get-AutomationVariable -Name 'EXO-CertificateThumbprint'
$AppId = Get-AutomationVariable -Name 'OPC-AppId'

############### TEAMS CALLING ###############

#Teams Calling Arrays
$Names            = @("BELR1", "HMNT1")
$PhoneBlockLower  = @(61294542510, 61730013710)
$PhoneBlockHigher = @(61294542599, 61730013799)
$PhoneNumberStart = @(612945425,   617300137)
$CallerIDs        = @("Belrose Office", "Hemmant Office")
$DialingPlans     = @("NSW","QLD")
$SiteNumber = [array]::indexof($Names, $Site)
$Log += "Site Index Selected: $SiteNumber - "+$Names[$SiteNumber]+"<br>"

if($SiteNumber -ne -1){
    try{
        $Log += "Connecting to Teams Online<br>"
        $TeamsCredentials = Get-AutomationPSCredential -Name "svc.TeamsAdmin@opecsystems.onmicrosoft.com"
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
    
}else{
    $Log += "Site does not have Teams Calling enabled. Returning phone number 0"
    $OfficePhone = 0
}



$Log += "Connecting to Exchange Online"
Connect-ExchangeOnline -CertificateThumbprint $CertificateThumbprint -AppId $AppId -Organization opecsystems.onmicrosoft.com

#####DISTRIBUTION LISTS#############

$Log += "AllAustralia = $AllAustralia<br>"
if ($AllAustralia -eq 'True'){
    $GroupName = "!AllAustralia@opecsystems.com"
    Add-DistributionGroupMember -Identity $GroupName -Member $Email -BypassSecurityGroupManagerCheck
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "AllStaff = $AllStaff<br>"
if ($AllStaff -eq 'True'){
    $GroupName = "!AllStaff@opecsystems.com"
    Add-DistributionGroupMember -Identity $GroupName -Member $Email -BypassSecurityGroupManagerCheck
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "AllMelbourne = $AllMelbourne<br>"
if ($AllMelbourne -eq 'True'){
    $GroupName = "!AllMelbourne@opecsystems.com"
    Add-DistributionGroupMember -Identity $GroupName -Member $Email -BypassSecurityGroupManagerCheck
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "AllBrisbane = $AllBrisbane<br>"
if ($AllBrisbane -eq 'True'){
    $GroupName = "!AllBrisbane@opecsystems.com"
    Add-DistributionGroupMember -Identity $GroupName -Member $Email -BypassSecurityGroupManagerCheck
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
$Log += "AllStaff = $AllStaff<br>"
if ($AllStaff -eq 'True'){
    $GroupName = "!AllStaff@opecsystems.com"
    Add-DistributionGroupMember -Identity $GroupName -Member $Email -BypassSecurityGroupManagerCheck
    $Log += "The user $Email has been added to the $GroupName group.<br>"
}
Disconnect-ExchangeOnline

$json = @"
{
"OfficePhone": $OfficePhone,
"Log":"$Log"
}
"@
Write-Output $json