param(
    #2
    [string]$Email='',

#Shared Mailboxes 41
    [string]$AppleAdmin ='',
    [string]$Defence ='',
    [string]$DefenceCalendar ='',
    [string]$DefenceIndustrySecurityProgram ='',
    [string]$DefenceInfrastructurePanel ='',
    [string]$EEPCalendar ='',
    [string]$Energy ='',
    [string]$EnergyProjects ='',
    [string]$EnergyRepairsAndMaintenance ='',
    [string]$EnviroBusiness ='',
    [string]$ENVIROCalendar ='',
    [string]$Feedback ='',
    [string]$FuelSamples ='',
    [string]$GeneralManager ='',
    [string]$HealthCalendar ='',
    [string]$Images ='',
    [string]$ISCPurchaseOrders ='',
    [string]$ITAdmin ='',
    [string]$LogisticsAdmin ='',
    [string]$Maintenance ='',
    [string]$MARINECalendar ='',
    [string]$NoReplyMailbox ='',
    [string]$NSWQuotes ='',
    [string]$OPECCareers ='',
    [string]$OPECCBRNeAccounts ='',
    [string]$OPECHR ='',
    [string]$OPECInfo ='',
    [string]$OPECPurchasing ='',
    [string]$OPECReceivables ='',
    [string]$OPECSales ='',
    [string]$Project ='',
    [string]$QFE ='',
    [string]$QLDQuotes ='',
    [string]$Reports ='',
    [string]$SUBSEACalendar ='',
    [string]$Training ='',
    [string]$Transfield ='',
    [string]$VICIndustrialCalendar ='',
    [string]$VICQuotes =''
)

$Log=""
$Log += "Connecting to Exchange Online"

# Get certificate thumbprint and App ID
$CertificateThumbprint = Get-AutomationVariable -Name 'EXO-CertificateThumbprint'
$AppId = Get-AutomationVariable -Name 'OPC-AppId'

Connect-ExchangeOnline -CertificateThumbprint $CertificateThumbprint -AppId $AppId -Organization opecsystems.onmicrosoft.com


#############SHARED MAILBOXES##############################
$Log += "AppleAdmin = $AppleAdmin<br>"
if ($createduser -AND $AppleAdmin -eq 'True'){
    Add-MailboxPermission "AppleAdmin@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to AppleAdmin.<br>"
}

$Log += "Defence = $Defence<br>"
if ($createduser -AND $Defence -eq 'True'){
    Add-MailboxPermission "Defence@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to Defence.<br>"
}

$Log += "DefenceCalendar = $DefenceCalendar<br>"
if ($createduser -AND $DefenceCalendar -eq 'True'){
    Add-MailboxPermission "dcalendar@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to @opecsystems.com.<br>"
}

$Log += "DefenceIndustrySecurityProgram = $DefenceIndustrySecurityProgram<br>"
if ($createduser -AND $DefenceIndustrySecurityProgram -eq 'True'){
    Add-MailboxPermission "disp@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to DefenceIndustrySecurityProgram.<br>"
}

$Log += "DefenceInfrastructurePanel = $DefenceInfrastructurePanel<br>"
if ($createduser -AND $DefenceInfrastructurePanel -eq 'True'){
    Add-MailboxPermission "dip@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to DefenceInfrastructurePanel.<br>"
}

$Log += "EEPCalendar = $EEPCalendar<br>"
if ($createduser -AND $EEPCalendar -eq 'True'){
    Add-MailboxPermission "i2@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to EEPCalendar.<br>"
}

$Log += "Energy = $Energy<br>"
if ($createduser -AND $Energy -eq 'True'){
    Add-MailboxPermission "Energy@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to Energy.<br>"
}

$Log += "EnergyProjects = $EnergyProjects<br>"
if ($createduser -AND $EnergyProjects -eq 'True'){
    Add-MailboxPermission "qldind@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to EnergyProjects.<br>"
}

$Log += "EnergyRepairsAndMaintenance = $EnergyRepairsAndMaintenance<br>"
if ($createduser -AND $EnergyRepairsAndMaintenance -eq 'True'){
    Add-MailboxPermission "ocalendar@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to EnergyRepairsAndMaintenance.<br>"
}

$Log += "EnviroBusiness = $EnviroBusiness<br>"
if ($createduser -AND $EnviroBusiness -eq 'True'){
    Add-MailboxPermission "enviro@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to EnviroBusiness.<br>"
}

$Log += "ENVIROCalendar = $ENVIROCalendar<br>"
if ($createduser -AND $ENVIROCalendar -eq 'True'){
    Add-MailboxPermission "ecalendar@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to ENVIROCalendar.<br>"
}

$Log += "Feedback = $Feedback<br>"
if ($createduser -AND $Feedback -eq 'True'){
    Add-MailboxPermission "Feedback@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to Feedback.<br>"
}

$Log += "FuelSamples = $FuelSamples<br>"
if ($createduser -AND $FuelSamples -eq 'True'){
    Add-MailboxPermission "FuelSamples@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to FuelSamples.<br>"
}

$Log += "GeneralManager = $GeneralManager<br>"
if ($createduser -AND $GeneralManager -eq 'True'){
    Add-MailboxPermission "gm@opeccollege.com.au" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to GeneralManager.<br>"
}

$Log += "HealthCalendar = $HealthCalendar<br>"
if ($createduser -AND $HealthCalendar -eq 'True'){
    Add-MailboxPermission "hcalendar@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to HealthCalendar.<br>"
}

$Log += "Images = $Images<br>"
if ($createduser -AND $Images -eq 'True'){
    Add-MailboxPermission "Images@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to Images.<br>"
}

$Log += "ISCPurchaseOrders = $ISCPurchaseOrders<br>"
if ($createduser -AND $ISCPurchaseOrders -eq 'True'){
    Add-MailboxPermission "ISCPurchaseOrders@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to ISCPurchaseOrders.<br>"
}

$Log += "ITAdmin = $ITAdmin<br>"
if ($createduser -AND $ITAdmin -eq 'True'){
    Add-MailboxPermission "ITAdmin@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to ITAdmin.<br>"
}

$Log += "LogisticsAdmin = $LogisticsAdmin<br>"
if ($createduser -AND $LogisticsAdmin -eq 'True'){
    Add-MailboxPermission "LogAdmin@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to LogisticsAdmin.<br>"
}

$Log += "Maintenance = $Maintenance<br>"
if ($createduser -AND $Maintenance -eq 'True'){
    Add-MailboxPermission "Maintenance@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to Maintenance.<br>"
}

$Log += "MARINECalendar = $MARINECalendar<br>"
if ($createduser -AND $MARINECalendar -eq 'True'){
    Add-MailboxPermission "mcalendar@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to MARINECalendar.<br>"
}

$Log += "NoReplyMailbox = $NoReplyMailbox<br>"
if ($createduser -AND $NoReplyMailbox -eq 'True'){
    Add-MailboxPermission "noreply@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to NoReplyMailbox.<br>"
}

$Log += "NSWQuotes = $NSWQuotes<br>"
if ($createduser -AND $NSWQuotes -eq 'True'){
    Add-MailboxPermission "NSWQuotes@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to NSWQuotes.<br>"
}

$Log += "OPECCareers = $OPECCareers<br>"
if ($createduser -AND $OPECCareers -eq 'True'){
    Add-MailboxPermission "careers@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to OPECCareers.<br>"
}

$Log += "OPECCBRNeAccounts = $OPECCBRNeAccounts<br>"
if ($createduser -AND $OPECCBRNeAccounts -eq 'True'){
    Add-MailboxPermission "accounts@opeccbrne.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to OPECCBRNeAccounts.<br>"
}

$Log += "OPECHR = $OPECHR<br>"
if ($createduser -AND $OPECHR -eq 'True'){
    Add-MailboxPermission "hr@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to OPECHR.<br>"
}

$Log += "OPECInfo = $OPECInfo<br>"
if ($createduser -AND $OPECInfo -eq 'True'){
    Add-MailboxPermission "info@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to OPECInfo.<br>"
}

$Log += "OPECPurchasing = $OPECPurchasing<br>"
if ($createduser -AND $OPECPurchasing -eq 'True'){
    Add-MailboxPermission "purchasing@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to OPECPurchasing.<br>"
}

$Log += "OPECReceivables = $OPECReceivables<br>"
if ($createduser -AND $OPECReceivables -eq 'True'){
    Add-MailboxPermission "receivables@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to OPECReceivables.<br>"
}

$Log += "OPECSales = $OPECSales<br>"
if ($createduser -AND $OPECSales -eq 'True'){
    Add-MailboxPermission "sales@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to OPECSales.<br>"
}

$Log += "Project = $Project<br>"
if ($createduser -AND $Project -eq 'True'){
    Add-MailboxPermission "Project@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to Project.<br>"
}

$Log += "QFE = $QFE<br>"
if ($createduser -AND $QFE -eq 'True'){
    Add-MailboxPermission "QFE@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to QFE.<br>"
}

$Log += "QLDQuotes = $QLDQuotes<br>"
if ($createduser -AND $QLDQuotes -eq 'True'){
    Add-MailboxPermission "QLDQuotes@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to QLDQuotes.<br>"
}

$Log += "Reports = $Reports<br>"
if ($createduser -AND $Reports -eq 'True'){
    Add-MailboxPermission "Reports@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to Reports.<br>"
}

$Log += "SUBSEACalendar = $SUBSEACalendar<br>"
if ($createduser -AND $SUBSEACalendar -eq 'True'){
    Add-MailboxPermission "scalendar@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to SUBSEACalendar.<br>"
}

$Log += "Training = $Training<br>"
if ($createduser -AND $Training -eq 'True'){
    Add-MailboxPermission "Training@opeccollege.edu.au" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to Training.<br>"
}

$Log += "Transfield = $Transfield<br>"
if ($createduser -AND $Transfield -eq 'True'){
    Add-MailboxPermission "Transfield@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to Transfield.<br>"
}

$Log += "VICIndustrialCalendar = $VICIndustrialCalendar<br>"
if ($createduser -AND $VICIndustrialCalendar -eq 'True'){
    Add-MailboxPermission "vicind@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to VICIndustrialCalendar.<br>"
}

$Log += "VICQuotes = $VICQuotes<br>"
if ($createduser -AND $VICQuotes -eq 'True'){
    Add-MailboxPermission "VICQuotes@opecsystems.com" -User $Email -AccessRights FullAccess -InheritanceType all | Out-Null
    $Log += "The user $Email has been added to VICQuotes.<br>"
}
Disconnect-ExchangeOnline

Write-Output $Log
