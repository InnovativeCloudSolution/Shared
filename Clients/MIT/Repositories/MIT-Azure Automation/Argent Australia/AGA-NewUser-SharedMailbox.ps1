param(
    #2
    [string]$Site='',
    [string]$Email='',

#Shared Mailboxes 17
    [string]$Marketing='',
    [string]$Promotions='',
    [string]$ArgentRetailVicSA='',
    [string]$ArthausOrders='',
    [string]$ArthausSales='',
    [string]$CrystalReports='',
    [string]$CytrackMail='',
    [string]$DCCommunications='',
    [string]$Enquiries='',
    [string]$Finance='',
    [string]$ManufacturingEmailPayslips='',
    [string]$MarketingCalendar='',
    [string]$Pod='',
    [string]$SalesBathroomKitchenCC='',
    [string]$VICCommercial='',
    [string]$WarehousePayslipEmail='',
    [string]$WarrantyClaims=''
)

$Log=""
$Log += "Connecting to Exchange Online"

# Establish connection variables
$ApplicationId = Get-AutomationVariable -Name 'AGA-EXO-ApplicationId'
$CertificateThumbprint = Get-AutomationVariable -Name 'AGA-EXO-CertificateThumbprint'
$Domain = 'argentaust1.onmicrosoft.com'

Connect-ExchangeOnline -CertificateThumbprint $CertificateThumbprint -AppId $ApplicationId -Organization $Domain

#############SHARED MAILBOXES##############################
$Log += "Marketing = $Marketing<br>"
if ($createduser -AND $Marketing -eq 'True'){
    Add-MailboxPermission "Marketing@argentaust.com.au" -User $Email -AccessRights FullAccess -InheritanceType all
    $Log += "The user $Email has been added to Marketing.<br>"
}

$Log += "Promotions = $Promotions<br>"
if ($createduser -AND $Promotions -eq 'True'){
    Add-MailboxPermission "Promotions@argentaust.com.au" -User $Email -AccessRights FullAccess -InheritanceType all
    $Log += "The user $Email has been added to Promotions.<br>"
}

$Log += "ArgentRetailVicSA = $ArgentRetailVicSA<br>"
if ($createduser -AND $ArgentRetailVicSA -eq 'True'){
    Add-MailboxPermission "retail.victoria@argentaust.com.au" -User $Email -AccessRights FullAccess -InheritanceType all
    $Log += "The user $Email has been added to ArgentRetailVicSA.<br>"
}

$Log += "ArthausOrders = $ArthausOrders<br>"
if ($createduser -AND $ArthausOrders -eq 'True'){
    Add-MailboxPermission "orders@arthausbk.com.au" -User $Email -AccessRights FullAccess -InheritanceType all
    $Log += "The user $Email has been added to ArthausOrders.<br>"
}

$Log += "ArthausSales = $ArthausSales<br>"
if ($createduser -AND $ArthausSales -eq 'True'){
    Add-MailboxPermission "sales@arthausbk.com.au" -User $Email -AccessRights FullAccess -InheritanceType all
    $Log += "The user $Email has been added to ArthausSales.<br>"
}

$Log += "CrystalReports = $CrystalReports<br>"
if ($createduser -AND $CrystalReports -eq 'True'){
    Add-MailboxPermission "Crystal.Reports@argentaust.com.au" -User $Email -AccessRights FullAccess -InheritanceType all
    $Log += "The user $Email has been added to CrystalReports.<br>"
}

$Log += "CytrackMail = $CytrackMail<br>"
if ($createduser -AND $CytrackMail -eq 'True'){
    Add-MailboxPermission "Cytrack.Mail@argentaust.com.au" -User $Email -AccessRights FullAccess -InheritanceType all
    $Log += "The user $Email has been added to CytrackMail.<br>"
}

$Log += "DCCommunications = $DCCommunications<br>"
if ($createduser -AND $DCCommunications -eq 'True'){
    Add-MailboxPermission "DC.Communications@argentaust.com.au" -User $Email -AccessRights FullAccess -InheritanceType all
    $Log += "The user $Email has been added to DCCommunications.<br>"
}

$Log += "Enquiries = $Enquiries<br>"
if ($createduser -AND $Enquiries -eq 'True'){
    Add-MailboxPermission "Enquiries@arthausbk.com.au" -User $Email -AccessRights FullAccess -InheritanceType all
    $Log += "The user $Email has been added to Enquiries.<br>"
}

$Log += "Finance = $Finance<br>"
if ($createduser -AND $Finance -eq 'True'){
    Add-MailboxPermission "Finance@argentaust.com.au" -User $Email -AccessRights FullAccess -InheritanceType all
    $Log += "The user $Email has been added to Finance.<br>"
}

$Log += "ManufacturingEmailPayslips = $ManufacturingEmailPayslips<br>"
if ($createduser -AND $ManufacturingEmailPayslips -eq 'True'){
    Add-MailboxPermission "Manufacturing@argentaust.com.au" -User $Email -AccessRights FullAccess -InheritanceType all
    $Log += "The user $Email has been added to ManufacturingEmailPayslips.<br>"
}

$Log += "MarketingCalendar = $MarketingCalendar<br>"
if ($createduser -AND $MarketingCalendar -eq 'True'){
    Add-MailboxPermission "MarketingCalendar@argentaust.com.au" -User $Email -AccessRights FullAccess -InheritanceType all
    $Log += "The user $Email has been added to MarketingCalendar.<br>"
}

$Log += "Pod = $Pod<br>"
if ($createduser -AND $Pod -eq 'True'){
    Add-MailboxPermission "Pod@arthausbk.com.au" -User $Email -AccessRights FullAccess -InheritanceType all
    $Log += "The user $Email has been added to Pod.<br>"
}

$Log += "SalesBathroomKitchenCC = $SalesBathroomKitchenCC<br>"
if ($createduser -AND $SalesBathroomKitchenCC -eq 'True'){
    Add-MailboxPermission "Sales@bathroomkitchencc.com.au" -User $Email -AccessRights FullAccess -InheritanceType all
    $Log += "The user $Email has been added to SalesBathroomKitchenCC.<br>"
}

$Log += "VICCommercial = $VICCommercial<br>"
if ($createduser -AND $VICCommercial -eq 'True'){
    Add-MailboxPermission "VIC.Commercial@argentaust.com.au" -User $Email -AccessRights FullAccess -InheritanceType all
    $Log += "The user $Email has been added to VICCommercial.<br>"
}

$Log += "WarehousePayslipEmail = $WarehousePayslipEmail<br>"
if ($createduser -AND $WarehousePayslipEmail -eq 'True'){
    Add-MailboxPermission "Warehouse@argentaust.com.au" -User $Email -AccessRights FullAccess -InheritanceType all
    $Log += "The user $Email has been added to WarehousePayslipEmail.<br>"
}

$Log += "WarrantyClaims = $WarrantyClaims<br>"
if ($createduser -AND $WarrantyClaims -eq 'True'){
    Add-MailboxPermission "Warranty.Claims@argentaust.com.au" -User $Email -AccessRights FullAccess -InheritanceType all
    $Log += "The user $Email has been added to WarrantyClaims.<br>"
}

Disconnect-ExchangeOnline -Confirm:$false

Write-Output $Log