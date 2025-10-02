<#

Mangano IT - Export List of Required Groups for Argent Australia
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    # User details
    [string]$Site,

    # Shared mailboxes
    [string]$SharedMailbox_Marketing,
    [string]$SharedMailbox_Promotions,
    [string]$SharedMailbox_ArgentRetailVicSA,
    [string]$SharedMailbox_ArthausOrders,
    [string]$SharedMailbox_ArthausSales,
    [string]$SharedMailbox_CrystalReports,
    [string]$SharedMailbox_CytrackMail,
    [string]$SharedMailbox_DCCommunications,
    [string]$SharedMailbox_Enquiries,
    [string]$SharedMailbox_Finance,
    [string]$SharedMailbox_ManufacturingEmailPayslips,
    [string]$SharedMailbox_MarketingCalendar,
    [string]$SharedMailbox_Pod,
    [string]$SharedMailbox_SalesBathroomKitchenCC,
    [string]$SharedMailbox_VICCommercial,
    [string]$SharedMailbox_WarehousePayslipEmail,
    [string]$SharedMailbox_WarrantyClaims,

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

## ESTABLISH VARIABLES ##

[string]$SecurityGroups = ''
[string]$PublicFolders = ''
[string]$SharedMailboxes = ''
[string]$DistributionGroups = ''
[string]$MailEnabledSecurityGroups = ''
[string]$TeamsChannels = ''
[string]$SharePointGroups = ''

## ON-PREM DISTRIBUTION GROUPS GROUPS ##

if ($AllArgentStaff -eq 'True') { $SecurityGroups += ";All Argent Staff" }
if ($AllFortitudeGroup -eq 'True') { $SecurityGroups += ";All Fortitude Group" }
if ($AllSalesStaff -eq 'True') { $SecurityGroups += ";All Sales Staff" }
if ($StockAdvice -eq 'True') { $SecurityGroups += ";Stock Advice" }
if ($AllMastercardHolders -eq 'True') { $SecurityGroups += ";All Mastercard Holders" }
if ($AllFortitudeValley -eq 'True') { $SecurityGroups += ";All Fortitude Valley" }
if ($AllMobilePhones -eq 'True') { $SecurityGroups += ";All Mobile Phones" }
if ($AllProjectSales -eq 'True') { $SecurityGroups += ";All Project Sales" }
if ($AllSalesReps -eq 'True') { $SecurityGroups += ";All Sales Representatives" }
if ($AllRetailSales -eq 'True') { $SecurityGroups += ";All Retail Sales" }
if ($AllNationalOffice -eq 'True') { $SecurityGroups += ";All National Office" }
if ($AllKAM -eq 'True') { $SecurityGroups += ";All KAM" }
if ($AllSydneyProjectGroup -eq 'True') { $SecurityGroups += ";All Sydney Project Group" }
if ($AllSydneyOffice -eq 'True') { $SecurityGroups += ";All Sydney Office" }
if ($SSC -eq 'True') { $SecurityGroups += ";SSC" }
if ($AllLogisticsGroup -eq 'True') { $SecurityGroups += ";All Logistics Group" }
if ($AllMelbourneOffice -eq 'True') { $SecurityGroups += ";All Melbourne Office" }
if ($AllCustomerCentral -eq 'True') { $SecurityGroups += ";All Customer-Central" }
if ($AllBrisbaneSalesOffice -eq 'True') { $SecurityGroups += ";All Brisbane Sales Office" }
if ($AllPerthOffice -eq 'True') { $SecurityGroups += ";All Perth Office" }

## SECURITY GROUPS ##

if ($CustomerCentralFolderAccess -eq 'True') { $SecurityGroups += ";Customer Central Folder Access" }
if ($CorporateFolder -eq 'True') { $SecurityGroups += ";Argent Corporate Folder Access" }

"ExecutiveFolder = $ExecutiveFolder<br>"
if ($ExecutiveFolder -eq 'True') {
    $SecurityGroups += ";Argent executive users"
    "The user $Email has been added to ExecutiveFolder.<br>"
}

"FinanceFolder = $FinanceFolder<br>"
if ($FinanceFolder -eq 'True') {
    $SecurityGroups += ";Argent Finance folder access"
    "The user $Email has been added to FinanceFolder.<br>"
}

"ForecastFolder = $ForecastFolder<br>"
if ($ForecastFolder -eq 'True') {
    $SecurityGroups += ";Argent Forecast Folder Access"
    "The user $Email has been added to ForecastFolder.<br>"
}

"HRFolder = $HRFolder<br>"
if ($HRFolder -eq 'True') {
    $SecurityGroups += ";Argent HR Folder Access"
    "The user $Email has been added to HRFolder.<br>"
}

"ITFolder = $ITFolder<br>"
if ($ITFolder -eq 'True') {
    $SecurityGroups += ";Argent IT folder access"
    "The user $Email has been added to ITFolder.<br>"
}

"LTIWorkingsFolder = $LTIWorkingsFolder<br>"
if ($LTIWorkingsFolder -eq 'True') {
    $SecurityGroups += ";Argent LTI Workings Folder Access"
    "The user $Email has been added to LTIWorkingsFolder.<br>"
}

"MarketingFolder = $MarketingFolder<br>"
if ($MarketingFolder -eq 'True') {
    $SecurityGroups += ";Argent Marketing Folder Access"
    "The user $Email has been added to MarketingFolder.<br>"
}

"MarketingFolderReadOnly = $MarketingFolderReadOnly<br>"
if ($MarketingFolderReadOnly -eq 'True') {
    $SecurityGroups += ";Argent Marketing ReadOnly Group"
    "The user $Email has been added to MarketingFolderReadOnly.<br>"
}

if ($AdobeLicensing -eq 'Adobe CC Suite (License required)') {
    $SecurityGroups += ";SG.AVD.FullDesktop"
    $SecurityGroups += ";SG.AVD.AdobeCCUsers"
} elseif ($CreatedUser) {
    "WVDAccess = $WVDAccess<br>"
    if ($WVDAccess -eq 'Full Desktop') {
        $SecurityGroups += ";SG.AVD.FullDesktop"
        "The user $Email has been added to SG.AVD.FullDesktop<br>"

        if ($AdobeLicensing -eq 'Adobe Acrobat Professional DC (License required)') {
            "The user $Email has been added to SG.AVD.AdobeDCUsers.Full<br>"
            $SecurityGroups += ";SG.AVD.AdobeDCUsers.Full"
        }
    }
    
    if ($WVDAccess -eq 'Published Apps') {
        $SecurityGroups += ";SG.AVD.RemoteApp"
        "The user $Email has been added to SG.AVD.RemoteApp.<br>"

        if ($AdobeLicensing -eq 'Adobe Acrobat Professional DC (License required)') {
            "The user $Email has been added to SG.AVD.AdobeDCUsers<br>"
            $SecurityGroups += ";SG.AVD.AdobeDCUsers"
        }
    }
}

if ($CreatedUser) {
	# Add user to site groups
	switch ($Site) {
		"Alexandria" {
			$SecurityGroups += ";SG.Citrix.Desktop.AXDR"
			$SecurityGroups += ";SG.Citrix.Printing.AXDR"
		}
		"Fortitude Valley" {
			$SecurityGroups += ";SG.Citrix.Desktop.FVAL"        
			$SecurityGroups += ";SG.Citrix.Printing.FVAL"        
		}
		"Osbourne Park" {
			$SecurityGroups += ";SG.Citrix.Desktop.OSPK"    
			$SecurityGroups += ";SG.Citrix.Printing.OSPK"    
		}
		"Pinkenba" {
			$SecurityGroups += ";SG.Citrix.Desktop.PKBA"    
			$SecurityGroups += ";SG.Citrix.Printing.PKBA"    
		}
		"South Melbourne" {
			$SecurityGroups += ";SG.Citrix.Desktop.SMEL"
			$SecurityGroups += ";SG.Citrix.Printing.SMEL"
		}
	}

	# Add user to CodeTwo signature groups
	switch ($Department) {
		"Admin" {
			$SecurityGroups += ";SG.User.ArgentGeneralSignature"
		}
		"Arthaus" {
			$SecurityGroups += ";SG.User.ArthausSignature"
		}
		"Clearance Centre" {
			$SecurityGroups += ";SG.User.ClearanceCenterSignature"
		}
		"Customer Central" {
			$SecurityGroups += ";SG.User.CustomerCentralSignature"
		}
		"Marketing" {
			$SecurityGroups += ";SG.User.ArgentGeneralSignature"
		}
		"Projects" {
			$SecurityGroups += ";SG.User.ArgentProjectsSignature"
		}
		"Retail" {
			$SecurityGroups += ";SG.User.ArgentRetailSignature"
		}
		"Warehouse" {
			$SecurityGroups += ";SG.User.CustomerCentralSignature"
		}
		Default {
			$SecurityGroups += ";SG.User.ArgentGeneralSignature"
		}
	}

	# Add user to generic groups
    $SecurityGroups += ";SG.AllArgentStaff"
    $SecurityGroups += ";SG.Access.CorpWifi"
    $SecurityGroups += ";SG.Access.DataDriveStandard"
	$SecurityGroups += ";SG.User.CodeTwoLicense"
}

## SHARED MAILBOXES ##

if ($SharedMailbox_Marketing -eq 'True') { $SharedMailboxes += ";Marketing@argentaust.com.au:FullAccess" }
if ($SharedMailbox_Promotions -eq 'True') { $SharedMailboxes += ";Promotions@argentaust.com.au:FullAccess" }
if ($SharedMailbox_ArgentRetailVicSA -eq 'True') { $SharedMailboxes += ";retail.victoria@argentaust.com.au:FullAccess" }
if ($SharedMailbox_ArthausOrders -eq 'True') { $SharedMailboxes += ";orders@arthausbk.com.au:FullAccess" }
if ($SharedMailbox_ArthausSales -eq 'True') { $SharedMailboxes += ";sales@arthausbk.com.au:FullAccess" }
if ($SharedMailbox_CrystalReports -eq 'True') { $SharedMailboxes += ";Crystal.Reports@argentaust.com.au:FullAccess" }
if ($SharedMailbox_CytrackMail -eq 'True') { $SharedMailboxes += ";Cytrack.Mail@argentaust.com.au:FullAccess" }
if ($SharedMailbox_DCCommunications -eq 'True') { $SharedMailboxes += ";DC.Communications@argentaust.com.au:FullAccess" }
if ($SharedMailbox_Enquiries -eq 'True') { $SharedMailboxes += ";Enquiries@arthausbk.com.au:FullAccess" }
if ($SharedMailbox_Finance -eq 'True') { $SharedMailboxes += ";Finance@argentaust.com.au:FullAccess" }
if ($SharedMailbox_ManufacturingEmailPayslips -eq 'True') { $SharedMailboxes += ";Manufacturing@argentaust.com.au:FullAccess" }
if ($SharedMailbox_MarketingCalendar -eq 'True') { $SharedMailboxes += ";MarketingCalendar@argentaust.com.au:FullAccess" }
if ($SharedMailbox_Pod -eq 'True') { $SharedMailboxes += ";Pod@arthausbk.com.au:FullAccess" }
if ($SharedMailbox_SalesBathroomKitchenCC -eq 'True') { $SharedMailboxes += ";Sales@bathroomkitchencc.com.au:FullAccess" }
if ($SharedMailbox_VICCommercial -eq 'True') { $SharedMailboxes += ";VIC.Commercial@argentaust.com.au:FullAccess" }
if ($SharedMailbox_WarehousePayslipEmail -eq 'True') { $SharedMailboxes += ";Warehouse@argentaust.com.au:FullAccess" }
if ($SharedMailbox_WarrantyClaims -eq 'True') { $SharedMailboxes += ";Warranty.Claims@argentaust.com.au:FullAccess" }

## SEND DATA TO FLOW ##

$Output = @{
	PublicFolders = $PublicFolders
	SharedMailboxes = $SharedMailboxes
	DistributionGroups = $DistributionGroups
	MailEnabledSecurityGroups = $MailEnabledSecurityGroups
	TeamsChannels = $TeamsChannels
    SharePointGroups = $SharePointGroups
}

Write-Output $Output | ConvertTo-Json