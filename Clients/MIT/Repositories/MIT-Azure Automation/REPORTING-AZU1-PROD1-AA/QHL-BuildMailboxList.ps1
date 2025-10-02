<#

Mangano IT - Export List of Required Mailboxes for Queensland Hydro
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
	# Email distribution lists
    [string]$DistributionGroup_QHEmployees,
    [string]$DistributionGroup_AllExtPartners,
    [string]$DistrubtionGroup_239GeorgeAllStaff,
    [string]$DistrubtionGroup_80AnnAllStaff,
    [string]$DistrubtionGroup_BorumbaAllStaff,
    [string]$DistrubtionGroup_PioneerBurdekinAllStaff,

	# Shared mailboxes
    [string]$SharedMailbox_Accounts,
    [string]$SharedMailbox_Board,
    [string]$SharedMailbox_Borumba,
    [string]$SharedMailbox_Careers,
    [string]$SharedMailbox_Communications,
    [string]$SharedMailbox_Community,
    [string]$SharedMailbox_Cybersecurity,
    [string]$SharedMailbox_DocumentControl,
    [string]$SharedMailbox_Docusign,
    [string]$SharedMailbox_Environment,
    [string]$SharedMailbox_Facilities,
    [string]$SharedMailbox_Finance,
    [string]$SharedMailbox_GovernmentRelations,
    [string]$SharedMailbox_IMandT,
    [string]$SharedMailbox_Info,
    [string]$SharedMailbox_Land,
    [string]$SharedMailbox_Media,
    [string]$SharedMailbox_Payroll,
    [string]$SharedMailbox_People,
    [string]$SharedMailbox_PioneerBurdekin,
    [string]$SharedMailbox_PMO,
    [string]$SharedMailbox_Privacy,
    [string]$SharedMailbox_Procurement,
    [string]$SharedMailbox_Steerco,
    [string]$SharedMailbox_Travel,
    [string]$SharedMailbox_WebsiteAnalytics,
    [string]$SharedMailbox_Wellbeing,

    # Shared mailbox access type
    [string]$SharedMailboxPermissions
)

## ESTABLISH VARIABLES ##

[string]$SharedMailboxes = ''
[string]$DistributionGroups = ''

## SHARED MAILBOXES ##

if ($SharedMailbox_Accounts -eq 'True') { $SharedMailboxes += ";accounts@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_Board -eq 'True') { $SharedMailboxes += ";board@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_Borumba -eq 'True') { $SharedMailboxes += ";borumba@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_Careers -eq 'True') { $SharedMailboxes += ";careers@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_Communications -eq 'True') { $SharedMailboxes += ";communications@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_Community -eq 'True') { $SharedMailboxes += ";community@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_Cybersecurity -eq 'True') { $SharedMailboxes += ";cybersecurity@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_DocumentControl -eq 'True') { $SharedMailboxes += ";document.control@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_Docusign -eq 'True') { $SharedMailboxes += ";docusign@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_Environment -eq 'True') { $SharedMailboxes += ";environment@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_Facilities -eq 'True') { $SharedMailboxes += ";facilities@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_Finance -eq 'True') { $SharedMailboxes += ";finance@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_GovernmentRelations -eq 'True') { $SharedMailboxes += ";governmentrelations@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_IMandT -eq 'True') { $SharedMailboxes += ";imt@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_Info -eq 'True') { $SharedMailboxes += ";info@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_Land -eq 'True') { $SharedMailboxes += ";land@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_Media -eq 'True') { $SharedMailboxes += ";media@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_Payroll -eq 'True') { $SharedMailboxes += ";payroll@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_People -eq 'True') { $SharedMailboxes += ";people@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_PioneerBurdekin -eq 'True') { $SharedMailboxes += ";pioneer-burdekin@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_PMO -eq 'True') { $SharedMailboxes += ";pmo@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_Privacy -eq 'True') { $SharedMailboxes += ";privacy@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_Procurement -eq 'True') { $SharedMailboxes += ";procurement@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_Steerco -eq 'True') { $SharedMailboxes += ";steerco@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_Travel -eq 'True') { $SharedMailboxes += ";travel@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_WebsiteAnalytics -eq 'True') { $SharedMailboxes += ";websiteanalytics@qldhydro.com.au:FullAccess" }
if ($SharedMailbox_Wellbeing -eq 'True') { $SharedMailboxes += ";wellbeing@qldhydro.com.au:FullAccess" }

## SHARED MAILBOX SEND AS ##

if ($SharedMailboxPermissions -eq 'True') {
    if ($SharedMailbox_Accounts -eq 'True') { $SharedMailboxes += ";accounts@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Board -eq 'True') { $SharedMailboxes += ";board@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Borumba -eq 'True') { $SharedMailboxes += ";borumba@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Careers -eq 'True') { $SharedMailboxes += ";careers@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Communications -eq 'True') { $SharedMailboxes += ";communications@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Community -eq 'True') { $SharedMailboxes += ";community@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Cybersecurity -eq 'True') { $SharedMailboxes += ";cybersecurity@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_DocumentControl -eq 'True') { $SharedMailboxes += ";document.control@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Docusign -eq 'True') { $SharedMailboxes += ";docusign@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Environment -eq 'True') { $SharedMailboxes += ";environment@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Facilities -eq 'True') { $SharedMailboxes += ";facilities@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Finance -eq 'True') { $SharedMailboxes += ";finance@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_GovernmentRelations -eq 'True') { $SharedMailboxes += ";governmentrelations@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_IMandT -eq 'True') { $SharedMailboxes += ";imt@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Info -eq 'True') { $SharedMailboxes += ";info@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Land -eq 'True') { $SharedMailboxes += ";land@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Media -eq 'True') { $SharedMailboxes += ";media@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Media -eq 'True') { $SharedMailboxes += ";pioneer-burdekin@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Payroll -eq 'True') { $SharedMailboxes += ";payroll@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_People -eq 'True') { $SharedMailboxes += ";people@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_PMO -eq 'True') { $SharedMailboxes += ";pmo@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Privacy -eq 'True') { $SharedMailboxes += ";privacy@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Procurement -eq 'True') { $SharedMailboxes += ";procurement@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Steerco -eq 'True') { $SharedMailboxes += ";steerco@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Travel -eq 'True') { $SharedMailboxes += ";travel@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_WebsiteAnalytics -eq 'True') { $SharedMailboxes += ";websiteanalytics@qldhydro.com.au:SendAs" }
    if ($SharedMailbox_Wellbeing -eq 'True') { $SharedMailboxes += ";wellbeing@qldhydro.com.au:SendAs" }
}

## DISTRIBUTION LISTS ##

if ($DistributionGroup_QHEmployees -eq 'True') { $DistributionGroups += ";QHEmployees@qldhydro.com.au" }
if ($DistributionGroup_AllExtPartners -eq 'True') { $DistributionGroups += ";AllExtPartners@qldhydro.com.au" }
if ($DistrubtionGroup_239GeorgeAllStaff -eq 'True') { $DistributionGroups += ";239GeorgeStAllStaff@qldhydro.com.au" }
if ($DistrubtionGroup_80AnnAllStaff -eq 'True') { $DistributionGroups += ";80AnnStreetAllStaff@qldhydro.com.au" }
if ($DistrubtionGroup_BorumbaAllStaff -eq 'True') { $DistributionGroups += ";BorumbaAllStaff@qldhydro.com.au" }
if ($DistrubtionGroup_PioneerBurdekinAllStaff -eq 'True') { $DistributionGroups += ";PioneerBurdekinAllStaff@qldhydro.com.au" }

## SEND DATA TO FLOW ##

$Output = @{
	SharedMailboxes = $SharedMailboxes
	DistributionGroups = $DistributionGroups
}

Write-Output $Output | ConvertTo-Json