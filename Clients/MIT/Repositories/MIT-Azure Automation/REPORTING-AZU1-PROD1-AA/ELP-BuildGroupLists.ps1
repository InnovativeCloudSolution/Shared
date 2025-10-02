<#

Mangano IT - Export List of Required Groups for Elston
Created by: Gabriel Nugent
Version: 1.4.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
	# Email distribution lists
    [string]$ElstonPrivateWealth,
    [string]$EPWBrisbane,
    [string]$Marketing,
    [string]$ElstonASOs,
    [string]$ElstonAssociateAdvisers,
    [string]$EAM,

	# Mail-enabled security groups
    [string]$ElstonGlobal,
    [string]$ElstonOOL,
    [string]$ElstonBAL,
    [string]$ElstonCBR,
    [string]$ElstonHVB,
    [string]$ElstonBNE,

	# Public folders
    [string]$Accounts,
    [string]$AssureComms,
    [string]$AssureInfo,
    [string]$CFO,
    [string]$ClientReviews,
    [string]$Comms,
    [string]$ContractNotes,
    [string]$ElstonReports,
    [string]$HR,
    [string]$Hub24,
    [string]$IPSTeam,
    [string]$ITAssist,
    [string]$Payroll,
    [string]$Portfolios,
    [string]$Spam,
    [string]$Succession,
    [string]$UBSAdmin,
    [string]$Voicemail,
    [string]$WebEnquiries,
    [string]$XplanAssist,

	# Shared mailboxes
    [string]$EAMShared,
    [string]$ElstonAdmin,
    [string]$ElstonEPFS,
    [string]$ElstonPhilanthropicServices,
    [string]$ElstonWealthPartner,
    [string]$FundApplications,
    [string]$InvestorAccounts,
    [string]$ShareRegistryInfo,
    [string]$UnitPrices,

	# Teams channels
	[string]$Teams_ASOs,
	[string]$Teams_IPS,

    # SharePoint groups
    [string]$SharePoint,

    # Site and Teams queues
    [string]$Site,
    [string]$IncludeInSiteSupportQueue,
    [string]$IncludeInEAMSupportQueue
)

## ESTABLISH VARIABLES ##

[string]$Microsoft365Groups = ''
[string]$PublicFolders = ''
[string]$SharedMailboxes = ''
[string]$DistributionGroups = ''
[string]$MailEnabledSecurityGroups = ''
[string]$TeamsChannels = ''
[string]$SharePointGroups = ''

## MICROSOFT 365 GROUPS ##

if ($IncludeInSiteSupportQueue -eq 'Include') {
    switch ($Site) {
        'Ballina' { $Microsoft365Groups += ";\teams.ball1.support" }
        'Brisbane' { $Microsoft365Groups += ";\teams.bris1.support" }
        'Canberra' { $Microsoft365Groups += ";\teams.canb1.support" }
        'Gold Coast' { $Microsoft365Groups += ";\teams.gold1.support" }
        'Hervey Bay' { $Microsoft365Groups += ";\teams.herv1.support" }
        'Melbourne' { $Microsoft365Groups += ";\teams.melb1.support" }
        'Sydney' { $Microsoft365Groups += ";\teams.sydn1.support" }
    }
}

if ($IncludeInEAMSupportQueue -eq 'Include') {
    $Microsoft365Groups += ";\teams.eam.support"
}

if ($Site -eq 'VBP - Back Office Solutions') {
    $Microsoft365Groups += ";SG.User.AVD.FullDesktop;SG.Role.PhilippinesUser"
}

## PUBLIC FOLDERS ##

if ($Accounts -eq 'True') { $PublicFolders += ";\Accounts:PublishingEditor" }
if ($AssureComms -eq 'True') { $PublicFolders += ";\Assure Comms:Editor" }
if ($AssureInfo -eq 'True') { $PublicFolders += ";\Assure Info:Editor" }
if ($CFO -eq 'True') { $PublicFolders += ";\CFO:Editor" }
if ($ClientReviews -eq 'True') { $PublicFolders += ";\ClientReviews:Editor" }
if ($Comms -eq 'True') { $PublicFolders += ";\Comms:Reviewer" }
if ($ContractNotes -eq 'True') { $PublicFolders += ";\Contract Notes:Reviewer" }
if ($ElstonReports -eq 'True') { $PublicFolders += ";\Elston Reports:PublishingEditor" }
if ($HR -eq 'True') { $PublicFolders += ";\HR:Editor" }
if ($Hub24 -eq 'True') { $PublicFolders += ";\Hub24:PublishingEditor" }
if ($IPSTeam -eq 'True') { $PublicFolders += ";\IPS Team:PublishingEditor" }
if ($ITAssist -eq 'True') { $PublicFolders += ";\IT Assist:Owner" }
if ($Payroll -eq 'True') { $PublicFolders += ";\Payroll:PublishingEditor" }
if ($Portfolios -eq 'True') { $PublicFolders += ";\Portfolios:Editor" }
if ($Spam -eq 'True') { $PublicFolders += ";\Spam:PublishingEditor" }
if ($Succession -eq 'True') { $PublicFolders += ";\Succession:Editor" }
if ($UBSAdmin -eq 'True') { $PublicFolders += ";\UBS Admin:Editor" }
if ($Voicemail -eq 'True') { $PublicFolders += ";\Voicemail:Owner" }
if ($WebEnquiries -eq 'True') { $PublicFolders += ";\Web Enquiries:Reviewer" }
if ($XplanAssist -eq 'True') { $PublicFolders += ";\Xplan Assist:Owner" }

## SHARED MAILBOXES ##

if ($EAMShared -eq 'True') { $SharedMailboxes += ";EAM.Shared@elston.com.au:FullAccess" }
if ($ElstonAdmin -eq 'True') { $SharedMailboxes += ";Elston.Admin@elston.com.au:FullAccess" }
if ($ElstonEPFS -eq 'True') { $SharedMailboxes += ";EPFS@elston.com.au:FullAccess" }
if ($ElstonPhilanthropicServices -eq 'True') { $SharedMailboxes += ";philanthropy@elston.com.au:FullAccess" }
if ($ElstonWealthPartner -eq 'True') { $SharedMailboxes += ";Wealth.Partner@elston.com.au:FullAccess" }
if ($FundApplications -eq 'True') { $SharedMailboxes += ";FundApplications@elston.com.au:FullAccess" }
if ($InvestorAccounts -eq 'True') { $SharedMailboxes += ";InvestorAccounts@elston.com.au:FullAccess" }
if ($ShareRegistryInfo -eq 'True') { $SharedMailboxes += ";ShareRegistryInfo@elston.com.au:FullAccess" }
if ($UnitPrices -eq 'True') { $SharedMailboxes += ";Unitprices@elston.com.au:FullAccess" }

## DISTRIBUTION LISTS ##

if ($ElstonPrivateWealth -eq 'True') { $DistributionGroups += ";epw@elston.com.au" }
if ($EPWBrisbane -eq 'True') { $DistributionGroups += ";EPW.Brisbane@elston.com.au" }
if ($Marketing -eq 'True') { $DistributionGroups += ";Marketing@elston.com.au" }
if ($ElstonASOs -eq 'True') { $DistributionGroups += ";Elston ASO's" }
if ($ElstonAssociateAdvisers -eq 'True') { $DistributionGroups += ";AssociateAdvisers@elston.com.au" }
if ($EAM -eq 'True') { $DistributionGroups += ";EAM@elston.com.au" }

## MAIL-ENABLED SECURITY GROUPS ##

if ($ElstonGlobal -eq 'True') { $MailEnabledSecurityGroups += ";Elston_Global@elston.com.au" }
if ($ElstonOOL -eq 'True') { $MailEnabledSecurityGroups += ";Elston_OOL@elston.com.au" }
if ($ElstonBAL -eq 'True') { $MailEnabledSecurityGroups += ";Elston_BNK@elston.com.au" }
if ($ElstonCBR -eq 'True') { $MailEnabledSecurityGroups += ";Elston_CBR@elston.com.au" }
if ($ElstonHVB -eq 'True') { $MailEnabledSecurityGroups += ";Elston_HVB@elston.com.au" }
if ($ElstonBNE -eq 'True') { $MailEnabledSecurityGroups += ";Elston_BNE@elston.com.au" }

## TEAMS CHANNELS ##

if ($Teams_ASOs -eq 'True') { $TeamsChannels += ";CSO's" }
if ($Teams_IPS -eq 'True') { $TeamsChannels += ";IPS" }

## SHAREPOINT GROUPS ##

if ($SharePoint -ne '') {
    $SharePoint = $SharePoint.replace(' ', '')
    $SharePointGroups = ';' + ($SharePoint.split('-'))[1]
}

## SEND DATA TO FLOW ##

$Output = @{
    Microsoft365Groups = $Microsoft365Groups
	PublicFolders = $PublicFolders
	SharedMailboxes = $SharedMailboxes
	DistributionGroups = $DistributionGroups
	MailEnabledSecurityGroups = $MailEnabledSecurityGroups
	TeamsChannels = $TeamsChannels
    SharePointGroups = $SharePointGroups
}

Write-Output $Output | ConvertTo-Json