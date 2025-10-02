<#

Mangano IT - Export List of Required Groups for OPEC Systems
Created by: Gabriel Nugent
Modified by: Joshua Ceccato
Version: 1.6

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    # Security Groups
    [string]$SecGroup_OSCARDefaultAccess = '',

    #Email Distribution Lists 12
    [string]$DistributionGroup_AllAustralia ='',
    [string]$DistributionGroup_AllMelbourne ='',
    [string]$DistributionGroup_AllStaff ='',
    [string]$DistributionGroup_AllSydney ='',
    [string]$DistributionGroup_AllBrisbane = '',
    [string]$DistributionGroup_EPOCUSA = '',
    [string]$DistributionGroup_FuelandTank ='',
    [string]$DistributionGroup_BLBOperations = '',
    [string]$DistributionGroup_DefenceForwarding = '',
    [string]$DistributionGroup_Marine = '',
    [string]$DistributionGroup_Subsea = '',
    [string]$DistributionGroup_Training = '',
    [string]$DistributionGroup_VesselHire = '',

    # Licenses
    [string]$M365License = '',
    [string]$TCOLicense = ''
)

## ESTABLISH VARIABLES ##

[string]$DistributionGroups = ''
[string]$TeamsChannels = ''
[string]$M365Groups = ''
[string]$Licenses = ''

$DistributionGroup_AllStaff = 'true'

## SECURITY GROUPS ##

if ($SecGroup_OSCARDefaultAccess -eq 'Yes') { $M365Groups += ";SG.SharePoint.Default" }

## DISTRIBUTION GROUPS ##

if ($DistributionGroup_AllAustralia -eq 'true') { $DistributionGroups += ";!AllAustralia@opecsystems.com" }
if ($DistributionGroup_AllStaff -eq 'true') { $DistributionGroups += ";!AllStaff@opecsystems.com" }
if ($DistributionGroup_AllMelbourne -eq 'true') { $DistributionGroups += ";!AllMelbourne@opecsystems.com" }
if ($DistributionGroup_AllBrisbane -eq 'true') { $DistributionGroups += ";!AllBrisbane@opecsystems.com" }
if ($DistributionGroup_EPOCUSA -eq 'true') { $DistributionGroups += ";!EOPCUSA@epocenviro.com" }
if ($DistributionGroup_FuelandTank -eq 'true') { $DistributionGroups += ";!Fuel&Tank@opecsystems.com" }
if ($DistributionGroup_BLBOperations -eq 'true') { $DistributionGroups += ";!AllBLB@opecsystems.com" }
if ($DistributionGroup_DefenceForwarding -eq 'true') { $DistributionGroups += ";defenceforwarding@opecsystems.com" }
if ($DistributionGroup_Marine -eq 'true') { $DistributionGroups += ";marine@opecsystems.com" }
if ($DistributionGroup_Subsea -eq 'true') { $DistributionGroups += ";subsea@opecsystems.com" }
if ($DistributionGroup_Training -eq 'true') { $DistributionGroups += ";Training@opecsystems.com" }
if ($DistributionGroup_VesselHire -eq 'true') { $DistributionGroups += ";vesselhire@opecsystems.com" }

## LICENSES ##
# To add new licenses use the following code split by colons, terminated with a semicolon Name:Platform:PlatformID(ForPax8):MicrosoftSkuPartNumber:BillingTerm;
# If platform is set to Pax8 and the platform ID is set to the Pax8 platform ID for the license, NewUserFlow-CheckIfLicensesRequired.ps1 will attempt to order licenses using Pax8-UpdateSubscription.ps1

if ($M365License -like "Microsoft 365 E5 with Teams*") {
    $M365Groups += ";SG.License.M365E5.NoTeams+TeamsEnterprise"
    $Licenses += "Microsoft 365 E5 (no Teams):Pax8Manual:n/a:Microsoft_365_E5_(no_Teams):Monthly;Microsoft Teams Enterprise:Pax8Manual:n/a:Microsoft_Teams_Enterprise_New:Monthly;"
} elseif ($M365License -like "Microsoft 365 E5 without Teams*") {
    $M365Groups += ";SG.License.M365E5.NoTeams"
    $Licenses += "Microsoft 365 E5 (no Teams):Pax8Manual:n/a:Microsoft_365_E5_(no_Teams):Monthly;"
} elseif ($M365License -like "Microsoft 365 F3*") {
    $M365Groups += ";SG.License.M365F3"
    $Licenses += "Microsoft 365 F3 [New Commerce Experience]:Telstra:n/a:SPE_F1:Monthly;"
}

if ($TCOLicense -eq 'true') {
    $M365Groups += ";SG.License.M365E5.NoTeams+TCO+TeamsEnterprise"
    $Licenses += "Telstra Calling for Office 365:Telstra:n/a:MCOPSTNEAU2:Monthly;Microsoft 365 E5 (no Teams):Pax8Manual:n/a:Microsoft_365_E5_(no_Teams):Monthly;Microsoft Teams Enterprise:Pax8Manual:n/a:Microsoft_Teams_Enterprise_New:Monthly;"
}

## SEND DATA TO FLOW ##

$Output = @{
	DistributionGroups = $DistributionGroups
	TeamsChannels = $TeamsChannels
    M365Groups = $M365Groups
    Licenses = $Licenses
}

Write-Output $Output | ConvertTo-Json