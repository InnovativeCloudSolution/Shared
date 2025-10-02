<#

Mangano IT - Export List of Required Groups for Seasons Living
Created by: Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    #Email Distribution Lists
    [string]$DistributionGroup_CommunityManagers ='',
    [string]$DistributionGroup_CareManagers ='',
    [string]$DistributionGroup_CareTeam ='',
    [string]$DistributionGroup_SupportOffice ='',
    [string]$LicenseType,
    [string]$TeamsCallingRequired
)

## ESTABLISH VARIABLES ##

[string]$DistributionGroups = ''
[string]$M365Groups = ''
[string]$Licenses = ''
[bool]$TeamsCalling = $false
[bool]$NonSupportUser = $false
[string]$Department = 'Seasons'

$DistributionGroup_AllStaff = 'true'

## DISTRIBUTION GROUPS ##

if ($DistributionGroup_AllStaff -eq 'true') { $DistributionGroups += ";AllStaff@seasonsliving.com.au" }
if ($DistributionGroup_CommunityManagers -eq 'true') { $DistributionGroups += ";CommunityManagers@seasonsliving.com.au" }
if ($DistributionGroup_CareManagers -eq 'true') { $DistributionGroups += ";CareManagers@seasonsliving.com.au" }
if ($DistributionGroup_CareTeam -eq 'true') { $DistributionGroups += ";CareTeam@seasonsliving.com.au" }
if ($DistributionGroup_SupportOffice -eq 'true') { $DistributionGroups += ";SupportOffice@seasonsliving.com.au" }

## LICENSES ##

switch ($LicenseType) {
    'Microsoft 365 F3' {
        $Licenses = 'NCE Microsoft 365 F3:Telstra:null:SPE_F1:Monthly;'
        $M365Groups += 'SG.License.OfficeF3;'
        $Department = 'F3 Seasons'
        $NonSupportUser = $true
    }
    'Microsoft 365 Business Premium' {
        $Licenses = 'NCE Microsoft 365 Business Premium:Telstra:null:SPB:Monthly;'
        if ($TeamsCallingRequired -eq 'User Requires Individual Teams Calling') {
            $Licenses += 'Telstra Calling for Office 365:Telstra:null:MCOPSTNEAU2:Monthly;NCE Microsoft Teams Phone Standard:Telstra:null:MCOEV:Monthly;'
            $M365Groups += 'SG.License.M365BusinessPremium_TCO;'
        } else {
            $M365Groups += 'SG.License.M365BusinessPremium;'
        }
    }
}

## SEND DATA TO FLOW ##

$Output = @{
	DistributionGroups = $DistributionGroups
    M365Groups = $M365Groups
    Licenses = $Licenses
    NonSupportUser = $NonSupportUser
    TeamsCallingRequired = $TeamsCalling
    Department = $Department
}

Write-Output $Output | ConvertTo-Json