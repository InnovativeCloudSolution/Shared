<#

Mangano IT - Export List of Required Groups for Queensland Hydro
Created by: Gabriel Nugent
Version: 1.5

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    # SharePoint sites
    [string]$SharePoint_Administration,
    [string]$SharePoint_Board,
    [string]$SharePoint_Borumba,
    [string]$SharePoint_Executive,
    [string]$SharePoint_Finance,
    [string]$SharePoint_HR,
    [string]$SharePoint_Legal,
    [string]$SharePoint_Marketing,
    [string]$SharePoint_Operations,
    [string]$SharePoint_PayRoll,
    [string]$SharePoint_Pioneer,
    [string]$SharePoint_PMO,
    [string]$SharePoint_SystemsAndIT,

    # Microsoft 365 security groups
    [string]$M365_KadenceBorumba,
    [string]$M365_KadenceCorporate,
    [string]$M365_KadenceCorporateFinance,
    [string]$M365_KadenceCorporateIMT,
    [string]$M365_KadenceCorporateLegal,
    [string]$M365_KadenceCorporateProcurement,
    [string]$M365_KadenceGeneral,
    [string]$M365_KadenceHR,
    [string]$M365_KadencePioneer,
    [string]$M365_KadencePMO,
    [string]$M365_KadenceStakeholderGovRelations,
    [string]$M365_KadenceStrategy
)

## ESTABLISH VARIABLES ##

[string]$M365Groups = ''

## SHAREPOINT SITES ##

if ($SharePoint_Administration -eq 'True') { $M365Groups += ";Administration" }
if ($SharePoint_Board -eq 'True') { $M365Groups += ";Board" }
if ($SharePoint_Borumba -eq 'True') { $M365Groups += ";Borumba" }
if ($SharePoint_Executive -eq 'True') { $M365Groups += ";Executive" }
if ($SharePoint_Finance -eq 'True') { $M365Groups += ";Finance" }
if ($SharePoint_HR -eq 'True') { $M365Groups += ";HR" }
if ($SharePoint_Legal -eq 'True') { $M365Groups += ";Legal" }
if ($SharePoint_Marketing -eq 'True') { $M365Groups += ";Marketing" }
if ($SharePoint_Operations -eq 'True') { $M365Groups += ";Operations" }
if ($SharePoint_PayRoll -eq 'True') { $M365Groups += ";Pay Roll" }
if ($SharePoint_Pioneer -eq 'True') { $M365Groups += ";Pioneer" }
if ($SharePoint_PMO -eq 'True') { $M365Groups += ";PMO" }
if ($SharePoint_SystemsAndIT -eq 'True') { $M365Groups += ";Systems and IT" }

## MICROSOFT 365 GROUPS ##

if ($M365_KadenceBorumba -eq 'True') { $M365Groups += ";sg.app.kadence.Borumba" }
if ($M365_KadenceCorporate -eq 'True') { $M365Groups += ";sg.app.kadence.Corporate" }
if ($M365_KadenceCorporateFinance -eq 'True') { $M365Groups += ";sg.app.kadence.CorporateFinance" }
if ($M365_KadenceCorporateIMT -eq 'True') { $M365Groups += ";sg.app.kadence.CorporateIMT" }
if ($M365_KadenceCorporateLegal -eq 'True') { $M365Groups += ";sg.app.kadence.CorporateLegal" }
if ($M365_KadenceCorporateProcurement -eq 'True') { $M365Groups += ";sg.app.kadence.CorporateProcurement" }
if ($M365_KadenceGeneral -eq 'True') { $M365Groups += ";sg.app.kadence.General" }
if ($M365_KadenceHR -eq 'True') { $M365Groups += ";sg.app.kadence.HR" }
if ($M365_KadencePioneer -eq 'True') { $M365Groups += ";sg.app.kadence.Pioneer" }
if ($M365_KadencePMO -eq 'True') { $M365Groups += ";sg.app.kadence.PMO" }
if ($M365_KadenceStakeholderGovRelations -eq 'True') { $M365Groups += ";sg.app.kadence.StakeholderGovRelations" }
if ($M365_KadenceStrategy -eq 'True') { $M365Groups += ";sg.app.kadence.Strategy" }

## SEND DATA TO FLOW ##

$Output = @{
	M365Groups = $M365Groups
}

Write-Output $Output | ConvertTo-Json