<#

Mangano IT - Export List of Required Mailboxes for OPEC Systems
Created by: Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param (
    # Shared Mailboxes 41
    [string]$SharedMailbox_AppleAdmin ='',
    [string]$SharedMailbox_Defence ='',
    [string]$SharedMailbox_DefenceCalendar ='',
    [string]$SharedMailbox_DefenceIndustrySecurityProgram ='',
    [string]$SharedMailbox_DefenceInfrastructurePanel ='',
    [string]$SharedMailbox_EEPCalendar ='',
    [string]$SharedMailbox_Energy ='',
    [string]$SharedMailbox_EnergyProjects ='',
    [string]$SharedMailbox_EnergyRepairsAndMaintenance ='',
    [string]$SharedMailbox_EnviroBusiness ='',
    [string]$SharedMailbox_ENVIROCalendar ='',
    [string]$SharedMailbox_Feedback ='',
    [string]$SharedMailbox_FuelSamples ='',
    [string]$SharedMailbox_GeneralManager ='',
    [string]$SharedMailbox_HealthCalendar ='',
    [string]$SharedMailbox_Images ='',
    [string]$SharedMailbox_ISCPurchaseOrders ='',
    [string]$SharedMailbox_ITAdmin ='',
    [string]$SharedMailbox_LogisticsAdmin ='',
    [string]$SharedMailbox_Maintenance ='',
    [string]$SharedMailbox_MARINECalendar ='',
    [string]$SharedMailbox_NoReplyMailbox ='',
    [string]$SharedMailbox_NSWQuotes ='',
    [string]$SharedMailbox_OPECCareers ='',
    [string]$SharedMailbox_OPECCBRNeAccounts ='',
    [string]$SharedMailbox_OPECHR ='',
    [string]$SharedMailbox_OPECInfo ='',
    [string]$SharedMailbox_OPECPurchasing ='',
    [string]$SharedMailbox_OPECReceivables ='',
    [string]$SharedMailbox_OPECSales ='',
    [string]$SharedMailbox_Project ='',
    [string]$SharedMailbox_QFE ='',
    [string]$SharedMailbox_QLDQuotes ='',
    [string]$SharedMailbox_Reports ='',
    [string]$SharedMailbox_SUBSEACalendar ='',
    [string]$SharedMailbox_Training ='',
    [string]$SharedMailbox_Transfield ='',
    [string]$SharedMailbox_VICIndustrialCalendar ='',
    [string]$SharedMailbox_VICQuotes =''
)

## ESTABLISH VARIABLES ##

[string]$SharedMailboxes = ''

## SHARED MAILBOXES ##

if ($SharedMailbox_AppleAdmin -eq 'true') { $SharedMailboxes += ";AppleAdmin@opecsystems.com:FullAccess" }
if ($SharedMailbox_Defence -eq 'true') { $SharedMailboxes += ";Defence@opecsystems.com:FullAccess" }
if ($SharedMailbox_DefenceCalendar -eq 'true') { $SharedMailboxes += ";dcalendar@opecsystems.com:FullAccess" }
if ($SharedMailbox_DefenceIndustrySecurityProgram -eq 'true') { $SharedMailboxes += ";disp@opecsystems.com:FullAccess" }
if ($SharedMailbox_DefenceInfrastructurePanel -eq 'true') { $SharedMailboxes += ";dip@opecsystems.com:FullAccess" }
if ($SharedMailbox_EEPCalendar -eq 'true') { $SharedMailboxes += ";i2@opecsystems.com:FullAccess" }
if ($SharedMailbox_Energy -eq 'true') { $SharedMailboxes += ";Energy@opecsystems.com:FullAccess" }
if ($SharedMailbox_EnergyProjects -eq 'true') { $SharedMailboxes += ";qldind@opecsystems.com:FullAccess" }
if ($SharedMailbox_EnergyRepairsAndMaintenance -eq 'true') { $SharedMailboxes += ";ocalendar@opecsystems.com:FullAccess" }
if ($SharedMailbox_EnviroBusiness -eq 'true') { $SharedMailboxes += ";enviro@opecsystems.com:FullAccess" }
if ($SharedMailbox_ENVIROCalendar -eq 'true') { $SharedMailboxes += ";ecalendar@opecsystems.com:FullAccess" }
if ($SharedMailbox_Feedback -eq 'true') { $SharedMailboxes += ";Feedback@opecsystems.com:FullAccess" }
if ($SharedMailbox_FuelSamples -eq 'true') { $SharedMailboxes += ";FuelSamples@opecsystems.com:FullAccess" }
if ($SharedMailbox_GeneralManager -eq 'true') { $SharedMailboxes += ";gm@opeccollege.com.au:FullAccess" }
if ($SharedMailbox_HealthCalendar -eq 'true') { $SharedMailboxes += ";hcalendar@opecsystems.com:FullAccess" }
if ($SharedMailbox_Images -eq 'true') { $SharedMailboxes += ";Images@opecsystems.com:FullAccess" }
if ($SharedMailbox_ISCPurchaseOrders -eq 'true') { $SharedMailboxes += ";ISCPurchaseOrders@opecsystems.com:FullAccess" }
if ($SharedMailbox_ITAdmin -eq 'true') { $SharedMailboxes += ";ITAdmin@opecsystems.com:FullAccess" }
if ($SharedMailbox_LogisticsAdmin -eq 'true') { $SharedMailboxes += ";LogAdmin@opecsystems.com:FullAccess" }
if ($SharedMailbox_Maintenance -eq 'true') { $SharedMailboxes += ";Maintenance@opecsystems.com:FullAccess" }
if ($SharedMailbox_MARINECalendar -eq 'true') { $SharedMailboxes += ";mcalendar@opecsystems.com:FullAccess" }
if ($SharedMailbox_NoReplyMailbox -eq 'true') { $SharedMailboxes += ";noreply@opecsystems.com:FullAccess" }
if ($SharedMailbox_NSWQuotes -eq 'true') { $SharedMailboxes += ";NSWQuotes@opecsystems.com:FullAccess" }
if ($SharedMailbox_OPECCareers -eq 'true') { $SharedMailboxes += ";careers@opecsystems.com:FullAccess" }
if ($SharedMailbox_OPECCBRNeAccounts -eq 'true') { $SharedMailboxes += ";accounts@opeccbrne.com:FullAccess" }
if ($SharedMailbox_OPECHR -eq 'true') { $SharedMailboxes += ";hr@opecsystems.com:FullAccess" }
if ($SharedMailbox_OPECInfo -eq 'true') { $SharedMailboxes += ";info@opecsystems.com:FullAccess" }
if ($SharedMailbox_OPECPurchasing -eq 'true') { $SharedMailboxes += ";purchasing@opecsystems.com:FullAccess" }
if ($SharedMailbox_OPECReceivables -eq 'true') { $SharedMailboxes += ";receivables@opecsystems.com:FullAccess" }
if ($SharedMailbox_OPECSales -eq 'true') { $SharedMailboxes += ";sales@opecsystems.com:FullAccess" }
if ($SharedMailbox_Project -eq 'true') { $SharedMailboxes += ";Project@opecsystems.com:FullAccess" }
if ($SharedMailbox_QFE -eq 'true') { $SharedMailboxes += ";QFE@opecsystems.com:FullAccess" }
if ($SharedMailbox_QLDQuotes -eq 'true') { $SharedMailboxes += ";QLDQuotes@opecsystems.com:FullAccess" }
if ($SharedMailbox_Reports -eq 'true') { $SharedMailboxes += ";Reports@opecsystems.com:FullAccess" }
if ($SharedMailbox_SUBSEACalendar -eq 'true') { $SharedMailboxes += ";scalendar@opecsystems.com:FullAccess" }
if ($SharedMailbox_Training -eq 'true') { $SharedMailboxes += ";Training@opeccollege.edu.au:FullAccess" }
if ($SharedMailbox_Transfield -eq 'true') { $SharedMailboxes += ";Transfield@opecsystems.com:FullAccess" }
if ($SharedMailbox_VICIndustrialCalendar -eq 'true') { $SharedMailboxes += ";vicind@opecsystems.com:FullAccess" }
if ($SharedMailbox_VICQuotes -eq 'true') { $SharedMailboxes += ";VICQuotes@opecsystems.com:FullAccess" }

## SEND DATA TO FLOW ##

$Output = @{
    SharedMailboxes = $SharedMailboxes
}

Write-Output $Output | ConvertTo-Json