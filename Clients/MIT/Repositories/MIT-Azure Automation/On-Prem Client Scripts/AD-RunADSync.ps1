<#

Mangano IT - Active Directory - Run AD Sync
Created by: Gabriel Nugent
Version: 1.0

This runbook is designed to be used in conjunction with an Azure Automation runbook.

#>

param(
    [string]$AzureADServer
)

## RUN SYNC ##

$SyncResult = Invoke-Command -ComputerName $AzureADServer -ScriptBlock {
    Start-ADSyncSyncCycle -PolicyType Delta
} | Select-Object result

## SEND BACK TO SCRIPT ##

if ($SyncResult.Result -eq "Success") { Write-Output $true } else { Write-Output $false }