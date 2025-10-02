<#

Mangano IT - DeskDirector - Update Microsoft 365 Group Form
Created by: Gabriel Nugent
Version: 1.0.2

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory)][string]$FormId,
    [Parameter(Mandatory=$true)][string]$TenantUrl,
    [string]$FormQuestionName = 'SharePoint Site/s',
    [string]$BlacklistedGroups
)

## SCRIPT VARIABLES ##

$Result = $false
$DesiredChoices = @()
$EntityId = [int]$FormId

## GET ALL GROUPS ##

# Fetch groups
$M365Groups = .\AAD-GetAllGroups.ps1 -TenantUrl $TenantUrl -GroupsWithOwners $true | ConvertFrom-Json

# Sort through and make array
foreach ($M365Group in $M365Groups) {
    if (!$BlacklistedGroups.Split(',').Contains($M365Group.DisplayName) -and $null -eq $M365Group.membershipRule) {
        $DesiredChoices += $M365Group.DisplayName
    }
}

## UPDATE FORM ##

$Request = .\DeskDirector-UpdateFormQuestionOptions.ps1 -EntityId $EntityId -FormQuestionName $FormQuestionName -DesiredChoices $DesiredChoices -DropDown | ConvertFrom-Json
$Result = $Request.Result

## SEND OUTPUT TO FLOW ##

$Output = @{
    Result = $Result
}

Write-Output $Output | ConvertTo-Json