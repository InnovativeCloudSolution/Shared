<#

Mangano IT - DeskDirector - Update Distribution List Form
Created by: Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory)][string]$FormId,
    [Parameter(Mandatory=$true)][string]$TenantUrl,
    [Parameter(Mandatory=$true)][string]$TenantSlug,
    [string]$FormQuestionName = 'Distribution list',
    [string]$BlacklistedAddresses
)

## SCRIPT VARIABLES ##

$Result = $false
$DesiredChoices = @()
$EntityId = [int]$FormId

## GET ALL MAILBOXES ##

# Fetch lists
$DistributionLists = .\EXO-GetAllDistributionGroups.ps1 -TenantUrl $TenantUrl -TenantSlug $TenantSlug

# Sort through and make array
foreach ($DistributionList in $DistributionLists) {
    if (!$BlacklistedAddresses.Split(',').Contains($DistributionList.PrimarySmtpAddress)) {
        $DistributionListName = $DistributionList.DisplayName
        $DistributionListAddress = $DistributionList.PrimarySmtpAddress
        $DesiredChoices += "$DistributionListName ($DistributionListAddress)"
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