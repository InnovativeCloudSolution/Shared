<#

Mangano IT - Exit User Flow - Find Configs
Created by: Gabriel Nugent
Version: 1.2.2

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param (
    [int]$TicketId,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$TaskNotes = 'Provide Asset List (manual)'

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## FETCH TASKS FROM TICKET ##

$Tasks = .\CWM-FindTicketTasks.ps1 -TicketId $TicketId -TaskNotes $TaskNotes -ApiSecrets $ApiSecrets | ConvertFrom-Json

## FETCH CONFIGS ##

# If task is not complete, fetch configs
foreach ($Task in $Tasks) {
    $Text = "$TaskNotes"
    $ConfigList = ''
    if ($Task.notes -like "$TaskNotes*" -and !$Task.closedFlag) {
        $TaskResolution = $Task.resolution | ConvertFrom-Json
        $CompanyName = $TaskResolution.CompanyName
        $EmailAddress = $TaskResolution.EmailAddress
        $ConfigsParams = @{
            EmailAddress = $EmailAddress
            CompanyName = $CompanyName
            ParentOnly = $true
            ActiveOnly = $true
            ApiSecrets = $ApiSecrets
        }
        $Configs = .\CWM-FindConfigsForContact.ps1 @ConfigsParams | ConvertFrom-Json

        if ($null -ne $Configs) {
            # Organise configs
            foreach ($Config in $Configs) {
                $ConfigList += "`n`n`Name: " + $Config.name
                $ConfigList += "`nType: " + $Config.type.name
                if ($Config.tagNumber -ne '' -and $null -ne $Config.tagNumber) {
                    $ConfigList += "`nAsset Tag: " + $Config.tagNumber
                }
                if ($Config.serialNumber -ne '' -and $null -ne $Config.serialNumber) {
                    $ConfigList += "`nSerial Number: " + $Config.serialNumber
                }
                if ($Config.modelNumber -ne '' -and $null -ne $Config.modelNumber) {
                    $ConfigList += "`nModel Number: " + $Config.modelNumber
                }
                if ($Config.manufacturer.name -ne '' -and $null -ne $Config.manufacturer.name) {
                    $ConfigList += "`nManufacturer: " + $Config.manufacturer.name
                }
            }

            # Prep text for ticket note
            $Text += "`n`nPlease review the following list of assets for any anomalies (devices that don't actually belong to the user, duplicates, etc.) before sending to the Asset Collection Nominee:$ConfigList"

            # Close task
            # $TaskId = $Task.id
            # .\CWM-UpdateSpecificTask.ps1 -TicketId $TicketId -TaskId $TaskId -ClosedStatus $true -ApiSecrets $ApiSecrets | Out-Null
        } else {
            # Prep text for ticket note
            $Text += "`n`nThe user ($EmailAddress) has no configs assigned. Please confirm this manually before closing the task."
        }

        .\CWM-AddTicketNote.ps1 -TicketId $TicketId -Text $Text -InternalFlag $true -ApiSecrets $ApiSecrets | Out-Null
    }
}

## SEND CONFIG LIST BACK TO FLOW ##

Write-Output $ConfigList