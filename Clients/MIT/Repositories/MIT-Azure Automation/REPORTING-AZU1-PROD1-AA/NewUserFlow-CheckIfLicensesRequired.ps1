<#

Mangano IT - New User Flow - Check if Licenses are Required
Created by: Gabriel Nugent
Version: 1.3

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory=$true)][int]$TicketId,
    [string]$TenantUrl,
    [string]$BearerToken,
    [string]$StatusToContinueFlow,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$TicketNoteDetails = ''
$TaskNotes = 'Purchase License'
[string]$Log = ''

# Fetch bearer token if needed
if ($BearerToken -eq '') {
    $Log += "Bearer token not supplied. Getting bearer token for $TenantUrl...`n`n"
    Write-Warning "Bearer token not supplied. Getting bearer token for $TenantUrl..."
	$EncryptedBearerToken = .\MS-GetBearerToken.ps1 -TenantUrl $TenantUrl
    $BearerToken = .\AzAuto-DecryptString.ps1 -String $EncryptedBearerToken
}

## GET CW MANAGE CREDENTIALS ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## FETCH TASKS FROM TICKET ##

$Tasks = .\CWM-FindTicketTasks.ps1 -TicketId $TicketId -TaskNotes $TaskNotes -ApiSecrets $ApiSecrets | ConvertFrom-Json

## CHECK LICENSE AVAILABILITY ##

foreach ($Task in $Tasks) {
    # Clear purchase variable
    $LicensePurchase = $null

    if (!$Task.closedFlag -and $Task.notes -eq $TaskNotes) {
        $Resolution = $Task.resolution | ConvertFrom-Json
        $Name = $Resolution.Name
        $Platform = $Resolution.Platform
        $PlatformId = $Resolution.PlatformId
        $SkuPartNumber = $Resolution.SkuPartNumber
        $BillingTerm = $Resolution.BillingTerm
        $FindLicenseArguments = @{
            SkuPartNumber = $SkuPartNumber
            BearerToken = $BearerToken
        }
        $Operation = .\AAD-FindLicenseDetails.ps1 @FindLicenseArguments | ConvertFrom-Json
        $Log += "`n`n" + $Operation.Log

        # Update task based on result
        if ($null -ne $Operation.Id) {
            # Check if a license is available
            if (!$Operation.LicenseAvailable) {
                # Purchase license if the platform is Pax8
                if ($Platform -eq "Pax8") {
                    $LicenseArguments = @{
                        SubscriptionId = $PlatformId
                        AddSubscriptions = $true
                        Quantity = 1
                        BillingTerm = $BillingTerm
                    }
                    $LicensePurchase = .\Pax8-UpdateSubscription.ps1 @LicenseArguments
                    $Log += $LicensePurchase.Log + "`n`n"
                    $Operation.LicenseAvailable = $LicensePurchase.Result
                }
            }

            # If there is a spare license, close the task
            if ($Operation.LicenseAvailable) {
                $TaskId = $Task.id
                $TaskArguments = @{
                    TicketId = $TicketId
                    TaskId = $TaskId
                    ClosedStatus = $true
                    ApiSecrets = $ApiSecrets
                }
                .\CWM-UpdateSpecificTask.ps1 @TaskArguments | Out-Null
            } else {
                # Add details for ticket note
                $TicketNoteDetails += "`n- $Name"
            }
        } else {
            # Add notes for if a check fails
            $TicketNoteDetails += "`n- $Name not located - please check availability manually."
        }
    }
}

# Set note and status details depending on result
if ($TicketNoteDetails -ne '') {
    $Text = $TaskNotes + " [automated task]`n`n"
    $Text += "The following licenses are required:$TicketNoteDetails`n`n"
    $Text += 'Once the licenses have been purchased, please change the ticket status to "' + $StatusToContinueFlow + '".'
    $Text += "`n`nAfter the status has been updated, this script will run again to make sure those licenses are available. "
    $Text += "If you see this message again, it means that another account has acquired the license you purchased."
    
    $StatusArguments = @{
        TicketId = $TicketId 
        StatusName = "Scheduling Required"
        CustomerRespondedFlag = $true 
        ApiSecrets = $ApiSecrets
    }
    $LicensesRequired = $true
} else {
    $Text = "$TaskNotes [automated task]`n`n"
    $Text += "A check has been performed against the company tenancy, and has confirmed that all required licenses are available. "
    $Text += "The automation will now continue."

    $StatusArguments = @{
        TicketId = $TicketId 
        StatusName = "Automation in Progress"
        CustomerRespondedFlag = $true 
        ApiSecrets = $ApiSecrets
    }
    $LicensesRequired = $false
}

$NoteArguments = @{
    TicketId = $TicketId
    Text = $Text
    ResolutionFlag = $true
    ApiSecrets = $ApiSecrets
}

.\CWM-AddTicketNote.ps1 @NoteArguments | Out-Null
.\CWM-UpdateTicketStatus.ps1 @StatusArguments | Out-Null

## SEND DETAILS TO FLOW ##

$Output = @{
    LicensesRequired = $LicensesRequired
    Log = $Log
}

Write-Output $Output | ConvertTo-Json