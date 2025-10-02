<#

Mangano IT - New User Flow - Create Tickets for New User Hardware
Created by: Gabriel Nugent
Version: 1.6

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory=$true)][int]$TicketId,
    [Parameter(Mandatory=$true)][int]$ContactId,
    [Parameter(Mandatory=$true)][string]$BoardName,
    [Parameter(Mandatory=$true)][string]$GivenName,
    [Parameter(Mandatory=$true)][string]$Surname,
    [Parameter(Mandatory=$true)][string]$StartDate,
    [int]$SiteId,
    [string]$Notes_Computer,
    [string]$Notes_MobileDevice,
    [string]$Notes_AdditionalEquipment,
    [string]$Notes_Deskphone,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

$TicketNoteDetails = ''
$TaskNotes_Computer = 'Make Ticket for Computer'
$TaskNotes_MobileDevice = 'Make Ticket for Mobile Device'
$TaskNotes_AdditionalEquipment = 'Make Ticket for Additional Equipment'
$TaskNotes_Deskphone = 'Make Ticket for Deskphone'
$TicketId_Computer = 0
$TicketId_MobileDevice = 0
$TicketId_AdditionalEquipment = 0
$TicketId_Deskphone = 0
[string]$Log = ''

## GET CW MANAGE CREDENTIALS ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## GET COMPANY ID ##

$CompanyId = .\CWM-FindCompanyId.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets

## FETCH TASKS FROM TICKET ##

$Tasks = .\CWM-FindTicketTasks.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets | ConvertFrom-Json

## CHECK LICENSE ASSIGNED STATUS ##

foreach ($Task in $Tasks) {
    # Clear temporary variables
    $TaskId = 0
    $TaskNotes = ''
    $Summary = " required for setup of $GivenName $Surname - $StartDate"
    $TicketArguments = @{
        InitialInternalNote = 'If a value (e.g. serial number) has not been supplied in the initial note, there will be a blank space instead.'
        CompanyId = $CompanyId
        ContactId = $ContactId
        BoardName = $BoardName
        StatusName = 'Pre-Process'
        TypeName = 'Request'
        PriorityName = 'P4 Normal Response'
        Level = 'Level 1'
    }

    # Add site ID if provided
    if ($SiteId -ne 0) {
        $TicketArguments += @{
            SiteId = $SiteId
        }
    }

    # Set ticket details based on device type
    if (!$Task.closedFlag -and $Task.notes -eq $TaskNotes_Computer) {
        $TaskId = $Task.id
        $TaskNotes = $TaskNotes_Computer
        $Summary = "Computer" + $Summary
        $TicketArguments += @{
            Summary = $Summary
            InitialDescriptionNote = $Notes_Computer
            SubtypeName = 'Laptop / Workstation'
            ItemName = 'SETUP device'
        }
    } elseif (!$Task.closedFlag -and $Task.notes -eq $TaskNotes_MobileDevice) {
        $TaskId = $Task.id
        $TaskNotes = $TaskNotes_MobileDevice
        $Summary = "Mobile device" + $Summary
        $TicketArguments += @{
            Summary = $Summary
            InitialDescriptionNote = $Notes_MobileDevice
            SubtypeName = 'Mobile Device'
            ItemName = 'ADD / SETUP Mobile Device'
        }
    } elseif (!$Task.closedFlag -and $Task.notes -eq $TaskNotes_AdditionalEquipment) {
        $TaskId = $Task.id
        $TaskNotes = $TaskNotes_AdditionalEquipment
        $Summary = "Additional equipment" + $Summary
        $TicketArguments += @{
            Summary = $Summary
            InitialDescriptionNote = $Notes_AdditionalEquipment
            SubtypeName = 'Laptop / Workstation'
            ItemName = 'Peripherals'
        }
    } elseif (!$Task.closedFlag -and $Task.notes -eq $TaskNotes_Deskphone) {
        $TaskId = $Task.id
        $TaskNotes = $TaskNotes_Deskphone
        $Summary = "Deskphone" + $Summary
        $TicketArguments += @{
            Summary = $Summary
            InitialDescriptionNote = $Notes_Deskphone
            SubtypeName = 'Specify New Equipment'
            ItemName = ''
        }
    }

    if ($TaskId -ne 0) {
        $Operation = .\New-CWTicketAndInitialNote.ps1 @TicketArguments
    
        # Close task if license assigned, re-open if not
        if ($null -ne $Operation) {
            # Update output
            switch($Summary) {
                "Computer required for setup of $GivenName $Surname - $StartDate" {
                    $TicketId_Computer = $Operation
                }

                "Mobile device required for setup of $GivenName $Surname - $StartDate" {
                    $TicketId_MobileDevice = $Operation
                }
                "Additional equipment required for setup of $GivenName $Surname - $StartDate" {
                    $TicketId_AdditionalEquipment = $Operation
                }
                "Deskphone required for setup of $GivenName $Surname - $StartDate" {
                    $TicketId_Deskphone = $Operation
                }
            }

            .\CWM-UpdateSpecificTask.ps1 -TicketId $TicketId -TaskId $TaskId -ClosedStatus $true -Resolution "#$Operation" -ApiSecrets $ApiSecrets | Out-Null
            $NoteArguments = @{
                TicketId = $TicketId
                Text = "$TaskNotes [automated task]`n`nA ticket was created for the procurement of equipment: #$Operation"
                ResolutionFlag = $true
                ApiSecrets = $ApiSecrets
            }
            .\CWM-AddTicketNote.ps1 @NoteArguments | Out-Null
            $Log += "SUCCESS: Created ticket #$Operation for task '$TaskNotes'.`n`n"
            Write-Warning "SUCCESS: Created ticket #$Operation for task '$TaskNotes'."
        } else {
            .\CWM-UpdateSpecificTask.ps1 -TicketId $TicketId -TaskId $TaskId -ClosedStatus $false -ApiSecrets $ApiSecrets | Out-Null
            $TicketNoteDetails += "`n- $Summary"
            $Log += "ERROR: Did not create ticket #$Operation for task '$TaskNotes'.`n`n"
            Write-Error "Did not create ticket #$Operation for task '$TaskNotes'."
        }
    }
}

# Add note if tickets not created
if ($TicketNoteDetails -ne '') {
    $NoteArguments = @{
        TicketId = $TicketId
        Text = "Hardware Tickets [automated task]`n`nThe following tickets were not created:$TicketNoteDetails."
        ResolutionFlag = $true
        ApiSecrets = $ApiSecrets
    }
    .\CWM-AddTicketNote.ps1 @NoteArguments | Out-Null
}

## SEND DETAILS TO FLOW ##

$Output = @{
    ComputerTicketId = $TicketId_Computer
    MobileTicketId = $TicketId_MobileDevice
    EquipmentTicketId = $TicketId_AdditionalEquipment
    DeskphoneTicketId = $TicketId_Deskphone
    Log = $Log
}

Write-Output $Output | ConvertTo-Json