Import-Module 'ConnectWiseManageAPI'

## VARIABLES ##
$MerakiContactEmail = 'support-noreply@meraki.com'
$CQLCompanyIdentifier = 'CQL'
$CQLCompanyId = 2791
$CNSCompanyId = 3075
$MITCompanyId = 2
$OPCCompanyId = 2410
$SEACompanyId = 2581
$SNRCompanyId = 2575
$BUPCompanyId = 3391
$OPCCompanyIdentifier = 'OPC'
$HelpDeskMSBoardId = 60
$HelpDeskTSBoardId = 53
$InternalBoardId = 75
$CNSBoardId = 84
$CQLP1P2Priorities = @(
    "P1 Emergency Response"
    "P2 Critical Response"
)
$CQLP1P2Boards = @(
    'ALERTS - Incidents'
    'CQL-HelpDesk (MS)'
    'CQL-HelpDesk (No SLA)'
)

## INTERNAL FUNCTIONS ##
function New-StandardNote {
    param (
        [Parameter(Mandatory)]$Ticket,
        [Parameter(Mandatory)][string]$StandardNoteName,
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][bool]$IsValidTicket
    )

    <#
    .DESCRIPTION
    Adds a given note to a ticket.

    .NOTES
    This function will only be called by other functions within this module.
    It also assumes that the function calling it is checking ticket validity, and building a note.
    When creating another standard note, please follow the examples of existing functions.
    #>

    # Add note to ticket if valid
    if ($IsValidTicket) {
        try {
            New-CWMTicketNote -TicketId $Ticket.id -Text $Text -internalAnalysisFlag $true | Out-Null
            Write-Warning "[#$($Ticket.id)] New-StandardNote: Added standard note '$StandardNoteName' to #$($Ticket.id)."
        }
        catch {
            Write-Error "[#$($Ticket.id)] New-StandardNote: Unable to add standard note '$StandardNoteName' to #$($Ticket.id) : $($_)"
        }
    }
    else {
        Write-Host "[#$($Ticket.id)] New-StandardNote: #$($Ticket.id) does not meet the criteria for '$StandardNoteName'."
    }
}

## STANDARD NOTES ##
function Add-NoteCQLMerakiUpdates {
    param ([Parameter(Mandatory)]$Ticket)

    <#
    .DESCRIPTION
    Adds a note to CleanCo Meraki tickets to point techs to documentation.
    #>

    # Create bool for whether or not the ticket meets the criteria
    $IsValidTicket = $false

    # Define note name
    $NoteName = "CQL - Meraki Firmware Updates"
    $Text = "Meraki updates for CleanCo CAN NOT be applied at their scheduled date and time. They will need to go through an RFC "
    $Text += "process and be set up to apply at an agreed date and time, as defined in the RFC.`n`n"
    $Text += "They should also, wherever possible be rolled into the quarterly ticket - 'CleanCo Scheduled Patching - Meraki'.`n`n"
    $Text += "This refers to the process listed here - https://mits.au.itglue.com/3210579/docs/9909532"
    
    # Check to see if ticket meets criteria (strict to prevent random notes from going out)
    if (($Ticket.summary -like "Scheduled maintenance for*") -and ($Ticket.summary -like "*in organization*")`
            -and ($Ticket.summary -like "*CleanCo Queensland Ltd*") -and ($Ticket.contactEmailAddress -eq $MerakiContactEmail)) {
        $IsValidTicket = $true
    }

    # Add note if required
    New-StandardNote -Ticket $Ticket -StandardNoteName $NoteName -Text $Text -IsValidTicket $IsValidTicket
}

function Add-NoteOPCExtendedHoursSupport {
    param ([Parameter(Mandatory)]$Ticket)

    <#
    .DESCRIPTION
    Adds a note to OPEC tickets from Extended Hours Support users that explains to techs how their extended support works.
    #>
    
    # Create bool for whether or not the ticket meets the criteria
    $IsValidTicket = $false

    # Define note name
    $NoteName = "$OPCCompanyIdentifier - Extended Hours Support"
    $Text = "This user is entitled to receive Extended Hours Support. This means they can call between 7am and 10pm, "
    $Text += "and receive the same level of support as someone calling between business hours.`n`n"
    $Text += "If the user requires support outside of our 7am-6pm business hours, please speak to the On Call technician "
    $Text += "so a time can be scheduled in for the client be contacted from 6pm - 10pm."
    
    # Check to see if ticket meets criteria (strict to prevent random notes from going out)
    if ($Ticket.company.identifier = $OPCCompanyIdentifier) {
        # Get contact details
        $Contact = Get-CWMCompanyContact -id $Ticket.contact.id
        if ($Contact.types.name -contains 'Extended Hours Support') {
            $IsValidTicket = $true
        }
    }

    # Add note if required
    New-StandardNote -Ticket $Ticket -StandardNoteName $NoteName -Text $Text -IsValidTicket $IsValidTicket
}

function Add-NoteNonSupportUser {
    param ([Parameter(Mandatory)]$Ticket)

    <#
    .DESCRIPTION
    Checks to see if the contact is a supported user at their company, then adds a note based on the ticket's company.

    .NOTES
    For Seasons, this function also removes the agreement from the ticket, and adds a note about passwords if that's more appropriate.
    For Sherrin, this function points the tech to the support matrix.
    #>
    
    # Create bool for whether or not the ticket meets the criteria
    $IsValidTicket = $false

    # Define note name
    $NoteName = "Non-Support User"
    $Text = 'N/A'

    # Checks to see if contact meets criteria
    $Contact = Get-CWMCompanyContact -id $Ticket.contact.id
    if ($Contact.types.name -contains 'Non-Support User') {
        $IsValidTicket = $true
        switch ($Ticket.company.id) {
            $SEACompanyId {
                # Define note
                $Text = "The user $($Contact.firstName) is not a supported staff member at Seasons.`n"
                $Text += "As such, the agreement will be removed from this ticket, and the ticket will be moved to HelpDesk (TS).`n`n"
                $Text += "Please ensure that the users license has not been upgraded to Business Premium before proceeding."
                $Text += " If so, remove the 'Non-Support User' tag from their contact and move this ticket back to MS.`n`n"
                $Text += "F3/RSW user support info: https://mits.au.itglue.com/2642483/docs/11610348"

                # Move ticket to HelpDesk (TS)
                Update-TicketBoard -TicketId $Ticket.id -BoardId $HelpDeskTSBoardId

                # Strip agreement from ticket
                Remove-TicketField -TicketId $Ticket.id -Path 'agreement'
            }
            $SNRCompanyId {
                # Define note
                $Text = "The user $($Contact.firstName) is a Basic License User at Sherrin Rentals. "
                $Text += "Please review the nature of the request, and determine if the ticket should be moved to HelpDesk (TS).`n`n"
                $Text += "Support matrix: https://mits.au.itglue.com/1828010/docs/11040840"
            }
            $BUPCompanyId {
                # Define note
                $Text = "The user $($Contact.firstName) is a Frontline User. In theory, they will only have reached out for "
                $Text += "a service that they are supported for. Please handle this ticket no differently.`n`n"
                $Text += "Support matrix: https://mits.au.itglue.com/6301189/docs/12655982"
            }
            Default {
                # Skip adding note if the ticket's company isn't featured
                $IsValidTicket = $false
            }
        }
    }
    elseif ($Ticket.summary.ToLower() -like "*password*" -and $Ticket.company.id -eq $SEACompanyId) {
        # Add note to Seasons password requests
        $IsValidTicket = $true

        # Define note
        $Text = "Based on the ticket summary, this may be a request for a password reset.`n`n"
        $Text += "Please confirm that the user whose password is being reset is a supported user by checking their ConnectWise "
        $Text += "contact. If they are not, please remove the agreement from this ticket.`n`n"
        $Text += "F3/RSW user support info: https://mits.au.itglue.com/2642483/docs/11610348"
    }

    # Add note if required
    New-StandardNote -Ticket $Ticket -StandardNoteName $NoteName -Text $Text -IsValidTicket $IsValidTicket
}

function Add-NoteCQLP1P2Incidents {
    param ([Parameter(Mandatory)]$Ticket)

    <#
    .DESCRIPTION
    Adds tasks and a note to all CleanCo P1s and P2s.
    #>
    
    # Create bool for whether or not the ticket meets the criteria
    $IsValidTicket = $false

    # Define note name
    $NoteName = "$CQLCompanyIdentifier - P1/P2 Incidents"

    # Define note text
    $Text = @"
Create Tasks [automated action]:

As this ticket has been flagged as a P1/P2, tasks have been added with important additional steps that need to be taken. Please remember to complete them as you go.

Add overall description of issue and resolution:
Once the ticket has been completed, a summary is required that explains what the issue was, and how it was resolved. This allows anyone who reviews the incident after to better understand what happened, especially when there are lots of notes.

Create problem ticket (if required):
Please create a PIR ticket to review the issue. Process - mits.au.itglue.com/3210579/docs/5279493
"@

    # Define tasks
    $Task1 = 'Add overall description of issue and resolution'
    $Task1Description = @"
Once the ticket has been completed, a summary is required that explains what the issue was, and how it was resolved.

This allows anyone who reviews the incident after to better understand what happened, especially when there are lots of notes.
"@
    $Task2 = 'Create problem ticket (if required)'
    $Task2Description = @"
Please create a project ticket to review the issue.

Process - mits.au.itglue.com/3210579/docs/5279493
"@
    $Task3 = 'Create PIR ticket'
    $Task3Description = @"
Please create a PIR ticket to review the issue.

Process - mits.au.itglue.com/3210579/docs/5279493
"@
    
    # Check to see if ticket meets criteria (strict to prevent random notes from going out)
    if (($Ticket.company.id -eq $CQLCompanyId) -and ($CQLP1P2Boards -contains $Ticket.board.name) -and `
        ($CQLP1P2Priorities -contains $Ticket.priority.name)) {
        $IsValidTicket = $true

        # Add tasks
        Add-TicketTask -TicketId $Ticket.id -Notes $Task1 -Resolution $Task1Description
        Add-TicketTask -TicketId $Ticket.id -Notes $Task2 -Resolution $Task2Description
        if ($Ticket.priority.name -eq 'P1 Emergency Response') {
            Add-TicketTask -TicketId $Ticket.id -Notes $Task3 -Resolution $Task3Description
        }
    }

    # Add note if required
    New-StandardNote -Ticket $Ticket -StandardNoteName $NoteName -Text $Text -IsValidTicket $IsValidTicket
}

function Add-NoteCQLOverseasAccessRequest {
    param ([Parameter(Mandatory)]$Ticket)

    <#
    .DESCRIPTION
    Adds a note to CleanCo overseas access requests to point techs to documentation.
    #>

    # Create bool for whether or not the ticket meets the criteria
    $IsValidTicket = $false

    # Define note name
    $NoteName = "CQL - Overseas Access Request"
    $Text = "If this request has come from Rich, it is pre-approved."
    $Text += "Otherwise, please ask Akshay for approval, and not the ICT Business Partners.`n`n"
    $Text += "Relevant doco: https://mits.au.itglue.com/3210579/docs/11345463"
    
    # Check to see if ticket meets criteria (strict to prevent random notes from going out)
    if (($Ticket.summary -like "Requested overseas access:*") -and ($Ticket.company.id -eq $CQLCompanyId)) {
        $IsValidTicket = $true
    }

    # Add note if required
    New-StandardNote -Ticket $Ticket -StandardNoteName $NoteName -Text $Text -IsValidTicket $IsValidTicket
}

function Add-NoteThreatLockerRequest {
    param ([Parameter(Mandatory)]$Ticket)

    <#
    .DESCRIPTION
    Adds notes to ThreatLocker access/elevation requests to point techs to documentation.

    .NOTES
    This function also updates the ticket summary to include the device and application name, and auto-triages the ticket.
    #>

    # Create bool for whether or not the ticket meets the criteria
    $IsValidTicket = $false

    # Define note name
    $NoteName = "ThreatLocker - New Application/Elevation Request"
    $Text = "Please consult the documentation for this request: https://mits.au.itglue.com/1018923/docs/11569605"
    
    # Check to see if ticket meets criteria (strict to prevent random notes from going out)
    if (($Ticket.summary -like "*ThreatLocker Application Request*") -or ($Ticket.summary -like "*ThreatLocker Elevation Request*")) {
        $IsValidTicket = $true

        # Build board moving details
        $BoardDetails = @{
            TicketId    = $Ticket.id
            BoardId     = $HelpDeskMSBoardId
            TypeName    = "Request"
            SubtypeName = "ThreatLocker"
            ItemName    = "MANAGE Requests - Application"
        }

        # Only add standard note if it's an application request and not an elevation request
        if ($Ticket.summary -like "*ThreatLocker Elevation Request*") {
            $BoardDetails.ItemName = "MANAGE Requests - Elevation"
        }

        # Prep board details based on company
        if ($Ticket.company.id -eq $MITCompanyId) {
            Write-Warning "[#$($TicketId)] Add-NoteThreatLockerRequest: Ticket belongs to Mangano IT and will be moved to Internal."
            $BoardDetails.BoardId = $InternalBoardId
            $Level = 'Internal Systems'
        }
        elseif ($Ticket.company.id -eq $CNSCompanyId) {
            Write-Warning "[#$($TicketId)] Add-NoteThreatLockerRequest: Ticket belongs to Coronis and will be moved to CNS-StreamlineIT."
            $BoardDetails.BoardId = $CNSBoardId
            $Level = 'Level 2'
        }
        else {
            Write-Warning "[#$($TicketId)] Add-NoteThreatLockerRequest: Ticket belongs to $($Ticket.company.name) and will be moved to HelpDesk (MS)."
            $Level = 'Level 2'
        }

        # Move to correct board
        Update-TicketBoard @BoardDetails

        # Update level, budget hours, etc.
        Update-TicketField -TicketId $Ticket.id -Path 'budgetHours' -Operation 'replace' -Value 0.5
        Update-TicketLevel -TicketId $Ticket.id -Level $Level

        # Pull device from summary
        $DeviceIndex = $Ticket.summary.IndexOf("Request for") + ("Request for ").Length
        $DeviceName = $Ticket.summary.Substring($DeviceIndex, $Ticket.summary.length - $DeviceIndex)

        # Construct new summary
        if ($Ticket.summary -like "*ThreatLocker Elevation Request*") {
            $Summary = "ThreatLocker Elevation Request: $($DeviceName)"
        }
        else {
            $Summary = "ThreatLocker Application Request: $($DeviceName)"
        }

        # Pull app name if it exists in initial note
        $Note = Get-CWMTicketNote -ticketId $Ticket.id -orderBy "id asc" -pageSize 1
        $ApplicationNameIndex = $Note.text.indexOf("Application Name") + ("Application Name: ").Length

        # Carve out application name if it exists
        if ($ApplicationNameIndex -ne (-1 + ("Application Name: ").Length)) {
            $HashIndex = $Note.text.indexOf("Hash")
            $StringLength = $HashIndex - $ApplicationNameIndex
            $ApplicationName = ($Note.text.Substring($ApplicationNameIndex, $StringLength)) -replace "`n", ""

            Write-Host "[#$($TicketId)] Add-NoteThreatLockerRequest: Application name located - $($ApplicationName)"

            # Remove Built-In from name if present
            if ($ApplicationName -like "* \(Built-In\)") {
                $BuiltInIndex = $ApplicationName.indexOf("\(Built-In\)") - 1
                $ApplicationName = $ApplicationName.Substring(0, $BuiltInIndex)
            }

            # Add application to summary
            $Summary += ", $($ApplicationName)"
        }

        # Update ticket summary
        Update-TicketField -TicketId $Ticket.id -Path 'summary' -Operation 'replace' -Value $Summary
    }

    # Add note if required
    New-StandardNote -Ticket $Ticket -StandardNoteName $NoteName -Text $Text -IsValidTicket $IsValidTicket
}

function Add-NoteOPCUnitedStatesUser {
    param ([Parameter(Mandatory)]$Ticket)

    <#
    .DESCRIPTION
    Adds a note for all tickets from OPEC users that are based in the USA to remind techs of their working hours.
    #>

    # Create bool for whether or not the ticket meets the criteria
    $IsValidTicket = $false

    # Define note name
    $NoteName = "OPC - United States User"
    $Text = "This ticket has been raised by a user in the USA. They are only contactable before 9AM AEST.`n`n"
    $Text += "Please keep this in mind and prioritise this ticket accordingly."
    
    # Check to see if ticket meets criteria (strict to prevent random notes from going out)
    if ($Ticket.company.id -eq $OPCCompanyId) {
        # Checks contact to see if they are in the US
        $Contact = Get-CWMCompanyContact -id $Ticket.contact.id
        if ($Contact.site.name -like "*USA*") {
            $IsValidTicket = $true
        }
    }

    # Add note if required
    New-StandardNote -Ticket $Ticket -StandardNoteName $NoteName -Text $Text -IsValidTicket $IsValidTicket
}

function Add-TriageQltyNote {
    param ([Parameter(Mandatory)]$Ticket)

    $TriageNotes = "https://mits.au.itglue.com/1018923/docs/3725809883908158"
    $Value = @([PSCustomObject]@{id = 68; caption = 'Triage Qlty Notes'; type = 'TextArea'; entryMethod = 'EntryField'; numberOfDecimals = 0; value = $TriageNotes })
    Update-TicketField -TicketId $Ticket.id -Path customFields -Operation replace -Value $Value
}

function Add-CrowdStrikeCI {
    param (
        [Parameter(Mandatory)] $Ticket
    )

    if ($Ticket.contact.name -eq "CrowdStrike Falcon Complete Detections") {
        $TicketSummary = $Ticket.summary
        $TicketInitialNote = (Get-CWMTicketNote -parentId $Ticket.id -orderBy "id asc" | Select-Object -First 1).text
        
        $CIpattern = '\b([A-Za-z0-9]{1,6}(?:-[A-Za-z0-9]{1,30}){1,4})(?:[\\_][A-Za-z0-9_]+)?'

        # Extract CIs from TicketSummary and TicketInitialNote
        $CIArray = @(
            [regex]::Matches($TicketSummary, $CIpattern) |
            ForEach-Object { $_.Value.Replace('\', '') }
            [regex]::Matches($TicketInitialNote, $CIpattern) |
            ForEach-Object { $_.Value.Replace('\', '') }
        ) | Sort-Object -Unique

        $CIChecks = Get-CWMTicketConfiguration -parentId $Ticket.id

        if ($null -ne $CIArray -and $CIArray.Count -gt 0) {
            Write-Warning "[#$($Ticket.id)] Add-CrowdStrikeCI: CI Matches found."

            foreach ($CINames in $CIArray) {
                $CheckCondition = "name like '%$CINames%' AND status/name='Active'"
                $CheckCIs = Get-CWMCompanyConfiguration -condition $CheckCondition

                if ($null -eq $CheckCIs) {
                    Write-Warning "[#$($Ticket.id)] Add-CrowdStrikeCI: No matching CI found for $CINames."
                    continue
                }

                foreach ($CheckCI in $CheckCIs) {
                    $IsAlreadyAttached = $CIChecks | Where-Object { $_.id -eq $CheckCI.id }
                    if ($IsAlreadyAttached) {
                        Write-Warning "[#$($Ticket.id)] Add-CrowdStrikeCI: CI $CINames already attached."
                        continue
                    }

                    Write-Warning "[#$($Ticket.id)] Add-CrowdStrikeCI: Adding CI [$($CheckCI.id)] - $($CheckCI.name)"
                    New-CWMTicketConfiguration -TicketID $Ticket.id -deviceIdentifier $CheckCI.id
                }
            }
        }
        else {
            Write-Warning "[#$($Ticket.id)] Add-CrowdStrikeCI: No CI matches found."
        }

        Write-Warning "[#$($Ticket.id)] Add-CrowdStrikeCI: Classifying Board/Type/subType/Item."
        Update-TicketBoard -TicketId $ticket.id -BoardId 63 -TypeName "Incident" -SubtypeName "Security" -ItemName "Suspicious Behaviour"
        Write-Warning "[#$($Ticket.id)] Add-CrowdStrikeCI: Closing ticket."
        Update-TicketField -TicketId $Ticket.id -Path "status/id" -Value 1316
    }
    else {
        Write-Warning "[#$($Ticket.id)] Add-CrowdStrikeCI: Ticket does not meet the criteria."
    }
}