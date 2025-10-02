Import-Module 'ConnectWiseManageAPI'

## VARIABLES ##
$PelotonSupportAddress = 'support@pelotoncyber.com.au'
$MerakiContactEmail = 'support-noreply@meraki.com'
$GenericExitUserForm = 54392
$GenericUserOffboardingForm = 630300
$GenericUserOffboardingFormPre = 630301
$ExitUserForms = @(
    @{
        CompanyId    = 2544    # ENAP
        FormEntityId = $GenericUserOffboardingFormPre
    },
    @{
        CompanyId    = 2641    # Argent
        FormEntityId = 14107
    },
    @{
        CompanyId    = 904     # Elston
        FormEntityId = $GenericUserOffboardingFormPre
    },
    @{
        CompanyId    = 2410    # OPEC
        FormEntityId = $GenericUserOffboardingFormPre
    },
    @{
        CompanyId    = 2581    # Seasons
        FormEntityId = $GenericUserOffboardingFormPre
    }
)

# Company-specific variables
$CleanCoCompanyId = 2791

## FUNCTIONS FOR THE AUTOBUNDLER ##
function Update-ParentThirdPartyCQLTickets {
    param (
        [Parameter(Mandatory = $true)]
        $Ticket
    )
    
    <#
    .DESCRIPTION
    Checks to see if the ticket is from a third party that works with CleanCo.
    If the ticket matches a given ticket number description, it searches for any tickets with matching ticket numbers.
    If one is found, it bundles the new ticket into the existing ticket.

    .NOTES
    This script will not bundle an open ticket into a closed ticket, unless the ticket is within the appropriate window for being re-opened.

    .EXAMPLE
    $Ticket = Get-CWMTicket -Id 12345
    Update-ParentThirdPartyCQLTickets -Ticket $Ticket
    #>
    
    $ThirdPartyEmailDomains = @(
        'manganoit.com.au',
        'mscmobility.com.au',
        'team.telstra.com',
        'paloaltonetworks.com',
        'citrix.com',
        'netapp.com',
        'cleancohelpdesk.freshservice.com',
        'servicenow@cleancoqld.com.au',
        'purple.telstra.com'
    )
    $Patterns = @{
        "INCV"   = "\(INCV[\d]{7}\D+$"
        "INC"    = "\(INC[\d]{7}\D+$"
        "REQ"    = "\(REQ[\d]{7}\D+$"
        "Case#"  = "\[Case#: \d{8,9} \]"
        "Citrix" = "Citrix - (\d{8,9})"
        "NetApp" = "NetApp Log # (\d{10})"
        "CHN"    = "CHN-\d+"
        "CHG"    = "CHG\d{7} \| "
    }
    $SummaryConditions = $null
    $ThirdPartyTicketDetails = $null

    # Validate if the ticket belongs to CleanCo
    if ($Ticket.company.id -eq $CleanCoCompanyId) {
        # Check if the contact email address matches any third-party domains
        $IsThirdPartyEmailAddress = ($null -ne ($ThirdPartyEmailDomains | Where-Object { $Ticket.contactEmailAddress -match $_ }))

        if ($IsThirdPartyEmailAddress) {
            Write-Host "[#$($Ticket.id)] Update-ParentThirdPartyCQLTickets: #$($Ticket.id) was last updated by a third-party. The ticket will now be checked for external ticket numbers."

            # Match ticket summary against patterns dynamically
            foreach ($Key in $Patterns.Keys) {
                if ($Ticket.summary -match $Patterns[$Key]) {
                    $IdentifierIndex = $Ticket.summary.IndexOf($Key)

                    # Handle each case based on the matched key
                    switch ($Key) {
                        "INCV" {
                            $ThirdPartyTicketDetails = $Ticket.summary.Substring($IdentifierIndex, ($Ticket.summary.Length - $IdentifierIndex)) -replace "\W", ""
                            $SummaryConditions = "*($($ThirdPartyTicketDetails))*"
                            Write-Warning "[#$($Ticket.id)] Update-ParentThirdPartyCQLTickets: Matched regex for INCV."
                            break
                        }
                        "INC" {
                            $ThirdPartyTicketDetails = $Ticket.summary.Substring($IdentifierIndex, ($Ticket.summary.Length - $IdentifierIndex)) -replace "\W", ""
                            $SummaryConditions = "*($($ThirdPartyTicketDetails))*"
                            Write-Warning "[#$($Ticket.id)] Update-ParentThirdPartyCQLTickets: Matched regex for INC."
                            break
                        }
                        "REQ" {
                            $ThirdPartyTicketDetails = $Ticket.summary.Substring($IdentifierIndex, ($Ticket.summary.Length - $IdentifierIndex)) -replace "\W", ""
                            $SummaryConditions = "*($($ThirdPartyTicketDetails))*"
                            Write-Warning "[#$($Ticket.id)] Update-ParentThirdPartyCQLTickets: Matched regex for REQ."
                            break
                        }
                        "Case#" {
                            $ThirdPartyTicketDetails = $Ticket.summary.Substring($IdentifierIndex, ($Ticket.summary.Length - $IdentifierIndex)) -replace "\D", ""
                            $SummaryConditions = "*Case#: $($ThirdPartyTicketDetails)*"
                            Write-Warning "[#$($Ticket.id)] Update-ParentThirdPartyCQLTickets: Matched regex for Case#."
                            break
                        }
                        "Citrix" {
                            $ThirdPartyTicketDetails = $Ticket.summary.Substring($IdentifierIndex, ($Ticket.summary.Length - $IdentifierIndex)) -replace "\D", ""
                            $SummaryConditions = "*Citrix* *$($ThirdPartyTicketDetails)*"
                            Write-Warning "[#$($Ticket.id)] Update-ParentThirdPartyCQLTickets: Matched regex for Citrix."
                            break
                        }
                        "NetApp" {
                            $ThirdPartyTicketDetails = $Ticket.summary.Substring($IdentifierIndex, ($Ticket.summary.Length - $IdentifierIndex)) -replace "\D", ""
                            $SummaryConditions = "*NetApp Log # $($ThirdPartyTicketDetails)*"
                            Write-Warning "[#$($Ticket.id)] Update-ParentThirdPartyCQLTickets: Matched regex for NetApp Log."
                            break
                        }
                        "CHN" {
                            $TicketNumberIndex = $Ticket.summary.IndexOf("Ticket#")
                            if ($TicketNumberIndex -eq -1) {
                                $ThirdPartyTicketDetails = $Ticket.summary.Substring($IdentifierIndex, ($Ticket.summary.Length - $IdentifierIndex)) -replace "\D", ""
                            }
                            else {
                                $ThirdPartyTicketDetails = $Ticket.summary.Substring($IdentifierIndex, ($TicketNumberIndex - $IdentifierIndex)) -replace "\D", ""
                            }
                            $SummaryConditions = "*#CHN-$($ThirdPartyTicketDetails)*"
                            Write-Warning "[#$($Ticket.id)] Update-ParentThirdPartyCQLTickets: Matched regex for CHN."
                            break
                        }
                        "CHG" {
                            $IdentifierIndex = $Ticket.summary.IndexOf('|')
                            $ThirdPartyTicketDetails = $Ticket.summary.Substring(0, $IdentifierIndex) -replace "\D", ""
                            $SummaryConditions = "CHG$($ThirdPartyTicketDetails) |*"
                            Write-Warning "[#$($Ticket.id)] Update-ParentThirdPartyCQLTickets: Matched regex for CHG."
                            break
                        }
                    }

                    # Exit the loop after the first match
                    break
                }
            }

            # Exit early if no SummaryConditions are found
            if (-not $SummaryConditions) {
                Write-Host "[#$($Ticket.id)] Update-ParentThirdPartyCQLTickets: #$($Ticket.id) does not contain any matched third-party ticket numbers in its summary."
                return
            }

            # Find parent ticket
            $Conditions = "summary like '$SummaryConditions' AND parentTicketId = null"
            Write-Host "[#$($Ticket.id)] Update-ParentThirdPartyCQLTickets: Completing search [$Conditions]..."
            $ParentTicket = Get-CWMTicket -condition $Conditions

            # Bundle new ticket into parent ticket (if one exists and isn't closed)
            if ($null -ne $ParentTicket[0].id) {
                Write-Warning "[#$($Ticket.id)] Update-ParentThirdPartyCQLTickets: Matching parent ticket located: #$($ParentTicket.id)."
                Add-ChildTickets -TicketId $ParentTicket[0].id -ChildTicketIds @($Ticket.id)
            }
            else {
                Write-Host "[#$($Ticket.id)] Update-ParentThirdPartyCQLTickets: No parent ticket found for #$($Ticket.id)."
            }
        }
        else {
            Write-Host "[#$($Ticket.id)] Update-ParentThirdPartyCQLTickets: #$($Ticket.id) was not updated by a listed third party, and will not be checked for third-party ticket numbers."
        }
    }
    else {
        Write-Host "[#$($Ticket.id)] Update-ParentThirdPartyCQLTickets: #$($Ticket.id) does not meet the required criteria, and will not be checked for third-party ticket numbers."
    }
}

function Update-ParentPelotonTickets {
    param ([Parameter(Mandatory)]$Ticket)

    <#
    .DESCRIPTION
    Bundles new Peloton service tickets into existing Peloton service tickets.

    .NOTES
    This script will not bundle an open ticket into a closed ticket, unless the ticket is within the appropriate window for being re-opened.
    #>

    # Checks if the ticket is a new Peloton ticket (or a reply to an existing one)
    if (($Ticket.contactEmailAddress -eq $PelotonSupportAddress -or $Ticket.contactEmailAddress -like "*@manganoit.com.au") `
            -and $Ticket.summary -like "*Service Ticket #*" -and $Ticket.summary -like "* regarding *") {
        # Create Peloton ticket number
        $PelotonTicketId = $Ticket.summary.Substring(0, $Ticket.summary.IndexOf("regarding")) -replace '[^0-9]', ""
        Write-Warning "[#$($Ticket.id)] Update-ParentPelotonTickets: #$($Ticket.id) is a new Peloton ticket (ID: #$($PelotonTicketId))."

        # Find main ticket
        $ParentTicket = Get-CWMTicket -condition "summary like 'Service Ticket #$PelotonTicketId regarding*' AND parentTicketId = null"

        # Bundle new ticket into parent ticket (if one exists and isn't closed)
        if ($null -ne $ParentTicket[0].id) {
            Add-ChildTickets -TicketId $ParentTicket[0].id -ChildTicketIds @($Ticket.id)
        }
        else {
            Write-Host "[#$($Ticket.id)] Update-ParentPelotonTickets: No parent ticket found for #$($Ticket.id)."
        }
    }
    else {
        Write-Host "[#$($Ticket.id)] Update-ParentPelotonTickets: #$($Ticket.id) is not a new Peloton ticket."
    }
}

function Update-ParentMerakiTickets {
    param ([Parameter(Mandatory)]$Ticket)

    <#
    .DESCRIPTION
    Bundles together Meraki tickets with identical initial notes and summaries.

    .NOTES
    This script will not bundle an open ticket into a closed ticket.
    #>

    # Checks if the ticket is a new Meraki ticket
    if (($Ticket.summary -like "Scheduled maintenance for*") -and ($Ticket.summary -like "*in organization*")`
            -and ($Ticket.contactEmailAddress -eq $MerakiContactEmail)) {
        Write-Warning "[#$($Ticket.id)] Update-ParentMerakiTickets: #$($Ticket.id) is a new Meraki maintenance notification."

        # Look for tickets with the same summary that should match
        $PossibleParentTickets = Get-CWMTicket -condition "summary = '$($Ticket.summary)' AND id != $($Ticket.id) AND parentTicketId = null AND closedFlag = false"

        # Check if any tickets were located
        if ($null -ne $PossibleParentTickets) {
            Write-Warning "[#$($Ticket.id)] Update-ParentMerakiTickets: There are other tickets that may be an appropriate parent ticket for #$($Ticket.id)."

            # Grab initial note on first ticket
            $InitialNote = Get-CWMTicketNote -ticketId $Ticket.id -orderBy "id asc" -pageSize 1

            # Shave off junk text
            $Index = $InitialNote.text.IndexOf("Cisco Meraki Technical Support")
            $InitialText = $InitialNote.text.Substring(0, $Index)

            # Check through each possible parent ticket and compare first notes
            foreach ($PossibleParentTicket in $PossibleParentTickets) {
                $ParentNote = Get-CWMTicketNote -ticketId $PossibleParentTicket.id -orderBy "id asc" -pageSize 1
                $ParentIndex = $ParentNote.text.IndexOf("Cisco Meraki Technical Support")
                $ParentText = $ParentNote.text.Substring(0, $ParentIndex)

                # Bundle ticket if they match
                if ($InitialText -eq $ParentText) {
                    Write-Warning "[#$($Ticket.id)] Update-ParentMerakiTickets: #$($PossibleParentTicket.id) is a match for #$($Ticket.id)."
                    Add-ChildTickets -TicketId $PossibleParentTicket.id -ChildTicketIds @($Ticket.id)
                    break
                }
                else {
                    Write-Host "[#$($Ticket.id)] Update-ParentMerakiTickets: #$($PossibleParentTicket.id) is not a match for #$($Ticket.id)."
                }
            }
        }
        else {
            Write-Host "[#$($Ticket.id)] Update-ParentMerakiTickets: There are no other tickets that may be an appropriate parent ticket for #$($Ticket.id)."
        }
    }
    else {
        Write-Host "[#$($Ticket.id)] Update-ParentMerakiTickets: #$($Ticket.id) is not a new Meraki maintenance notification."
    }
}

function Update-ExitUserTicketSummary {
    param ([Parameter(Mandatory)]$Ticket)

    <#
    .DESCRIPTION
    Checks to see if the incoming ticket is an exit user ticket.
    If it is, it reviews the list of DeskDirector exit user forms, pulls the user details from a matching form, and then cleans up the summary.

    .NOTES
    Exit user tickets will typically present a user dropdown from a list of contacts when selecting the user to exit.
    This leaves the user's email address in the ticket summary, which cuts off the exit date.
    #>

    # Check if ticket meets criteria
    if ($Ticket.summary -like "User Exit*") {
        Write-Warning "[#$($Ticket.id)] Update-ExitUserTicketSummary: #$($Ticket.id) is an exit user ticket with a potentially obtuse summary."

        # Update ticket level
        Update-TicketLevel -TicketId $Ticket.id -Level 'Level 1'

        # Build var for checking company
        foreach ($Form in $ExitUserForms) {
            if ($Ticket.company.id -eq $Form.CompanyId) {
                $FormEntityId = $Form.FormEntityId
                Write-Host "[#$($Ticket.id)] Update-ExitUserTicketSummary: Found specific form for $($Ticket.company.name): $($Form.FormEntityId)"
            }
        }

        # Provide generic form if one not already supplied
        if ($null -eq $FormEntityId) {
            $FormEntityId = $GenericExitUserForm
            Write-Host "[#$($Ticket.id)] Update-ExitUserTicketSummary: No specific form for $($Ticket.company.name). Generic form to be checked."
        }

        # Get related submission
        $FormResult = Get-DeskDirectorFormResultByTicket -TicketId $Ticket.id -FormEntityId $FormEntityId

        # Update summary with submission details
        if ($null -ne $FormResult) {
            Write-Host $FormResult.form.sections
            # Pull contact info out of submission
            foreach ($Section in $FormResult.form.sections) {
                foreach ($Field in $Section.fields) {
                    if ($Field.name -eq "User To Exit") {
                        $ContactId = $Field.choices[0].identifier
                        Write-Warning "[#$($Ticket.id)] Update-ExitUserTicketSummary: Contact ID located - $($ContactId)"
                    }
                    elseif ($Field.name -like "Exit Date*") {
                        $ExitDateAndTime = $Field.value
                        Write-Warning "[#$($Ticket.id)] Update-ExitUserTicketSummary: Exit Date and Time located - $($ExitDateAndTime)"
                    }
                }
            }

            # Construct summary
            if ($null -ne $ContactId -and $null -ne $ExitDateAndTime) {
                $UserToExit = Get-CWMCompanyContact -id $ContactId
                $ExitDateFormatted = Get-Date -Date $ExitDateAndTime -Format "ddd d MMM, yyyy a\t h:mm tt"
                $Summary = "User Exit: $($UserToExit.firstName) $($UserToExit.lastName), departing on $($ExitDateFormatted)"
                Update-TicketField -TicketId $Ticket.id -Path 'summary' -Value $Summary -Operation 'replace'
            }
            else {
                Write-Host "[#$($Ticket.id)] Update-ExitUserTicketSummary: Contact ID or Exit Date and Time equals null. Summary will not be updated."
                Write-Host "[#$($Ticket.id)] Update-ExitUserTicketSummary: Contact ID: $($ContactId)"
                Write-Host "[#$($Ticket.id)] Update-ExitUserTicketSummary: Exit Date and Time: $($ExitDateAndTime)"
            }
        }
        else {
            Write-Host "[#$($Ticket.id)] Update-ExitUserTicketSummary: No valid form result found. Summary will not be updated."
        }
    }
    else {
        Write-Host "[#$($Ticket.id)] Update-ExitUserTicketSummary: #$($Ticket.id) is not an exit user ticket from a matching company."
    }
}

function Update-UserOffboardingTicketSummary {
    param ([Parameter(Mandatory)]$Ticket)
    if ($Ticket.summary -like "User Offboarding*") {
        Write-Warning "[#$($Ticket.id)] Update-UserOffboardingTicketSummary: #$($Ticket.id) is an exit user ticket with a potentially obtuse summary."

        # Update ticket level
        Update-TicketLevel -TicketId $Ticket.id -Level 'Level 1'

        # Build var for checking company
        foreach ($Form in $ExitUserForms) {
            if ($Ticket.company.id -eq $Form.CompanyId) {
                $FormEntityId = $Form.FormEntityId
                Write-Host "[#$($Ticket.id)] Update-UserOffboardingTicketSummary: Found specific form for $($Ticket.company.name): $($Form.FormEntityId)"
            }
        }

        # Provide generic form if one not already supplied
        if ($null -eq $FormEntityId) {
            $FormEntityId = $GenericUserOffboardingForm
            Write-Host "[#$($Ticket.id)] Update-UserOffboardingTicketSummary: No specific form for $($Ticket.company.name). Generic form to be checked."
        }

        # Get related submission
        $FormResult = Get-DeskDirectorFormResultByTicket -TicketId $Ticket.id -FormEntityId $FormEntityId

        # Update summary with submission details
        if ($null -ne $FormResult) {
            Write-Host $FormResult.form.sections
            # Pull contact info out of submission
            foreach ($Section in $FormResult.form.sections) {
                foreach ($Field in $Section.fields) {
                    if ($Field.name -eq "User To Offboard") {
                        $ContactId = $Field.choices[0].identifier
                        Write-Warning "[#$($Ticket.id)] Update-UserOffboardingTicketSummary: Contact ID located - $($ContactId)"
                    }
                    elseif ($Field.name -like "Exit Date*") {
                        $ExitDateAndTime = $Field.value
                        Write-Warning "[#$($Ticket.id)] Update-UserOffboardingTicketSummary: Exit Date and Time located - $($ExitDateAndTime)"
                    }
                }
            }

            # Construct summary
            if ($null -ne $ContactId -and $null -ne $ExitDateAndTime) {
                $UserToExit = Get-CWMCompanyContact -id $ContactId
                $ExitDateFormatted = Get-Date -Date $ExitDateAndTime -Format "ddd d MMM, yyyy a\t h:mm tt"
                $Summary = "User Offboarding: $($UserToExit.firstName) $($UserToExit.lastName), departing on $($ExitDateFormatted)"
                Update-TicketField -TicketId $Ticket.id -Path 'summary' -Value $Summary -Operation 'replace'
            }
            else {
                Write-Host "[#$($Ticket.id)] Update-UserOffboardingTicketSummary: Contact ID or Exit Date and Time equals null. Summary will not be updated."
                Write-Host "[#$($Ticket.id)] Update-UserOffboardingTicketSummary: Contact ID: $($ContactId)"
                Write-Host "[#$($Ticket.id)] Update-UserOffboardingTicketSummary: Exit Date and Time: $($ExitDateAndTime)"
            }
        }
        else {
            Write-Host "[#$($Ticket.id)] Update-UserOffboardingTicketSummary: No valid form result found. Summary will not be updated."
        }
    }
    else {
        Write-Host "[#$($Ticket.id)] Update-UserOffboardingTicketSummary: #$($Ticket.id) is not an exit user ticket from a matching company."
    }
}

function Update-UserOffboardingTicketSummaryPre {
    param ([Parameter(Mandatory)]$Ticket)
    if ($Ticket.summary -like "User Offboarding*") {
        Write-Warning "[#$($Ticket.id)] Update-UserOffboardingTicketSummaryPre: #$($Ticket.id) is an exit user ticket with a potentially obtuse summary."

        # Update ticket level
        Update-TicketLevel -TicketId $Ticket.id -Level 'Level 1'

        # Build var for checking company
        foreach ($Form in $ExitUserForms) {
            if ($Ticket.company.id -eq $Form.CompanyId) {
                $FormEntityId = $Form.FormEntityId
                Write-Host "[#$($Ticket.id)] Update-UserOffboardingTicketSummaryPre: Found specific form for $($Ticket.company.name): $($Form.FormEntityId)"
            }
        }

        # Provide generic form if one not already supplied
        if ($null -eq $FormEntityId) {
            $FormEntityId = $GenericUserOffboardingFormPre
            Write-Host "[#$($Ticket.id)] Update-UserOffboardingTicketSummaryPre: No specific form for $($Ticket.company.name). Generic form to be checked."
        }

        # Get related submission
        $FormResult = Get-DeskDirectorFormResultByTicket -TicketId $Ticket.id -FormEntityId $FormEntityId

        # Update summary with submission details
        if ($null -ne $FormResult) {
            Write-Host $FormResult.form.sections
            # Pull contact info out of submission
            foreach ($Section in $FormResult.form.sections) {
                foreach ($Field in $Section.fields) {
                    if ($Field.name -eq "User To Offboard") {
                        $ContactId = $Field.choices[0].identifier
                        Write-Warning "[#$($Ticket.id)] Update-UserOffboardingTicketSummaryPre: Contact ID located - $($ContactId)"
                    }
                    elseif ($Field.name -like "Exit Date*") {
                        $ExitDateAndTime = $Field.value
                        Write-Warning "[#$($Ticket.id)] Update-UserOffboardingTicketSummaryPre: Exit Date and Time located - $($ExitDateAndTime)"
                    }
                }
            }

            # Construct summary
            if ($null -ne $ContactId -and $null -ne $ExitDateAndTime) {
                $UserToExit = Get-CWMCompanyContact -id $ContactId
                $ExitDateFormatted = Get-Date -Date $ExitDateAndTime -Format "ddd d MMM, yyyy a\t h:mm tt"
                $Summary = "User Offboarding: $($UserToExit.firstName) $($UserToExit.lastName), departing on $($ExitDateFormatted)"
                Update-TicketField -TicketId $Ticket.id -Path 'summary' -Value $Summary -Operation 'replace'
            }
            else {
                Write-Host "[#$($Ticket.id)] Update-UserOffboardingTicketSummaryPre: Contact ID or Exit Date and Time equals null. Summary will not be updated."
                Write-Host "[#$($Ticket.id)] Update-UserOffboardingTicketSummaryPre: Contact ID: $($ContactId)"
                Write-Host "[#$($Ticket.id)] Update-UserOffboardingTicketSummaryPre: Exit Date and Time: $($ExitDateAndTime)"
            }
        }
        else {
            Write-Host "[#$($Ticket.id)] Update-UserOffboardingTicketSummaryPre: No valid form result found. Summary will not be updated."
        }
    }
    else {
        Write-Host "[#$($Ticket.id)] Update-UserOffboardingTicketSummaryPre: #$($Ticket.id) is not an exit user ticket from a matching company."
    }
}