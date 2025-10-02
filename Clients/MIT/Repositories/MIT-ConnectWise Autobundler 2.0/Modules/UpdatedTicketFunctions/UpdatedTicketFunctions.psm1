Import-Module 'ConnectWiseManageAPI'

## VARIABLES ##
$ChildTicketStatusName = '** Child Ticket **'
$StatusSchedulingRequired = 1309
$NoSLAAgreement = 882
$CleanCoCompanyId = 2791
$PiaBoardId = 80

## FUNCTIONS FOR THE AUTOBUNDLER ##
function Update-ChildTicket {
    param ([Parameter(Mandatory)]$Ticket)

    <#
    .DESCRIPTION
    Checks to see if a ticket is a child ticket.
    If the ticket is a child ticket, it checks to see if the status matches "** Child Ticket **", and if the ticket is on the same board.
    If the status is incorrect, it updates the status. If the board is incorrect, it moves the ticket to the correct board.
    #>
    
    # Check if ticket is a child ticket
    if ($null -ne $Ticket.parentTicketId) {
        Write-Warning "[#$($Ticket.id)] Update-ChildTicket: #$($Ticket.id) is a child ticket."

        # Get parent ticket details
        try {
            $ParentTicket = Get-CWMTicket -TicketID $Ticket.parentTicketId
            Write-Host "[#$($Ticket.id)] Update-ChildTicket: Fetched ticket details for #$($ParentTicket.id)."
        }
        catch {
            Write-Error "[#$($Ticket.id)] Update-ChildTicket: Unable to fetch ticket details for #$($Ticket.parentTicketId) : $($_)"
        }

        # Check to see if child ticket has the same company assigned as the parent ticket
        if ($Ticket.company.id -ne $ParentTicket.company.id) {
            Write-Warning "[#$($Ticket.id)] Update-ChildTicket: #$($Ticket.id) is not assigned to the same company as #$($ParentTicket.id), and must be moved."
            Update-TicketField -TicketId $Ticket.id -Path 'company' -ValueId $ParentTicket.company.id

            # Update contact to match original contact
            if ($null -ne $Ticket.contact.id -and $Ticket.contact.id -ne 0) {
                Write-Warning "[#$($Ticket.id)] Update-ChildTicket: #$($Ticket.id) had a valid contact that must be re-applied."
                Update-TicketField -TicketId $Ticket.id -Path 'contact' -ValueId $Ticket.contact.id
            }
        }
        else {
            Write-Host "[#$($Ticket.id)] Update-ChildTicket: #$($Ticket.id) is assigned to the same company as #$($ParentTicket.id)."
        }

        # Check to see if child ticket is on the same board as the parent ticket
        if ($Ticket.board.id -ne $ParentTicket.board.id) {
            Write-Warning "[#$($Ticket.id)] Update-ChildTicket: #$($Ticket.id) is not on the same board as #$($ParentTicket.id), and must be moved."
            Update-TicketField -TicketId $Ticket.id -Path 'board' -ValueId $ParentTicket.board.id
        }
        else {
            Write-Host "[#$($Ticket.id)] Update-ChildTicket: #$($Ticket.id) is on the same board as #$($ParentTicket.id)."
        }

        # Checks to see if the child ticket has the same type as the parent ticket, and updates if required
        if ($Ticket.type.id -ne $ParentTicket.type.id) {
            Write-Warning "[#$($Ticket.id)] Update-ChildTicket: #$($Ticket.id) does not have the same type as #$($ParentTicket.id), and will need to be updated."
            Update-TicketField -TicketId $Ticket.id -Path 'type' -ValueId $ParentTicket.type.id
        }
        else {
            Write-Host "[#$($Ticket.id)] Update-ChildTicket: #$($Ticket.id)'s type does not need to be updated."
        }

        # Checks to see if the child ticket has the same subtype as the parent ticket, and updates if required
        if ($Ticket.subType.id -ne $ParentTicket.subType.id) {
            Write-Warning "[#$($Ticket.id)] Update-ChildTicket: #$($Ticket.id) does not have the same subtype as #$($ParentTicket.id), and will need to be updated."
            Update-TicketField -TicketId $Ticket.id -Path 'subType' -ValueId $ParentTicket.subType.id
        }
        else {
            Write-Host "[#$($Ticket.id)] Update-ChildTicket: #$($Ticket.id)'s subtype does not need to be updated."
        }

        # Checks to see if the child ticket has the same item as the parent ticket, and updates if required
        if ($Ticket.item.id -ne $ParentTicket.item.id) {
            Write-Warning "[#$($Ticket.id)] Update-ChildTicket: #$($Ticket.id) does not have the same item as #$($ParentTicket.id), and will need to be updated."
            Update-TicketField -TicketId $Ticket.id -Path 'item' -ValueId $ParentTicket.item.id
        }
        else {
            Write-Host "[#$($Ticket.id)] Update-ChildTicket: #$($Ticket.id)'s item does not need to be updated."
        }

        # Check if the ticket has the right status
        if ($Ticket.status.name -eq $ChildTicketStatusName) {
            Write-Host "[#$($Ticket.id)] Update-ChildTicket: #$($Ticket.id)'s status does not need to be updated."
        }
        else {
            Write-Warning "[#$($Ticket.id)] Update-ChildTicket: #$($Ticket.id)'s status is currently set to $($Ticket.status.name), and will need to be updated to '$ChildTicketStatusName'."

            # Locate the right status for the board
            $ChildTicketStatus = Get-CWMBoardStatus -serviceBoardId $ParentTicket.board.id -condition "name = '$ChildTicketStatusName'"

            # Update child ticket's status if located
            if ($null -ne $ChildTicketStatus.id) {
                Write-Host "[#$($Ticket.id)] Update-ChildTicket: Status '$($ChildTicketStatus.name)' ($($ChildTicketStatus.id)) has been fetched."
                Update-TicketField -TicketId $Ticket.id -Path 'status' -ValueId $ChildTicketStatus.id -ValueName $ChildTicketStatus.name
            }
            else {
                Write-Host "[#$($Ticket.id)] Update-ChildTicket: Board $($ParentTicket.board.name) does not have a status named '$ChildTicketStatusName'."
            }
        }

        # Update child ticket's note with Parent Ticket ID
        if ($Ticket.board.name -ne "Sales") {
            $ChildTicketNote = "This issue has already been raised by another member of your organisation and has been grouped accordingly. All related work is now being managed under ticket #$($ParentTicket.id) - $($ParentTicket.summary).`n`nIf you have any further questions, please contact the Service Desk team, quoting ticket #$($ParentTicket.id)."
            New-CWMTicketNote -ticketId $Ticket.id -text $ChildTicketNote -detailDescriptionFlag $true
            Write-Host "[#$($Ticket.id)] Update-ChildTicket: Note 'Parent is Ticket [#$($ParentTicket.id)]' added to Discussion."
        }
        else {
            Write-Host "[#$($Ticket.id)] Update-ChildTicket: Ticket is in Sales Board."
        }

    }
    else {
        Write-Host "[#$($Ticket.id)] Update-ChildTicket: #$($Ticket.id) is not a child ticket."
    }
}

function Update-TicketOwner {
    param ([Parameter(Mandatory)]$Ticket)

    <#
    .DESCRIPTION
    Checks to see if the ticket owner matches the assigned resource.

    .NOTES
    This does not work with tickets that have more than one resource assigned.
    #>
    
    # Check if the ticket only has one resource assigned
    if ($null -eq $Ticket.resources) {
        Write-Host "[#$($Ticket.id)] Update-TicketOwner: #$($Ticket.id) has no resources assigned."
    }
    elseif ($Ticket.resources -like "*,*") {
        Write-Host "[#$($Ticket.id)] Update-TicketOwner: #$($Ticket.id) has more than one resource assigned ($($Ticket.resources))."
    }
    elseif ($Ticket.resources -eq $Ticket.owner.identifier) {
        Write-Host "[#$($Ticket.id)] Update-TicketOwner: Ticket owner for #$($Ticket.id) already matches the single assigned resource ($($Ticket.resources))."
    }
    else {
        Write-Warning "[#$($Ticket.id)] Update-TicketOwner: #$($Ticket.id) only has one resource assigned ($($Ticket.resources)), and they have not been set as the owner."

        # Grab schedules from ticket
        $Schedules = Get-CWMScheduleEntry -condition "objectId = $($Ticket.id) AND doneFlag = false" -pageSize 1

        # Update ticket owner using existing schedule
        if ($null -ne $Schedules) {
            try {
                Update-CWMScheduleEntry -id $Schedules[0].id -Operation 'replace' -Path 'ownerFlag' -Value $true | Out-Null
                Write-Warning "[#$($Ticket.id)] Update-TicketOwner: Changed ticket owner for #$($Ticket.id) to $($Schedules[0].member.identifier)."
            }
            catch {
                Write-Error "[#$($Ticket.id)] Update-TicketOwner: Unable to change ticket owner for #$($Ticket.id) to $($Schedules[0].member.identifier) : $($_)"
            }
        }
        else {
            Write-Host "[#$($Ticket.id)] Update-TicketOwner: No valid schedules returned. Skipping updating of ticket owner."
        }
    }
}

function Send-CleanCoCosolEmail {
    param (
        [Parameter(Mandatory)]$Ticket
    )

    # Constants
    $BCC = "support@manganoit.com.au"
    $CosolEmail = "clnco@cosol.atlassian.net"
    $CosolDomain = "cosol\.atlassian\.net"
    $CQLDomain = "cleancoqld\.com\.au"
    $WebhookUrl = "https://prod-11.australiasoutheast.logic.azure.com:443/workflows/6973220a4a734bd395a11c65047aba87/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=Z2pIM0a81r4yuF4Bm-0Nx2V5U8zqaHs2QosJKN5cdF0"

    # Helper function to get contact email
    function Get-ContactEmail {
        param (
            $ContactId
        )
        $CWMContacts = Get-CWMContact -id $ContactId
        $CommunicationItems = $CWMContacts.communicationItems | Where-Object -Property "communicationType" -eq "email"
        return $CommunicationItems.value
    }

    # Get necessary details
    $TicketCompleteNote = Get-CWMTicketNote -parentId $Ticket.id -condition "(text contains '%COMPLETE%') AND createdby='jira@cosol.atlassian.net'" -orderBy "id desc" | Select-Object -First 1
    $TicketLatestNote = Get-CWMTicketNote -parentId $Ticket.id -orderBy "id desc" | Select-Object -First 1
    $TicketLatestNoteTest = $TicketLatestNote.text -split "`n" | Select-Object -First 1
    
    if ($TicketLatestNote.createdBy -ne "jira@cosol.atlassian.net") {
        $TicketLatestEmail = Get-ContactEmail $TicketLatestNote.contact.id
    }
    else {
        $TicketLatestEmail = $TicketLatestNote.createdBy
    }

    # Check if the ticket meets the criteria
    if (($Ticket.company.id -eq $CleanCoCompanyId) -and `
        ($($Ticket.board.name) -eq "CQL-HelpDesk (No SLA)") -and `
        ($Ticket.summary -match "\[CLNCO-\d{1,6}\]") -and `
        ($Ticket.status.name -match "Client Responded") -and `
        ($TicketLatestNoteTest -ne "Response From") -and `
        ($TicketLatestNoteTest -notlike "*has updated the ticket*")) {

        Write-Warning "[#$($Ticket.id)] Send-CleanCoCosolEmail: CleanCo QLD SAP/SuccessFactor ticket with a status of 'Client Responded'."

        # Determine who responded
        if ($TicketLatestEmail -match $CQLDomain) {
            Write-Warning "[#$($Ticket.id)] Send-CleanCoCosolEmail: Response is from CleanCo."
            $ResponseFrom = "CleanCo - $($Ticket.contactName) - $($Ticket.contactEmailAddress)"
            $TicketNoteUpdate = "$($Ticket.contactName) has updated the ticket."
            $SendTo = $CosolEmail
            $Action = "CQLUpdated"
        }
        elseif ($TicketLatestEmail -match $CosolDomain) {
            Write-Warning "[#$($Ticket.id)] Send-CleanCoCosolEmail: Response is from COSOL."
            $ResponseFrom = "COSOL - $TicketLatestEmail"
            if ($TicketLatestNote.id -eq $TicketCompleteNote.id) {
                Write-Warning "[#$($Ticket.id)] Send-CleanCoCosolEmail: COSOL has updated the ticket to complete."
                $TicketNoteUpdate = "COSOL has updated the ticket to complete."
                $SendTo = $Ticket.contactEmailAddress
                $Action = "COSOLCompleted"
            }
            elseif ($TicketLatestNote.text -match "Accepted") {
                Write-Warning "[#$($Ticket.id)] Send-CleanCoCosolEmail: COSOL has accepted the ticket."
                $TicketNoteUpdate = "COSOL has accepted the ticket."
                $SendTo = $Ticket.contactEmailAddress
                $Action = "COSOLAccepted"
            }
            else {
                Write-Warning "[#$($Ticket.id)] Send-CleanCoCosolEmail: COSOL has updated the ticket."
                $TicketNoteUpdate = "COSOL has updated the ticket."
                $SendTo = $Ticket.contactEmailAddress
                $Action = "COSOLUpdated"
            }
        }

        # Add the ticket note and log the action
        New-CWMTicketNote -ticketId $Ticket.id -text $TicketNoteUpdate -internalAnalysisFlag $true
        Write-Host "[#$($Ticket.id)] Send-CleanCoCosolEmail: Ticket Note '$TicketNoteUpdate' added to Internal."

        # Prepare and send JSON Payload to PowerAutomate
        $Subject = "Ticket # $($Ticket.id) / CQL / $($Ticket.summary)"
        $Payload = @{
            TicketID     = $Ticket.id
            Subject      = $Subject
            Sendto       = $SendTo
            Action       = $Action
            ResponseFrom = $ResponseFrom
            Body         = $TicketLatestNote.text
            BCC          = $BCC
        } | ConvertTo-Json

        try {
            Invoke-RestMethod -Uri $WebhookUrl -Method Post -Body $Payload -ContentType "application/json"
            Write-Host "[#$($Ticket.id)] Send-CleanCoCosolEmail: Email sent successfully."
        }
        catch {
            Write-Error "[#$($Ticket.id)] Send-CleanCoCosolEmail: Failed to send JSON Payload. $_"
        }
    }
    else {
        Write-Host "[#$($Ticket.id)] Send-CleanCoCosolEmail: Ticket does not meet the criteria."
    }
}

function Add-CleanCoCosolNumber {
    param (
        [Parameter(Mandatory)]$Ticket
    )

    # Constants
    $TicketErrorThreshold = "Add-CleanCoCosolNumber-Tracker-5"
    
    # Retrieve relevant ticket notes and documents
    $TicketTracker = Get-CWMTicketNote -parentId $Ticket.id -condition "text like '%Add-CleanCoCosolNumber-Tracker-%'" -orderBy "id desc" | Select-Object -First 1
    $TicketError = Get-CWMTicketNote -parentId $Ticket.id -condition "text like '%Add-CleanCoCosolNumber-Pre-Error%'" -orderBy "id desc" | Select-Object -First 1
    $TicketSkip = Get-CWMTicketNote -parentId $Ticket.id -condition "text like '%This SAP/SuccessFactor ticket has not received the initial email from Cosol%'" -orderBy "id desc" | Select-Object -First 1
    $TicketDocumentDatas = Get-CWMDocument -recordType Ticket -recordId $Ticket.id -orderBy "id asc" -all

    # Find the relevant document
    foreach ($TicketDocumentData in $TicketDocumentDatas) {
        if ($TicketDocumentData.documentType.name -eq "EML" -and $TicketDocumentData.fileName -match ".*CLNCO-.*") {
            if (-not $TicketDocument -or $TicketDocumentData.id -lt $TicketDocument.id) {
                $TicketDocument = $TicketDocumentData
            }
        }
    }

    # Check if the ticket meets the criteria
    if (($Ticket.company.id -eq $CleanCoCompanyId) -and `
        ($($Ticket.board.name) -eq "CQL-HelpDesk (No SLA)") -and `
        ($Ticket.summary -notmatch "\[CLNCO-\d{1,6}\]") -and `
        ($Ticket.summary -match "(?i)\b(SAP|SuccessFactor|SuccessFactors)\b")) {
        
        Write-Warning "[#$($Ticket.id)] Add-CleanCoCosolNumber: #$($Ticket.id) - CleanCo QLD SAP/SuccessFactor ticket without a Cosol number in the ticket summary."

        if ($TicketDocument) {
            Write-Warning "[#$($Ticket.id)] Add-CleanCoCosolNumber: #$($Ticket.id) - Received the initial email from Cosol containing the CLNCO number."
            $CLNCOIndex = $TicketDocument.fileName.indexof("CLNCO-")
            $PostCLNCOIndex = $TicketDocument.fileName.indexof(" Service Ticket")
            if ($PostCLNCOIndex -eq -1) {
                $PostCLNCOIndex = $TicketDocument.fileName.indexof(" Ticket")
            }
            $CLNCO = ($TicketDocument.fileName.Substring($CLNCOIndex, ($PostCLNCOIndex - $CLNCOIndex)))
            $CLNCONumber = $CLNCO -replace "\D", ""
            Write-Host "[#$($Ticket.id)] Add-CleanCoCosolNumber: Extracted Cosol number - $($CLNCONumber)."
            if ($Ticket.summary.Length -ge 85) {
                $Summary = "[CLNCO-$($CLNCONumber)] $($Ticket.summary.Substring(0, 85))"
            }
            else {
                $Summary = "[CLNCO-$($CLNCONumber)] $($Ticket.summary)"
            }

            Update-TicketField -TicketId $Ticket.id -Path 'summary' -Value $Summary
            Write-Host "[#$($Ticket.id)] Add-CleanCoCosolNumber: #$($Ticket.id) New summary constructed and updated '$($Summary)'."
            $TicketNoteUpdate = @("CleanCo Cosol number [CLNCO-$($CLNCONumber)] has been added to the ticket summary.")
            New-CWMTicketNote -ticketId $Ticket.id -text $($TicketNoteUpdate) -internalAnalysisFlag $true
            Write-Host "[#$($Ticket.id)] Add-CleanCoCosolNumber: Note '$($TicketNoteUpdate)' added to Discussion."

            Update-TicketField -TicketId $Ticket.id -Path 'agreement' -ValueId $NoSLAAgreement
            Write-Host "[#$($Ticket.id)] Add-CleanCoCosolNumber: #$($Ticket.id) Agreement updated to '2-No SLA Managed Services Agreement'."
            $TicketNoteUpdate = @("Agreement updated to '2-No SLA Managed Services Agreement'.")
            New-CWMTicketNote -ticketId $Ticket.id -text $($TicketNoteUpdate) -internalAnalysisFlag $true
            Write-Host "[#$($Ticket.id)] Add-CleanCoCosolNumber: Note '$($TicketNoteUpdate)' added to Discussion."

        }
        elseif ($TicketSkip) {
            Write-Warning "[#$($Ticket.id)] Add-CleanCoCosolNumber: #$($Ticket.id) - SKIP - Existing error."
        }
        elseif ($TicketError) {
            Write-Warning "[#$($Ticket.id)] Add-CleanCoCosolNumber: #$($Ticket.id) - ERROR - Ticket has not received the initial email from Cosol."
            $TicketNoteUpdate = @("This SAP/SuccessFactor ticket has not received the initial email from Cosol. Please create a new internal ticket and assign to an Automation technician to review '[CQL] SAP - Submission'.")
            New-CWMTicketNote -ticketId $Ticket.id -text $($TicketNoteUpdate) -internalAnalysisFlag $true
            Write-Host "[#$($Ticket.id)] Add-CleanCoCosolNumber: Ticket Note '$($TicketNoteUpdate)' added to Discussion."
            Update-CWMTicket -id $Ticket.id -Operation replace -Path status/id -Value $StatusSchedulingRequired
            Write-Host "[#$($Ticket.id)] Add-CleanCoCosolNumber: Ticket Status updated to 'Scheduling Required'."
        }
        elseif ($TicketTracker) {
            $TicketTrackerMatches = [regex]::Matches($TicketTracker.text, '(\d+)$')
            if ($TicketTrackerMatches.Count -gt 0) {
                $TicketTrackerNumber = [int]$TicketTrackerMatches[0].Value + 1
            }
            else {
                $TicketTrackerNumber = 1
            }
            $TicketNoteUpdate = "Add-CleanCoCosolNumber-Tracker-$TicketTrackerNumber"
            if ($TicketNoteUpdate -match $TicketErrorThreshold) {
                Write-Warning "[#$($Ticket.id)] Add-CleanCoCosolNumber: #$($Ticket.id) - Final Check - Ticket has not received the initial email from Cosol."
                $TicketNoteUpdate = "Add-CleanCoCosolNumber-Pre-Error"
                New-CWMTicketNote -ticketId $Ticket.id -text $($TicketNoteUpdate) -internalAnalysisFlag $true
            }
            else {
                Write-Warning "[#$($Ticket.id)] Add-CleanCoCosolNumber: #$($Ticket.id) - Appending Check - Ticket has not received the initial email from Cosol."
                New-CWMTicketNote -ticketId $Ticket.id -text $($TicketNoteUpdate) -internalAnalysisFlag $true
            }
        }
        else {
            Write-Warning "[#$($Ticket.id)] Add-CleanCoCosolNumber: #$($Ticket.id) - Initial Check - Ticket has not received the initial email from Cosol."
            $TicketNoteUpdate = "Add-CleanCoCosolNumber-Tracker-1"
            New-CWMTicketNote -ticketId $Ticket.id -text $($TicketNoteUpdate) -internalAnalysisFlag $true
        }
    }
    else {
        Write-Host "[#$($Ticket.id)] Add-CleanCoCosolNumber: Ticket does not meet the criteria."
    }
}

function Add-CleanCoChangeNumber {
    param ([Parameter(Mandatory)]$Ticket)

    <#
    .DESCRIPTION
    Scrapes the ticket number of CleanCo RFCs, and adds it to the ticket's summary.
    #>
    
    # Check if ticket is for CQL, doesn't have a change number in the summary, and was last updated by email
    if (($Ticket.company.id -eq $CleanCoCompanyId) -and ($Ticket.summary -notmatch "CHN-\d+")`
            -and ($Ticket._info.updatedBy -eq "Email Connector")) {
        Write-Warning "[#$($Ticket.id)] Add-CleanCoChangeNumber: #$($Ticket.id) is a CleanCo QLD ticket updated by email with no change number."

        # Get most recent note
        $Note = Get-CWMTicketNote -ticketId $Ticket.id -orderBy "id desc" -pageSize 1

        # Check if the most recent note was added by Freshservice
        if ($Note.createdBy -match "helpdesk@cleancohelpdesk\.freshservice\.com") {
            Write-Warning "[#$($Ticket.id)] Add-CleanCoChangeNumber: #$($Ticket.id)'s last note ($($Note.id)) was added by Freshservice. Change number will be extracted."

            # Find change number in note
            $CHNIndex = $Note.text.indexof("CHN-")
            $PostCHNIndex = $Note.text.indexof("**Change Requester**")
            $CHN = ($Note.text.Substring($CHNIndex, ($PostCHNIndex - $CHNIndex)))
            $CHNNumber = $CHN -replace "\D", ""
            Write-Host "[#$($Ticket.id)] Add-CleanCoChangeNumber: Extracted change number - $($CHNNumber)."

            # Build new summary with change number
            if ($Ticket.summary.Length -ge 85) {
                $Summary = "$($Ticket.summary.Substring(0, 85)) [#CHN-$($CHNNumber)]"
            }
            else {
                $Summary = "$($Ticket.summary) [#CHN-$($CHNNumber)]"
            }
            Write-Host "[#$($Ticket.id)] Add-CleanCoChangeNumber: New summary constructed - $($Summary)."

            # Update summary
            Update-TicketField -TicketId $Ticket.id -Path 'summary' -Value $Summary
        }
        else {
            Write-Host "[#$($Ticket.id)] Add-CleanCoChangeNumber: #$($Ticket.id)'s last note ($($Note.id)) was added by '$($Note.createdBy)'."
        }
    }
    else {
        Write-Host "[#$($Ticket.id)] Add-CleanCoChangeNumber: #$($Ticket.id) does not meet the criteria."
    }
}

function Add-CleanCoCHGNumber {
    param ([Parameter(Mandatory)]$Ticket)
    
    # Check if ticket is for CQL, doesn't have a change number in the summary, and was last updated by email
    if (($Ticket.company.id -eq $CleanCoCompanyId) -and ($Ticket.summary -notmatch "CHG\d{7}")`
            -and ($Ticket._info.updatedBy -eq "Email Connector")) {
        Write-Warning "[#$($Ticket.id)] Add-CleanCoCHGNumber: #$($Ticket.id) is a CleanCo QLD ticket updated by email with no change number."

        # Get most recent note
        $Note = Get-CWMTicketNote -ticketId $Ticket.id -orderBy "id desc" -pageSize 1

        # Check if the most recent note was added by Service Now
        if ($Note.createdBy -match "servicenow@cleancoqld\.com\.au") {
            Write-Warning "[#$($Ticket.id)] Add-CleanCoCHGNumber: #$($Ticket.id)'s last note ($($Note.id)) was added by ServiceNow. Change number will be extracted."
            
            # Use regex to extract the CHG number
            $CHGPattern = "CHG(\d{7})"
            if ($Note.text -match $CHGPattern) {
                $CHGNumber = $Matches[1]
                Write-Warning "[#$($Ticket.id)] Add-CleanCoCHGNumber: Extracted change number - $($CHGNumber)."
            }
            else {
                Write-Warning "[#$($Ticket.id)] Add-CleanCoCHGNumber: No change number found in the note text."
            }

            # Build new summary with change number
            if ($Ticket.summary.Length -ge 85) {
                $Summary = "$($Ticket.summary.Substring(0, 85)) [CHG$($CHGNumber)]"
            }
            else {
                $Summary = "$($Ticket.summary) [CHG$($CHGNumber)]"
            }
            Write-Host "[#$($Ticket.id)] Add-CleanCoCHGNumber: New summary constructed - $($Summary)."

            # Update summary
            Update-TicketField -TicketId $Ticket.id -Path 'summary' -Value $Summary
        }
        else {
            Write-Host "[#$($Ticket.id)] Add-CleanCoCHGNumber: #$($Ticket.id)'s last note ($($Note.id)) was added by '$($Note.createdBy)'."
        }
    }
    else {
        Write-Host "[#$($Ticket.id)] Add-CleanCoCHGNumber: #$($Ticket.id) does not meet the criteria."
    }
}

function Update-PiaTicketDefaultContactFlag {
    param ([Parameter(Mandatory)]$Ticket)

    <#
    .DESCRIPTION
    Enables the contact flag for all tickets on the Pia board.

    .NOTES
    This is required for emails to be sent from Pia Canvas, as their setup has no way to update this flag on our behalf.
    #>

    # Check if ticket is on the Pia board and has an incorrectly set contact flag
    if ($Ticket.board.id -eq $PiaBoardId -and !$Ticket.automaticEmailContactFlag) {
        Write-Warning "[#$($Ticket.id)] Update-PiaTicketDefaultContactFlag: #$($Ticket.id) is on the Pia board, but does not have the automatic email contact flag enabled."

        # Update flag
        Update-TicketField -TicketId $Ticket.id -Path 'automaticEmailContactFlag' -Value $true
    }
    else {
        Write-Host "[#$($Ticket.id)] Update-PiaTicketDefaultContactFlag: #$($Ticket.id) does not meet the criteria."
    }
}

function Update-LMTicketCI {
    param (
        [Parameter(Mandatory)] $Ticket
    )

    $ConfigCheck = Get-CWMTicketConfiguration -parentId $Ticket.id
    $TicketSummaryPattern = '^(?<Alert>alert (error|warn|critical)|eventalert (error|warn|critical))|(?<Collector>Collector Down Alert)'
    $TicketNoteUpdates = @()

    if (($null -eq $ConfigCheck) -and
        (($Ticket.board.name -eq "ALERTS - Events") -or ($Ticket.board.name -eq "ALERTS - Incidents")) -and
        ($Ticket.summary -match $TicketSummaryPattern)) {

        Write-Warning "[#$($Ticket.id)] Update-LMTicketCI: LM Alert ticket without a CI attached."

        $IsAlert = $matches['Alert']
        $IsCollector = $matches['Collector']

        if ($IsAlert) {
            Write-Warning "[#$($Ticket.id)] Update-LMTicketCI: Matched Alert - $($matches[0])"

            $TicketAlertPattern = "^alert (error|warn|critical) |eventalert (error|warn|critical) "
            $TicketAlertCleanedLine = $Ticket.summary -replace $TicketAlertPattern, ""
            $CIName = $TicketAlertCleanedLine.Split(" ")[0]
            if ($CIName -like "*:*") {
                Write-Warning "[#$($Ticket.id)] Update-LMTicketCI: Azure Resource"
                $TicketNoteUpdates += "This alert is for an Azure Resource"
                $CIName = $CIName.Split(":", [System.StringSplitOptions]::RemoveEmptyEntries)[-1]
            }            

            $Condition = "name='$CIName' AND status/name='Active'"
            $CIs = Get-CWMCompanyConfiguration -condition $Condition
            if ($CIs.Count -ge 1) {
                foreach ($CI in $CIs) {
                    Write-Warning "[#$($Ticket.id)] Update-LMTicketCI: Matched CI [$($CI.id)] - $($CI.name)"
                    New-CWMTicketConfiguration -TicketID $Ticket.id -deviceIdentifier $CI.id
                    $TicketNoteUpdates += "Attached CI [$($CI.id)] - $($CI.name) to the ticket."
                }
            }
            else {
                $TicketNoteUpdates += "There is no CI for $CIName. Please consider if this object requires a CI. If unsure please consult with the Senior Technician Team."
            }

            $CILive = $CIs | Where-Object { $_.type.name -notlike "Managed*" }
            if ($CILive.Count -gt 1) {
                $TicketNoteUpdates += "There are duplicate 'Non-Managed' tagged CIs for $CIName. Please amend this manually."
            }
        }

        elseif ($IsCollector) {
            Write-Warning "[#$($Ticket.id)] Update-LMTicketCI: Matched Collector Down Alert - $($matches[0])"

            $TicketInitial = Get-CWMTicketNote -parentId $Ticket.id -orderBy "id asc" | Select-Object -First 1
            $TicketInitialNote = $TicketInitial.text
            $TicketInitialFirstLine = $TicketInitialNote.Split("`n")[0]
            $TicketInitialPattern = "^(?:LMA\d+\sLogicMonitor\sCollector\s\d+\s-\s(?:[A-Z0-9\\]+\\)?)?(.*)"
            $TicketInitialCleanedLine = $TicketInitialFirstLine -replace $TicketInitialPattern, ""
            $CIName = $TicketInitialCleanedLine.Split(" ")[0]

            $Condition = "name='$CIName' AND status/name='Active'"
            $CIs = Get-CWMCompanyConfiguration -condition $Condition
            if ($CIs.Count -ge 1) {
                foreach ($CI in $CIs) {
                    Write-Warning "[#$($Ticket.id)] Update-LMTicketCI: Matched CI [$($CI.id)] - $($CI.name)"
                    New-CWMTicketConfiguration -TicketID $Ticket.id -deviceIdentifier $CI.id
                    $TicketNoteUpdates += "Attached CI [$($CI.id)] - $($CI.name) to the ticket."
                }
            }
            else {
                $TicketNoteUpdates += "There is no CI for $CIName. Please consider if this object requires a CI. If unsure please consult with the Senior Technician Team."
            }

            $CILive = $CIs | Where-Object { $_.type.name -notlike "Managed*" }
            if ($CILive.Count -gt 1) {
                $TicketNoteUpdates += "There are duplicate 'Non-Managed' tagged CIs for $CIName. Please amend this manually."
            }
        }

        foreach ($TicketNoteUpdate in $TicketNoteUpdates) {
            New-CWMTicketNote -ticketId $Ticket.id -text $TicketNoteUpdate -internalAnalysisFlag $true
            Write-Warning "[#$($Ticket.id)] Update-LMTicketCI: $TicketNoteUpdate"
        }
    }    
    else {
        Write-Warning "[#$($Ticket.id)] Update-LMTicketCI: Ticket does not meet the criteria."
    }
}

function Update-TicketCI {
    param (
        [Parameter(Mandatory)] $Ticket
    )

    $ExcludedCIs = @(
        "USB-C",
        "Automation - Persona",
        "Automation - Submission"
    )

    if ($ticket) {
        Write-Warning "[#$($Ticket.id)] Update-TicketCI: Ticket Match."
        $TicketNoteEntry = [DateTime]::Parse((Get-CWMTicketNote -parentId $Ticket.id -orderBy "id desc" | Select-Object -First 1)._info.lastUpdated)
        $LatestTimeEntry = Get-CWMTimeEntry -condition "ticket/id=$($Ticket.id)" -orderBy "id desc" | Select-Object -First 1
        if ($LatestTimeEntry -and $LatestTimeEntry._info.lastUpdated) {
            $TicketTimeEntry = [DateTime]::Parse($LatestTimeEntry._info.lastUpdated)
            Write-Output "Parsed DateTime: $TicketTimeEntry"
        } else {
            Write-Warning "No valid time entry found for ticket ID $($Ticket.id)."
        }

        if ($TicketNoteEntry -gt $TicketTimeEntry) {
            Write-Warning "TicketNoteEntry is the later one: $TicketNoteEntry"
            $TicketLatestNote = (Get-CWMTicketNote -parentId $Ticket.id -orderBy "id desc" | Select-Object -First 1).text
        }
        elseif ($TicketTimeEntry -gt $TicketNoteEntry) {
            Write-Warning "TicketTimeEntry is the later one: $TicketTimeEntry"
            $TicketLatestNote = (Get-CWMTimeEntry -condition "ticket/id=$($Ticket.id)" -orderBy "id desc" | Select-Object -First 1).notes
        }
        else {
            Write-Warning "ERROR: Both timestamps are the same: $TicketNoteEntry"
        }

        $CIpattern = '\b([A-Za-z][A-Za-z0-9]{0,5}(?:-[A-Za-z0-9]{1,30}){1,4})(?:[\\_][A-Za-z0-9_]+)?'
        $CIArray = [regex]::Matches($TicketLatestNote, $CIpattern) |
            ForEach-Object { $_.Value.Replace('\', '') } |
            Where-Object {
                $value = $_.ToLower()
                -not ($ExcludedCIs | Where-Object { $value -match [regex]::Escape($_.ToLower()) })
            } |
            Sort-Object -Unique

        $CIChecks = Get-CWMTicketConfiguration -parentId $Ticket.id

        if ($null -ne $CIArray) {
            Write-Warning "[#$($Ticket.id)] Update-TicketCI: CI Matches found."

            foreach ($CINames in $CIArray) {
                $CheckCondition = "name like '%$CINames%' AND status/name='Active'"
                $CheckCIs = Get-CWMCompanyConfiguration -condition $CheckCondition

                if ($null -eq $CheckCIs) {
                    Write-Warning "[#$($Ticket.id)] Update-TicketCI: No matching CI found for $CINames."
                    continue
                }

                foreach ($CheckCI in $CheckCIs) {
                    $IsAlreadyAttached = $CIChecks | Where-Object { $_.id -eq $CheckCI.id }
                    if ($IsAlreadyAttached) {
                        Write-Warning "[#$($Ticket.id)] Update-TicketCI: CI $CINames already attached."
                        continue
                    }

                    Write-Warning "[#$($Ticket.id)] Update-TicketCI: Adding CI [$($CheckCI.id)] - $($CheckCI.name)"
                    New-CWMTicketConfiguration -TicketID $Ticket.id -deviceIdentifier $CheckCI.id
                }
            }
        }
        else {
            Write-Warning "[#$($Ticket.id)] Update-TicketCI: No matches found."
        }
    }
    else {
        Write-Warning "[#$($Ticket.id)] Update-LMTicketCI: Ticket does not meet the criteria."
    }
}