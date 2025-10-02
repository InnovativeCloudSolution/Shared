using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Robot ticket updaters that we want to ignore
$Bots = @(
    "AzFuncBundler"
    "mits-cwm-azure"
)

# Complete validation from secret key in query
if ($Request.Query.secret -eq $env:SECRETKEY) {
    # Grab ticket details from request body
    $Ticket = $Request.Body.Entity | ConvertFrom-Json

    # Only run rest of function if the ticket isn't null
    if ($null -ne $Ticket) {
        # Write to the Azure Functions log stream.
        Write-Host "[#$($Ticket.id)] TICKET $(($Request.Body.Action).ToUpper()) | SUMMARY: $($Ticket.summary) | UPDATED BY: $($Ticket._info.updatedBy)"

        # Check if the ticket isn't a service ticket and should therefore be ignored
        if ($Ticket.recordType -ne "ServiceTicket") {
            Write-Output "[#$($Ticket.id)] Ticket is not a service desk ticket ($($Ticket.recordType)). Request ignored."
            Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::Forbidden
                Body = "Ticket is not a service desk ticket ($($Ticket.recordType)). Request ignored."
            })
        }
        
        # Check if ticket was updated by an automated process and should therefore be ignored
        elseif ($Bots -contains $Ticket._info.updatedBy) {
            Write-Output "[#$($Ticket.id)] Updated by bot ($($Ticket._info.updatedBy)). Request ignored."
            Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::Forbidden
                Body = "Ticket updated by bot ($($Ticket._info.updatedBy)). Request ignored."
            })
        }
        
        # Run functions after checks and balances complete
        else {
            # Connect to CWM
            Connect-CWManage -TicketId $Ticket.id
                        
            if ($Request.Body.Action -eq 'added') {
                # Add standard notes if ticket is new
                Write-Host "[#$($Ticket.id)] Checking for standard note validity..."
                Add-TriageQltyNote -Ticket $Ticket
                Add-CrowdStrikeCI -Ticket $Ticket
                Add-NoteCQLMerakiUpdates -Ticket $Ticket
                Add-NoteOPCExtendedHoursSupport -Ticket $Ticket
                Add-NoteNonSupportUser -Ticket $Ticket
                Add-NoteCQLP1P2Incidents -Ticket $Ticket
                Add-NoteCQLOverseasAccessRequest -Ticket $Ticket
                Add-NoteThreatLockerRequest -Ticket $Ticket
                Add-NoteOPCUnitedStatesUser -Ticket $Ticket

                # Check to see if the ticket can be bundled
                Write-Host "[#$($Ticket.id)] Checking to see if it can be bundled anywhere..."
                Update-ParentPelotonTickets -Ticket $Ticket
                Update-ParentThirdPartyCQLTickets -Ticket $Ticket
                Update-ParentMerakiTickets -Ticket $Ticket

                # Check for other new ticket functions
                Update-ExitUserTicketSummary -Ticket $Ticket
                Update-UserOffboardingTicketSummary -Ticket $Ticket
                Update-UserOffboardingTicketSummaryPre -Ticket $Ticket
                Update-LMTicketCI -Ticket $Ticket
            }

            # Specific functions for updated tickets
            elseif ($Request.Body.Action -eq 'updated') {
                # Check if ticket is a CleanCo SAP/SuccessFactor with a response from CQL/COSOL
                Send-CleanCoCosolEmail -Ticket $Ticket
                # Check if ticket is a CleanCo SAP/SuccessFactor ticket
                Add-CleanCoCosolNumber -Ticket $Ticket
                # Check if ticket is a CleanCo ServiceNow RFC with change number
                Add-CleanCoCHGNumber -Ticket $Ticket
                # Check if ticket is a CleanCo FreshService RFC with change number
                Add-CleanCoChangeNumber -Ticket $Ticket
                # Check if ticket is from Argent/Elston and scrape CI from latest ticket note and attempts to attach CI
                Update-TicketCI -Ticket $Ticket
            }

            # Check if ticket is a child ticket, and if the status is wrong
            Update-ChildTicket -Ticket $Ticket

            # Checks to make sure the ticket's owner matches its resource
            Update-TicketOwner -Ticket $Ticket

            # Checks to see if the ticket is on the Pia board, and if the contact flag is set correctly
            Update-PiaTicketDefaultContactFlag -Ticket $Ticket

            # Disconnects from CW Manage
            Disconnect-CWM

            # Associate values to output bindings by calling 'Push-OutputBinding'.
            Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::OK
                Body = "Success. Please check the Azure Functions app 'MIT-TESTING-DEV-CRM' for logs."
            })
        }
    } else {
        Write-Output "Request received, but ticket object unable to be converted. Ticket object: $($Request.Body.Entity)"
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body = "Request received, but ticket object unable to be converted."
        })
    }
} else {
    Write-Output "Request not authorised."
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::Unauthorized
        Body = "Request not authorised."
    })
}