<#

Mangano IT - New User Flow - Create Tasks for New User Flow
Created by: Gabriel Nugent
Version: 1.20.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory=$true)]$TicketId,
    [string]$CustomTasks,
    [string]$UsageLocation,
    [string]$Licenses,
	[string]$SecurityGroups,
    [string]$365Groups,
    [string]$SharedMailboxes,
    [string]$DistributionGroups,
    [string]$MailEnabledSecurityGroups,
    [string]$TeamsChannels,
	[string]$PublicFolders,
    [string]$CalendarFolders,
    [bool]$GroupBasedLicensing = $true,
    [string]$SMSContacts,
    [string]$ComputerRequiredTicket,
    [string]$MobileRequiredTicket,
    [string]$EquipmentRequiredTicket,
    [string]$DeskphoneRequiredTicket,
    [bool]$TeamsCalling,
    [bool]$CQLHelpDeskBoard = $false,
    [bool]$Bullphish = $false
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

# Initialize class for ticket task
class Task {
    [string]$Notes
    [string]$Resolution
    [boolean]$Exists
    [boolean]$Closed
}

# Define standard tasks
$Tasks = @(
    [Task]@{Notes='Review the ticket';Resolution='I have read the ticket and all actions have been taken as requested.';Exists=$false;Closed=$false},
    [Task]@{Notes='Review all additional details';Resolution='I have ensured any additional notes/requested items have been reviewed and action taken as required.';Exists=$false;Closed=$false},
    [Task]@{Notes='Have the ticket peer reviewed';Resolution='I have had this ticket peer reviewed to ensure all steps have been met.';Exists=$false;Closed=$false},
    [Task]@{Notes='Reach out to user';Resolution='I have spoken to the new user (or their manager) and confirmed that they can sign in.';Exists=$false;Closed=$false},
    [Task]@{Notes='Create New User';Resolution='To be filled with user details upon completion.';Exists=$false;Closed=$false},
    [Task]@{Notes='Create ConnectWise Contact';Exists=$false;Closed=$false}    
)

# Define tasks where multiple can have the same name
$MultiTasks = @()
$LicenseTasks = @()

# Add custom tasks
foreach ($CustomTask in $CustomTasks.split(";")) {
    if ($CustomTask -ne '') {
        $TaskDetails = $CustomTask.split(":")
        $Notes = $TaskDetails[0]
        $Resolution = $TaskDetails[1]
        $Tasks += [Task]@{Notes=$Notes;Resolution=$Resolution;Exists=$false;Closed=$false}
    }
}

# Add task for usage location if needed
if ($UsageLocation -ne '') {
    $Tasks += [Task]@{Notes='Set Usage Location';Resolution=$UsageLocation;Exists=$false;Closed=$false}
}

# Add task for Teams Calling if needed
if ($TeamsCalling) {
    $Tasks += [Task]@{Notes='Setup Teams Calling';Resolution='';Exists=$false;Closed=$false}
}

# Add task for HelpDesk board depending on company
if ($CQLHelpDeskBoard) {
    $Tasks += [Task]@{Notes='Move to CQL-HelpDesk (MS)';Resolution='Once this task is complete, automation should be finished.';Exists=$false;Closed=$false}
} else {
    $Tasks += [Task]@{Notes='Move to HelpDesk (MS)';Resolution='Once this task is complete, automation should be finished.';Exists=$false;Closed=$false}
}

# Add tasks for licenses
# Platform can be Pax8 or Telstra
# Id represents the ID for that platform
# SkuPartNumber is a specific Microsoft field, generic across clients
foreach ($License in $Licenses.split(";")) {
    if ($License -ne '') {
        $TaskDetails = $License.split(":")
        $Resolution = @{
            Name = $TaskDetails[0]
            Platform = $TaskDetails[1]
            PlatformId = $TaskDetails[2]
            SkuPartNumber = $TaskDetails[3]
            BillingTerm = $TaskDetails[4]
        } | ConvertTo-Json
        $LicenseTasks += [Task]@{Notes='Purchase License';Resolution=$Resolution;Exists=$false;Closed=$false}
        if (!$GroupBasedLicensing) {
            $LicenseTasks += [Task]@{Notes='Assign License';Resolution=$Resolution;Exists=$false;Closed=$false}
        }
        $LicenseTasks += [Task]@{Notes='Confirm License is Assigned';Resolution=$Resolution;Exists=$false;Closed=$false}
    }
}

# Add tasks for security groups
foreach ($SecurityGroup in $SecurityGroups.split(";")) {
    if ($SecurityGroup -ne '') {
        $Resolution = @{
            Name = $SecurityGroup
        } | ConvertTo-Json
        $MultiTasks += [Task]@{Notes='Add to On-Premises Security Group';Resolution=$Resolution;Exists=$false;Closed=$false}
    }
}

# Add tasks for M365 groups
foreach ($365Group in $365Groups.split(";")) {
    if ($365Group -ne '') {
        $TaskDetails = $365Group.split(":")
        if ($TaskDetails[1] -eq 'true') { $Owner = $true } else { $Owner = $false } 
        $Resolution = @{
            Name = $TaskDetails[0]
            Owner = $Owner
        } | ConvertTo-Json
        $MultiTasks += [Task]@{Notes='Add to Microsoft 365/AAD Security Group';Resolution=$Resolution;Exists=$false;Closed=$false}
    }
}

# Add tasks for shared mailboxes
foreach ($SharedMailbox in $SharedMailboxes.split(";")) {
    if ($SharedMailbox -ne '') {
        $TaskDetails = $SharedMailbox.split(":")
        $Resolution = @{
            Name = $TaskDetails[0]
            AccessRights = $TaskDetails[1]
        } | ConvertTo-Json
        $MultiTasks += [Task]@{Notes='Grant Access to Shared Mailbox';Resolution=$Resolution;Exists=$false;Closed=$false}
    }
}

# Add tasks for calendar folders
foreach ($CalendarFolder in $CalendarFolders.split(";")) {
    if ($CalendarFolder -ne '') {
        $TaskDetails = $CalendarFolder.split(":")
        $Resolution = @{
            Name = $TaskDetails[0]
            AccessRights = $TaskDetails[1]
            SharingPermissionFlags = $TaskDetails[2]
        } | ConvertTo-Json
        $MultiTasks += [Task]@{Notes='Grant Access to Calendar';Resolution=$Resolution;Exists=$false;Closed=$false}
    }
}

# Add tasks for distribution groups
foreach ($DistributionGroup in $DistributionGroups.split(";")) {
    if ($DistributionGroup -ne '') {
        $Resolution = @{
            Name = $DistributionGroup
        } | ConvertTo-Json
        $MultiTasks += [Task]@{Notes='Add to Distribution Group';Resolution=$Resolution;Exists=$false;Closed=$false}
    }
}

# Add tasks for mail-enabled security groups
foreach ($MailEnabledSecurityGroup in $MailEnabledSecurityGroups.split(";")) {
    if ($MailEnabledSecurityGroup -ne '') {
        $Resolution = @{
            Name = $MailEnabledSecurityGroup
        } | ConvertTo-Json
        $MultiTasks += [Task]@{Notes='Add to Mail-Enabled Security Group';Resolution=$Resolution;Exists=$false;Closed=$false}
    }
}

# Add tasks for Teams channels
foreach ($TeamsChannel in $TeamsChannels.split(";")) {
    if ($TeamsChannel -ne '') {
        $Resolution = @{
            Name = $TeamsChannel
        } | ConvertTo-Json
        $MultiTasks += [Task]@{Notes='Add to Teams Channel';Resolution=$Resolution;Exists=$false;Closed=$false}
    }
}

# Add tasks for public folders
foreach ($PublicFolder in $PublicFolders.split(";")) {
    if ($PublicFolder -ne '') {
        $TaskDetails = $PublicFolder.split(":")
        $Resolution = @{
            Name = $TaskDetails[0]
            AccessRights = $TaskDetails[1]
        } | ConvertTo-Json
        $MultiTasks += [Task]@{Notes='Grant Access to Public Folder';Resolution=$Resolution;Exists=$false;Closed=$false}
    }
}

# Add tasks for SMS recipients
foreach ($SMSContact in $SMSContacts.split(";")) {
    if ($SMSContact -ne '') {
        $TaskDetails = $SMSContact.split(":")
        $MobileNumber = .\SMS-FormatValidPhoneNumber.ps1 -PhoneNumber $TaskDetails[0] -IncludePlus $false -KeepSpaces $false
        $Resolution = @{
            MobileNumber = $MobileNumber
            ContactType = $TaskDetails[1]
            CountryCode = $TaskDetails[2]
        } | ConvertTo-Json
        $MultiTasks += [Task]@{Notes='Send SMS';Resolution=$Resolution;Exists=$false;Closed=$false}
    }
}

# Add task for computer ticket
if ($ComputerRequiredTicket -eq 'Yes') {
    $Tasks += [Task]@{Notes='Make Ticket for Computer';Exists=$false;Closed=$false}
}

# Add task for mobile device ticket
if ($MobileRequiredTicket -eq 'Yes') {
    $Tasks += [Task]@{Notes='Make Ticket for Mobile Device';Exists=$false;Closed=$false}
}

# Add task for equipment ticket
if ($EquipmentRequiredTicket -eq 'Yes') {
    $Tasks += [Task]@{Notes='Make Ticket for Additional Equipment';Exists=$false;Closed=$false}
}

# Add task for deskphone ticket
if ($DeskphoneRequiredTicket -eq 'Yes') {
    $Tasks += [Task]@{Notes='Make Ticket for Deskphone';Exists=$false;Closed=$false}
}

# Add task for Bullphish senders
if ($Bullphish) {
    $Tasks += [Task]@{Notes='Add Bullphish Addresses to Safe Senders List';Exists=$false;Closed=$false}
}

# Set CRM variables, connect to server
$Server = Get-AutomationVariable -Name 'CWManageUrl'
$Company = Get-AutomationVariable -Name 'CWManageCompanyId'
$PublicKey = Get-AutomationVariable -Name 'PublicKey'
$PrivateKey = Get-AutomationVariable -Name 'PrivateKey'
$ClientId = Get-AutomationVariable -Name 'clientId'

# Create an object with CW credentials 
$Connection = @{
	Server = $Server
	Company = $Company
	pubkey = $PublicKey
	privatekey = $PrivateKey
	clientId = $ClientId
}

# Connect to our CWM server
Connect-CWM @Connection

# Get all currently existing tasks on this ticket
try {
    $Log += "Fetching ticket tasks for #$TicketId..."
    $TicketTasks = Get-CWMTicketTask -parentId $TicketId -all
    $Log += "`nSUCCESS: Ticket tasks fetched for #$TicketId.`n"
} catch {
    $Log += "`nERROR: Ticket tasks unable to be fetched for #$TicketId.`nERROR DETAILS: " + $_ + "`n"
    Write-Error "Ticket tasks unable to be fetched for #$TicketId : $_"
}

## FETCH TICKET TASKS ##

# For each task on the ticket, if it already exists, check if it's already marked as closed
:Outer foreach ($TicketTask in $TicketTasks) {
    foreach ($Task in $Tasks) {
        if ($Task.Notes -like $TicketTask.notes) {
            $TaskNotes = $Task.Notes
            $TaskClosed = ($Task.Closed).ToString()
            $Task.Exists = $true
            $Task.Closed = $TicketTask.closedFlag
            $Log += "`nINFO: Task '$TaskNotes' already exists. Closed status: $TaskClosed"
            Write-Warning "INFO: Task '$TaskNotes' already exists. Closed status: $TaskClosed"
            continue Outer
        }
    }

    foreach ($Task in $MultiTasks) {
        if ($Task.Notes -like $TicketTask.notes -and $Task.Resolution -like $TicketTask.resolution) {
            $TaskNotes = $Task.Notes
            $TaskResolution = $Task.Resolution
            $TaskClosed = ($Task.Closed).ToString()
            $Task.Exists = $true
            $Task.Closed = $TicketTask.closedFlag
            $Log += "`nINFO: Task $TaskNotes (multi-task) already exists. Closed status: $TaskClosed"
            Write-Warning "INFO: Task $TaskNotes (multi-task) already exists. Closed status: $TaskClosed"
            continue Outer
        }
    }

    foreach ($Task in $LicenseTasks) {
        $ResolutionDetails = $Task.Resolution | ConvertFrom-Json
        $LicenseName = $ResolutionDetails.Name
        if ($Task.Notes -like $TicketTask.notes -and $TicketTask.resolution -like "*$LicenseName*") {
            $TaskNotes = $Task.Notes
            $TaskResolution = $Task.Resolution
            $TaskClosed = ($Task.Closed).ToString()
            $Task.Exists = $true
            $Task.Closed = $TicketTask.closedFlag
            $Log += "`nINFO: Task $TaskNotes (license task) already exists. Closed status: $TaskClosed"
            Write-Warning "INFO: Task $TaskNotes (license task) already exists. Closed status: $TaskClosed"
            continue Outer
        }
    }
}

# Add multitasks to tasks
$Tasks += $MultiTasks
$Tasks += $LicenseTasks

# Add extra line break
$Log += "`n"

## ADD TICKET TASKS ##

foreach ($Task in $Tasks) {
    if (!$Task.Exists) {
        $TaskNotes = $Task.Notes
        $TaskResolution = $Task.Resolution
        try {
            $Log += "`nCreating task '$TaskNotes - $TaskResolution'..."
            New-CWMTicketTask -parentId $TicketId -notes $Task.Notes -resolution $Task.Resolution -closedFlag $false | Out-Null
        }
        catch {
            $Log += "`nERROR: Ticket task '$TaskNotes - $TaskResolution' not created.`nERROR DETAILS: " + $_
            Write-Error "Ticket task '$TaskNotes - $TaskResolution' not created : $_"
        }
    }
}

## ADD NOTE TO TICKET TO EXPLAIN TASKS ##

$Text = "Create Tasks [automated action]:`n`n"
$Text += "As part of the updated new user flows, tasks are created for every action that the automation intends to take. "
$Text += "These tasks represent the steps of a new user request that needs to be completed (account creation, license purchasing, etc.). "
$Text += "Once the automation has completed a task, it will close it automatically. In some cases, it will add details to the task's resolution.`n`n"
$Text += "In the event that the automation fails (Internal Systems Review Required), the flow will need to be reviewed by an Internal Systems technician. "
$Text += "If one is not available to re-run the flow after fixing it, the tasks are your guide to what is yet to be completed.`n`n"
$Text += "See the company's new user doco for more information."

$NoteArguments = @{
    TicketId = $TicketId
    Text = $Text
    InternalFlag = $true
    ApiSecrets = $ApiSecrets
}
.\CWM-AddTicketNote.ps1 @NoteArguments | Out-Null

## SEND DETAILS BACK TO FLOW ##

# Get all currently existing tasks on this ticket again
try {
    $Log += "`nFetching ticket tasks again for #$TicketId..."
    $TicketTasks = Get-CWMTicketTask -parentId $TicketId -all
    $Log += "`nSUCCESS: Ticket tasks fetched for #$TicketId."
} catch {
    $Log += "`nERROR: Ticket tasks unable to be fetched for #$TicketId.`nERROR DETAILS: " + $_
    Write-Error "Ticket tasks unable to be fetched for #$TicketId : $_"
}

$Output = @{
    Log = $Log
    TicketTasks = $TicketTasks
}

# Send back to Power Automate
Write-Output $Output | ConvertTo-Json