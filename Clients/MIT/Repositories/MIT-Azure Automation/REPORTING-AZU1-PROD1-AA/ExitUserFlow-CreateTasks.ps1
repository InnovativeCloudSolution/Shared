<#

Mangano IT - Exit User Flow - Create Tasks for Exit User Flow
Created by: Gabriel Nugent
Version: 1.7.5

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory=$true)][int]$TicketId,
    $ApiSecrets = $null,
    [string]$DisabledUsersOU,
    [bool]$RemoveManager = $true,
    [bool]$HideFromGal = $true,
    [string]$SharedMailboxDelegate,
    [bool]$SetUpAutoResponse,
    [string]$AutoResponseTarget,
    [string]$AadUserId,
    [string]$AadUserPrincipalName,
    [string]$EmailAddress,
    [string]$AdminPrefix,
    [string]$CWCompanyName,
    [bool]$OnPremisesUser,
    [bool]$TeamsCallingClient,
    [bool]$CleanCoTasks = $false,
    [bool]$ElstonTasks = $false
)

## GET CREDENTIALS IF NOT PROVIDED ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

# Provide value if not given for disabled users OU
if ($DisabledUsersOU -eq '') { $DisabledUsersOU = 'N/A' }

# Initialize class for ticket task
class Task {
    [string]$Notes
    [string]$Resolution
    [boolean]$Exists
    [boolean]$Closed
}

## TASKS ##

# Create body for Provide Asset List task
$AssetListResolution = @{
    EmailAddress = $EmailAddress
    CompanyName = $CWCompanyName
} | ConvertTo-Json

# Create body for Disable User Account task
if ($OnPremisesUser) {
    $DisableUserResolution = @{
        DisabledUsersOU = $DisabledUsersOU
        RemoveManager = $RemoveManager
        HideFromGlobalAddressList = $HideFromGal
    } | ConvertTo-Json
} else { $DisableUserResolution = '' }

# Define standard tasks
$Tasks = @(
    [Task]@{Notes='Provide Asset List (manual)';Resolution=$AssetListResolution;Exists=$false;Closed=$false}
)

# Add task for privileged user workflow (CleanCo only)
if ($CleanCoTasks) {
    $Tasks += [Task]@{Notes='Complete Privileged User Exit Workflow (manual)';Resolution="This task will be closed if the user is not privileged.`n`nhttps://mits.au.itglue.com/3210579/docs/3446869";Exists=$false;Closed=$false}
}

$Tasks += [Task]@{Notes='Disable User Account';Resolution=$DisableUserResolution;Exists=$false;Closed=$false}

# Add task to remove admin accounts
$AdminResolution = @{
    UserPrincipalName = $EmailAddress
    AdminPrefix = $AdminPrefix
} | ConvertTo-Json
$Tasks += [Task]@{Notes='Check For and Disable Admin Account/s';Resolution=$AdminResolution;Exists=$false;Closed=$false}

if ($AadUserId -ne '') {
    $UserIdResolution = @{
        UserId = $AadUserId
        UserPrincipalName = $AadUserPrincipalName
    } | ConvertTo-Json
    $Tasks += [Task]@{Notes='Revoke Sign In Sessions in AAD';Resolution=$UserIdResolution;Exists=$false;Closed=$false}
    $Tasks += [Task]@{Notes='Disable MFA in AAD (manual)';Resolution='This step must be completed manually.';Exists=$false;Closed=$false}
}

# Add task for shared mailbox delegate
if ($SharedMailboxDelegate -ne '') {
    $SharedMailboxResolution = @{
        Delegate = $SharedMailboxDelegate
    } | ConvertTo-Json
} else {
    $SharedMailboxResolution = @{
        Delegate = 'null'
    } | ConvertTo-Json
}

# More standard tasks - creating in order
if ($AadUserId -ne '') {
    $Tasks += [Task]@{Notes='Remove from Cloud Groups';Resolution=$UserIdResolution;Exists=$false;Closed=$false}
}
$Tasks += [Task]@{Notes='Remove from Exchange Services';Resolution='';Exists=$false;Closed=$false}
$Tasks += [Task]@{Notes='Convert To Shared Mailbox';Resolution=$SharedMailboxResolution;Exists=$false;Closed=$false}

if ($OnPremisesUser) {
    $Tasks += [Task]@{Notes='Remove from On-Premises Security Groups';Resolution='';Exists=$false;Closed=$false}
}

if ($AadUserId -ne '') {
    $Tasks += [Task]@{Notes='Remove Azure AD Role Assignments (manual)';Resolution="This step cannot be completed by automation due to permission issues. Please complete this manually.";Exists=$false;Closed=$false}
    $Tasks += [Task]@{Notes='Remove Azure Enterprise Applications';Resolution=$UserIdResolution;Exists=$false;Closed=$false}
    $Tasks += [Task]@{Notes='Unassign Licenses';Resolution=$UserIdResolution;Exists=$false;Closed=$false}
}
$Tasks += [Task]@{Notes='Make Contact Inactive in ConnectWise Manage';Resolution='';Exists=$false;Closed=$false}

if ($SetUpAutoResponse) {
    if ($AutoResponseTarget -eq '') { $AutoResponseTarget = 'N/A' }
    $AutoResponseResolution = @{
        Target = $AutoResponseTarget
    } | ConvertTo-Json
    $Tasks += [Task]@{Notes='Set Up Auto-Response Email';Resolution=$AutoResponseResolution;Exists=$false;Closed=$false}
}

# Add Teams Calling task
<#
if ($TeamsCallingClient -and $AadUserId -ne '') {
    $TeamsCallingResolution = @{
        UserId = $AadUserId
        UserPrincipalName = $AadUserPrincipalName
    } | ConvertTo-Json
    $Tasks += [Task]@{Notes='Disable Teams Calling for User';Resolution=$TeamsCallingResolution;Exists=$false;Closed=$false}
}
#>

# Add CleanCo-specific tasks
if ($CleanCoTasks) {
    $Tasks += [Task]@{Notes='Remove StreamlineNX Card (manual)';Resolution='https://mits.au.itglue.com/3210579/docs/3704564';Exists=$false;Closed=$false}
    $Tasks += [Task]@{Notes='Reassign ICT Assets (manual)';Resolution='https://mits.au.itglue.com/3210579/docs/3446859 - step 20';Exists=$false;Closed=$false}
    $Tasks += [Task]@{Notes='MDM - Return Device/s';Resolution='Automation will send an email to MDM with user and recipient details. That email will BCC the service desk, so a new ticket will be created to complete device retrieval.';Exists=$false;Closed=$false}
    $Tasks += [Task]@{Notes='Advise Internal CleanCo Groups';Resolution='Automation will send emails to internal groups and to the P&C team, which will appear as separate tickets to be bundled into this one.';Exists=$false;Closed=$false}
    $Tasks += [Task]@{Notes='Move to CQL-HelpDesk (MS)';Resolution='Once this task is complete, automation should be finished.';Exists=$false;Closed=$false}
} else {
    $Tasks += [Task]@{Notes='Move to HelpDesk (MS)';Resolution='Once this task is complete, automation should be finished.';Exists=$false;Closed=$false}
}

# Add Elston-specific tasks
if ($ElstonTasks) {
    $Tasks += [Task]@{Notes='Provide Screenshot of Last Sign In';Resolution='https://mits.au.itglue.com/797907/docs/4784318';Exists=$false;Closed=$false}
    $Tasks += [Task]@{Notes='Disable IRESS Access (manual)';Resolution='https://mits.au.itglue.com/797907/docs/6198041';Exists=$false;Closed=$false}
    $Tasks += [Task]@{Notes='Disable XPLAN Access (manual)';Resolution='https://mits.au.itglue.com/797907/docs/379148';Exists=$false;Closed=$false}
    $Tasks += [Task]@{Notes='Disable TeamViewer Access (manual)';Resolution='https://mits.au.itglue.com/797907/docs/7223559';Exists=$false;Closed=$false}
}

## FILTER THROUGH TASKS ##

try {
    $Log += "Getting tasks for #$TicketId...`n"
    $TicketTasks = .\CWM-FindTicketTasks.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets | ConvertFrom-Json
    $Log += "SUCCESS: Fetched tasks for #$TicketId.`n`n"
    Write-Warning "SUCCESS: Fetched tasks for #$TicketId."
} catch {
    $Log += "ERROR: Unable to fetch tasks for #$TicketId.`nERROR DETAILS: " + $_
    Write-Error "Unable to fetch tasks for #$TicketId : $_"
    $TicketTasks = $null
}

# Ignore tasks that already exist
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
}

## ADD TICKET TASKS ##

foreach ($Task in $Tasks) {
    if (!$Task.Exists) {
        $TaskNotes = $Task.Notes
        $TaskResolution = $Task.Resolution
        $TaskParameters = @{
            TicketId = $TicketId
            Note = $TaskNotes
            Resolution = $TaskResolution
            ApiSecrets = $ApiSecrets
        }
        $NewTask = .\CWM-AddTicketTask.ps1 @TaskParameters | ConvertFrom-Json
        $Log += $NewTask.Log + "`n`n"
    }
}

## ADD NOTE TO TICKET TO EXPLAIN TASKS ##

$Text = "Create Tasks [automated action]:`n`n"
$Text += "As part of exit user flows, tasks are created for every action that the automation intends to take. "
$Text += "These tasks represent the steps of a exit user request that needs to be completed. "
$Text += "Once the automation has completed a task, it will close it automatically. In some cases, it will add details to the task's resolution.`n`n"
$Text += "In the event that the automation fails (Internal Systems Review Required), the flow will need to be reviewed by an Internal Systems technician. "
$Text += "If one is not available to re-run the flow after fixing it, the tasks are your guide to what is yet to be completed.`n`n"
$Text += "See the company's exit user doco for more information."

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
    $TicketTasks = .\CWM-FindTicketTasks.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets | ConvertFrom-Json
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