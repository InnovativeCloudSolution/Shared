<#

Mangano IT - ConnectWise Manage - Create Multiple Tasks and Note
Created by: Gabriel Nugent
Version: 1.7

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [Parameter(Mandatory=$true)]$TicketId,
    [Parameter(Mandatory=$true)]$TasksToAdd,
    $Text = ''
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

# Array of tasks
$Tasks = @()

# Only add note if tasks were added to the ticket
$TasksAdded = $false

# Initialize class for ticket task
class Task {
    [string]$Notes
    [string]$Resolution
    [boolean]$Exists
    [boolean]$Closed
}

# Define tasks
foreach ($TaskToAdd in $TasksToAdd.split(';')) {
    $TaskDetails = $TaskToAdd.split(':')
    $Tasks += [Task]@{
        Notes = $TaskDetails[0]
        Resolution = $TaskDetails[1]
        Exists = $false
        Closed = $false
    }
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
}

# Add extra line break
$Log += "`n"

## ADD TICKET TASKS ##

foreach ($Task in $Tasks) {
    if (!$Task.Exists) {
        $TaskNotes = $Task.Notes
        $TaskResolution = $Task.Resolution
        try {
            $Log += "`nCreating task '$TaskNotes - $TaskResolution'..."
            $TaskArguments = @{
                parentId = $TicketId
                notes = $Task.Notes
                resolution = $Task.Resolution
                closedFlag = $false
            }
            New-CWMTicketTask @TaskArguments | Out-Null
            $TasksAdded = $true
        }
        catch {
            $Log += "`nERROR: Ticket task '$TaskNotes - $TaskResolution' not created.`nERROR DETAILS: " + $_
            Write-Error "Ticket task '$TaskNotes - $TaskResolution' not created : $_"
        }
    }
}

## ADD NOTE TO TICKET TO EXPLAIN TASKS ##

if ($TasksAdded) {
    if ($Text -eq '' -or $null -eq $Text) {
        $Text = "Create Tasks [automated action]:`n`n"
        $Text += "Tasks have been added with important additional steps that need to be taken. Please remember to complete them as you go."
    }
    $NoteArguments = @{
        TicketId = $TicketId
        Text = $Text
        InternalFlag = $true
    }
    .\CWM-AddTicketNote.ps1 @NoteArguments | Out-Null
}

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