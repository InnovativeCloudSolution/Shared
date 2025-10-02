<#

Mangano IT - ConnectWise Manage - Create ConnectWise Manage PO Report
Created by: Alex Williams, Gabriel Nugent
Version: 1.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$CompanyIdentifier = "MIT",
    [string]$PONumber = '',
    [string]$AgreementType = 'All',
    [string]$ReportSpan = 'Forever',
    [string]$Requestor = ''
)

Set-StrictMode -Off

## SCRIPT VARIABLES ##

$Results = @()
$PONumber = $PONumber -replace '\s',''
$Conditions = 'company/identifier like "%'+$CompanyIdentifier+'%"'
#$AEST = [System.TimeZoneInfo]::GetSystemTimeZones().Where({$_.StandardName -eq 'AUS Eastern Standard Time'})[0]
$SendGridApiKey = Get-AutomationVariable -Name 'SendGrid Azure Automation API'

## CONNECT TO CWM ##

# Set CRM variables
$Server = 'https://api-aus.myconnectwise.net'
$Company = 'mit'
$PublicKey = Get-AutomationVariable -Name 'PublicKey'
$PrivateKey = Get-AutomationVariable -Name 'PrivateKey'
$ClientId = Get-AutomationVariable -Name 'clientId'

# Create a credential object
$Connection = @{
    Server = $Server
    Company = $Company
    pubkey = $PublicKey
    privatekey = $PrivateKey
    clientId = $ClientId
}

# Connect to Manage server
try { Connect-CWM @Connection } catch { Write-Error "CONNECT TO CWM: $_" }

## TIME ENTRY MANAGEMENT FUNCTION ##

function Format-TimeEntry {
    param (
        $TimeEntry
    )

    $Cost = $TimeEntry.actualHours * $TimeEntry.hourlyRate
    if (![string]::IsNullOrEmpty($TimeEntry.notes)) {
        if ($TimeEntry.notes[0] -eq '-') { $TimeEntry.notes = $TimeEntry.notes.substring(1) }
        $TimeEntry.notes =$TimeEntry.notes -replace '\n',' '
        if ($TimeEntry.notes.length -gt 32000) { $TimeEntry.notes = $TimeEntry.notes.substring(0,32000) }
    }
    $TimeDate = [datetime]::parseexact($TimeEntry.timeStart, 'yyyy-MM-ddTHH:mm:ssZ', $null).AddHours(10).ToString('yyyy-MM-dd');

    $Ticket = Get-CWMTicket -id $TimeEntry.chargeToId
    $Type = $Ticket.type.name
    $Subtype = $Ticket.subType.name
    $Item = $Ticket.item.name
    $FormattedTimeEntry = New-Object -TypeName PSObject -Property ([ordered]@{
        TimeID = $TimeEntry.id
        TicketNumber = $TimeEntry.chargeToId
        Summary = $TimeEntry.ticket.Summary
        Technician = $TimeEntry.member.name
        timeDate = $TimeDate
        hoursBilled = $TimeEntry.hoursBilled
        Cost = $Cost
        Notes = $TimeEntry.notes
        Agreement = $TimeEntry.agreement.name
        Type = $Type
        subType = $Subtype
        item = $Item
        status = $Ticket.status.name
    })

    Write-Output $FormattedTimeEntry
}

## ESTABLISH REQUEST ##

if ($AgreementType -eq 'Managed Services') { $Conditions += " and agreement/name like '%Managed Services%'" }
elseif ($AgreementType -eq "PO") { $Conditions +=' and agreement/name like "*'+$PONumber+'*"' }

if ($AgreementType -eq 'Managed Services' -or $AgreementType -eq 'All') {
    if ($ReportSpan -eq '1 Month') {
        $ReportStart = (Get-Date).AddMonths(-1).AddHours(-10).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $Conditions += " and lastUpdated >= [$ReportStart]"
    } elseif ($ReportSpan -eq '3 Months') {
        $ReportStart = (Get-Date).AddMonths(-3).AddHours(-10).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $Conditions += " and lastUpdated >= [$ReportStart]"
    } elseif ($ReportSpan -eq '6 Months') {
        $ReportStart = (Get-Date).AddMonths(-6).AddHours(-10).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $Conditions += " and lastUpdated >= [$ReportStart]"
    } else {
        $ReportStart = (Get-Date).AddMonths(-6).AddHours(-10).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $Conditions += " and lastUpdated >= [$ReportStart]"
    }
}

Write-Output $Conditions

## FIND TIME ENTRIES ##

# Fetch time entries
try { $TimeEntries = Get-CWMTimeEntry -Condition $Conditions -all } catch { Write-Error "FETCH TIME ENTRIES: $_" }

# If the time entry count is over 999, it splits it into an array of arrays. This will check if we have an array of arrays of objects or an array of objects
if ($TimeEntries.GetType() -eq $TimeEntries[0].GetType()) {
    foreach ($Subset in $TimeEntries) {
        foreach ($TimeEntry in $Subset) {
            try { $Results += Format-TimeEntry($TimeEntry) } catch { Write-Error "FORMAT TIME ENTRY: $_" }
        }
    }
} else {
    foreach ($TimeEntry in $TimeEntries) {
        try { $Results += Format-TimeEntry($TimeEntry) } catch { Write-Error "FORMAT TIME ENTRY: $_" }
    }
}
$Now = (Get-Date).ToString('yyyy-MM-dd')
$FileName = "$CompanyIdentifier $AgreementType$PONumber $ReportSpan Report - $Now"

# Exports the CSV to be emailed
try { $Results | Export-Csv "C:\temp\$FileName.csv" -NoTypeInformation } catch { Write-Error "EXPORT CSV: $_" }

## SEND RESULT TO RECIPIENT ##

$Subject = "Agreement Reconciliation: "+$FileName
$Body = "Please find attached the agreement reconciliation report for "+$FileName
$SendGridEmail = @{
	From = 'workflows@manganoit.com.au'
	To = $Requestor
	Subject = $Subject
	Body = $Body
	SmtpServer = 'smtp.sendgrid.net'
	Port = 587
	UseSSL = $true
	Credential = New-Object PSCredential 'apikey', (ConvertTo-SecureString $SendGridApiKey -AsPlainText -Force)	
    Attachments = "C:\temp\$FileName.csv"
}

# Sends a message with the SendGrid API
try { Send-MailMessage @SendGridEmail } catch { Write-Error "SEND MAIL MESSAGE: $_" }

# Disconnect from ConnectWise Manage
try { Disconnect-CWM } catch { Write-Error "DISCONNECT FROM CWM: $_" }