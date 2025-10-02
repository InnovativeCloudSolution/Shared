<#

Mangano IT - TransmitSMS - Send Message
Created by: Gabriel Nugent
Version: 1.2.1

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [int]$TicketId,
    [Parameter(Mandatory=$true)][string]$MobileNumber,
    [Parameter(Mandatory=$true)][string]$Message,
    [string]$CountryCode = 'AU',
    [string]$SenderName = 'Mangano IT',
    [string]$ScheduledSendDate
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

# Format phone
$FormatNumberParams = @{
    PhoneNumber = $MobileNumber
    IsMobileNumber = $true
    IncludePlus = $false
    KeepSpaces = $false
}
$MobileNumber = .\SMS-FormatValidPhoneNumber.ps1 @FormatNumberParams

## GET API VARIABLES ##

$PublicKey = .\KeyVault-GetSecret.ps1 -Name 'MIT-TransmitSMS-Username'
$PrivateKey = .\KeyVault-GetSecret.ps1 -Name 'MIT-TransmitSMS-Password'

# Build authentication
$Credentials = "$($PublicKey):$($PrivateKey)"
$EncodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($Credentials))
$BasicAuthentication = "Basic $EncodedCredentials"

# Build arguments
$ApiArguments = @{
    Uri = "https://api.transmitsms.com/send-sms.json?message=$Message&to=$MobileNumber&from=$SenderName"
    Method = 'POST'
    Headers = @{ Authorization = $BasicAuthentication }
    UseBasicParsing = $true
}

# If country code supplied, add to request
if ($CountryCode -ne '') {
    $ApiArguments.Uri += "&countrycode=$CountryCode"
}

# If scheduled send date provided, add to request
if ($ScheduledSendDate -ne '') {
    $ApiArguments.Uri += "&send_at=$ScheduledSendDate"
}

## API REQUEST ##

try {
    $Log += "Sending message to $MobileNumber...`n"
    Invoke-WebRequest @ApiArguments | Out-Null
    $Log += "SUCCESS: Message sent!"
    Write-Warning "SUCCESS: Message sent!"
    $Result = $true
} catch {
    $Log += "ERROR: Message not sent.`nERROR DETAILS: " + $_
    Write-Error "Message not sent : $_"
    $Result = $false
}

## UPDATE RELATED TASK IF TICKET ID PROVIDED ##

if ($TicketId -ne 0 -and $Result) {
    # Update task with info
    $Task = .\CWM-UpdateTask.ps1 -TicketId $TicketId -Note 'Send SMS' -ClosedStatus $true
    $Log += "`n`n" + $Task.Log

    # Add note to ticket
    $Note = .\CWM-AddTicketNote.ps1 -TicketId $TicketId -Text "Send SMS: SMS has been sent to $MobileNumber." -ResolutionFlag $true
    $Log += "`n`n" + $Note.Log
}

## SEND DETAILS TO FLOW ##

$Output = @{
    Result = $Result
    MobileNumber = $MobileNumber
    Message = $Message
    Log = $Log
}

Write-Output $Output | ConvertTo-Json