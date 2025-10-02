<#

Mangano IT - Teams - Set Up Calling
Created by: Gabriel Nugent
Version: 1.2.5

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$EmailAddress,
    [int]$SiteId,
    [int]$TicketId,
    [string]$TenantId,
    [array]$TeamsCallingSites, # JSON array of hashtables with site name, phone block, caller ID, 
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''
$Result = $true
$TaskNotes = "Setup Teams Calling"

## GET CW MANAGE CREDENTIALS ##

if ($null -eq $ApiSecrets) {
    $ApiSecrets = .\CWM-GetApiSecretsFromVault.ps1 | ConvertFrom-Json
}

# Fetch tenant slug
$TenantSlug = .\CWM-FindCompanySlug.ps1 -TicketId $TicketId -ApiSecrets $ApiSecrets

# Grab connection variables
$ApplicationId = .\KeyVault-GetSecret.ps1 -SecretName "$TenantSlug-TMS-ApplicationID"
$CertificateThumbprint = .\KeyVault-GetSecret.ps1 -SecretName "$TenantSlug-TMS-CertificateThumbprint"

## SORT THROUGH ARRAY OF HASHTABLES ##

foreach ($TeamsCallingSite in $TeamsCallingSites) {
    if ($TeamsCallingSite.SiteId -eq $SiteId) {
        Write-Warning "Picked Teams site: $($TeamsCallingSite)"
        $TeamsSite = $TeamsCallingSite
    }
}

## CHECK IF TEAMS CALLING IS ALREADY DONE ##

$TicketTasks = .\CWM-FindTicketTasks.ps1 -TicketId $TicketId -TaskNotes $TaskNotes

if (!$TicketTasks.closedFlag -and $null -ne $TicketTasks) {
    ## CONNECT TO TEAMS ##
    
    try {
        $Log += "Connecting to Teams Online...`n"
        Connect-MicrosoftTeams -CertificateThumbprint $CertificateThumbprint -ApplicationId $ApplicationId -TenantId $TenantId | Out-Null
        $Log += "SUCCESS: Connected to Teams Online.`n`n"
    } catch {
        $Log += "ERROR: Failed to connect to Teams Online.`nERROR DETAILS: " + $_
        Write-Error "Failed to connect to Teams Online : $_"
    }
    
    ## ASSIGN PHONE DETAILS ##
    
    # Assign a phone number 
    $Log += "Picking phone number...`n"
	$GetNumbersParameters = @{
		TelephoneNumberStartsWith = $TeamsSite.PhoneNumberStart
		ActivationState = 'Activated'
		NumberType = 'CallingPlan'
		PstnAssignmentStatus = 'Unassigned'
		CapabilitiesContain = 'UserAssignment'
	}
    $PhoneNumbers = Get-CsPhoneNumberAssignment @GetNumbersParameters
    foreach ($Number in $PhoneNumbers) {
		$TelephoneInt = $Number.TelephoneNumber -replace '[^0-9]', ''
        if ($TelephoneInt -gt $TeamsSite.PhoneBlockLower) {
            $OfficePhone = $Number.TelephoneNumber
            $Log += "SUCCESS: Found phone number: $OfficePhone`n"
            Write-Warning "SUCCESS: Found phone number: $OfficePhone"
            break
        }
    }
    
    try {
        Set-CsPhoneNumberAssignment -Identity $EmailAddress -PhoneNumber $OfficePhone -PhoneNumberType CallingPlan
        $Log += "SUCCESS: Added office phone to $EmailAddress in Teams.`n`n"
        Write-Warning "SUCCESS: Added office phone to $EmailAddress in Teams."
        $Result_PhoneNumber = $true
    } catch {
        $Log += "ERROR: Failed to add office phone to $EmailAddress in Teams.`nERROR DETAILS: " + $_
        Write-Error "Failed to add office phone to $EmailAddress in Teams : $_"
		$Result_PhoneNumber = $false
    }
    
    # Complete next steps only if a phone number was assigned
    if ($Result_PhoneNumber) {
        # Assign dialling plan
        $Log += "Assigning dial plan $($TeamsSite.DialingPlan) to $EmailAddress...`n"
        try {
            Grant-CsTenantDialPlan -Identity $EmailAddress -PolicyName $TeamsSite.DialingPlan
            $Log += "SUCCESS: Assigned dial plan $($TeamsSite.DialingPlan) to $EmailAddress.`n`n"
            Write-Warning "SUCCESS: Assigned dial plan $($TeamsSite.DialingPlan) to $EmailAddress."
            $Result_DialPlan = $true
        } catch {
            $Log += "ERROR: Failed to assign dial plan $($TeamsSite.DialingPlan) to user.`nERROR DETAILS: " + $_
            Write-Error "Failed to assign dial plan $($TeamsSite.DialingPlan) to user : $_"
            $Result_DialPlan = $false
        }
        
        # Assign Calling ID Policy
        try {
            $Log += "Assigning calling ID policy $($TeamsSite.CallerId) to $EmailAddress...`n"
            Grant-CsCallingLineIdentity -Identity $EmailAddress -PolicyName $TeamsSite.CallerId
            $Log += "SUCCESS: Assigned calling ID policy $($TeamsSite.CallerId) to $EmailAddress.`n`n"
            Write-Warning "SUCCESS: Assigned calling ID policy $($TeamsSite.CallerId) to $EmailAddress."
            $Result_CallingId = $true
        } catch {
            $Log += "ERROR: Failed to assign calling ID policy $($TeamsSite.CallerId).`nERROR DETAILS: " + $_
            Write-Error "Failed to assign calling ID policy $($TeamsSite.CallerId) : $_"
            $Result_CallingId = $false
        }
    }
    
    $Log += "INFO: Teams calling additions complete. Removing PSSession.."
    Disconnect-MicrosoftTeams | Out-Null

    if ($Result_PhoneNumber -and $Result_DialPlan -and $Result_CallingId) {
        $Result = $true
    }
    
    ## CLOSE TASK AND ADD NOTE ##
    
    if ($TicketId -ne 0) {
        # Close task if all three steps worked
        if ($Result_PhoneNumber -and $Result_DialPlan -and $Result_CallingId) {
            $Task = .\CWM-UpdateTask.ps1 -TicketId $TicketId -Note $TaskNotes -ClosedStatus $true -ApiSecrets $ApiSecrets
            $Log += "`n`n" + $Task.Log
        }
    
        $NoteText = "The following has been set up in Teams Calling for $EmailAddress.`n`n"
        $NoteText += "Office phone: $OfficePhone`n`nPhone number assigned: $Result_PhoneNumber`n"
        $NoteText += "Dial plan assigned: $Result_DialPlan`nCalling ID assigned: $Result_CallingId"
    
        # Add note to ticket
        $Note = .\CWM-AddTicketNote.ps1 -TicketId $TicketId -Text $NoteText -ResolutionFlag $true -ApiSecrets $ApiSecrets
        $Log += "`n`n" + $Note.Log
    }
}

## SEND DETAILS TO FLOW ##
    
$Output = @{
    Result = $Result
    Result_PhoneNumber = $Result_PhoneNumber
    OfficePhone = $OfficePhone
    Log = $Log
}

# Send back to Power Automate
Write-Output $Output | ConvertTo-Json