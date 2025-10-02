<#

Elston - Setup Teams Calling
Created by: Gabriel Nugent
Version: 1.3.4

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$EmailAddress,
    [string]$Site,
    [int]$TicketId,
    $ApiSecrets = $null
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''
$Result = $false
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
$TenantId = "2939c5b6-63d7-430c-a345-aba7b3d6ab1b"

# Teams calling arrays
$Names            = @("Brisbane", "Gold Coast", "Hervey Bay", "Canberra", "Ballina", "Sydney")
$PhoneBlockLower  = @(61730023810, 61755573010, 61731513310, 61261534010, 61279032510, 61256462010)
# $PhoneBlockHigher = @(61730023899, 61755573099, 61731513399, 61261534099, 61279032599, 61256462099)
$PhoneNumberStart = @(617300238, 617555730, 617315133, 612615340, 612790325, 612564620)
$CallerIds        = @("bris1", "gold1", "herv1", "canb1", "ball1", "sydn1")
$DialingPlans     = @("QLDDialOut","QLDDialOut","QLDDialOut","NSWDialOut","NSWDialOut","NSWDialOut")
$SiteNumber = [array]::indexof($Names, $Site)
$Log += "INFO: Site index selected: $SiteNumber - "+$Names[$SiteNumber]+"`n`n"

## CHECK IF TEAMS CALLING IS ALREADY DONE ##

$TicketTasks = .\CWM-FindTicketTasks.ps1 -TicketId $TicketId -TaskNotes $TaskNotes

if (!$TicketTasks.closedFlag) {
    ## CONNECT TO TEAMS ##
    
    try {
        $Log += "Connecting to Teams Online...`n"
        Connect-MicrosoftTeams -CertificateThumbprint $CertificateThumbprint -ApplicationId $ApplicationId -TenantId $TenantId | Out-Null
        $Log += "SUCCESS: Connected to Teams Online.`n`n"
    }
    catch {
        $Log += "ERROR: Failed to connect to Teams Online.`nERROR DETAILS: " + $_
        Write-Error "Failed to connect to Teams Online : $_"
    }
    
    ## ASSIGN PHONE DETAILS ##
    
    # Assign a phone number 
    $Log += "Picking phone number...`n"
	$GetNumbersParameters = @{
		TelephoneNumberStartsWith = $PhoneNumberStart[$SiteNumber]
		ActivationState = 'Activated'
		NumberType = 'CallingPlan'
		PstnAssignmentStatus = 'Unassigned'
		CapabilitiesContain = 'UserAssignment'
	}
    $PhoneNumbers = Get-CsPhoneNumberAssignment @GetNumbersParameters
    foreach ($Number in $PhoneNumbers) {
		$TelephoneInt = $Number.TelephoneNumber -replace '[^0-9]', ''
        if ($TelephoneInt -gt $PhoneBlockLower[$SiteNumber]) {
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
    
    # Assign dialling plan
    $DialingPlan = $DialingPlans[$SiteNumber]
    $Log += "Assigning dial plan $DialingPlan to $EmailAddress...`n"
    try {
        Grant-CsTenantDialPlan -Identity $EmailAddress -PolicyName $DialingPlan
        $Log += "SUCCESS: Assigned dial plan $DialingPlan to $EmailAddress.`n`n"
        Write-Warning "SUCCESS: Assigned dial plan $DialingPlan to $EmailAddress."
        $Result_DialPlan = $true
    }
    catch {
        $Log += "ERROR: Failed to assign dial plan $DialingPlan to user.`nERROR DETAILS: " + $_
        Write-Error "Failed to assign dial plan $DialingPlan to user : $_"
		$Result_DialPlan = $false
    }
    
    # Assign Calling ID Policy
    $CallerId = $CallerIds[$SiteNumber]
    try {
        $Log += "Assigning calling ID policy $CallerId to $EmailAddress...`n"
        Grant-CsCallingLineIdentity -Identity $EmailAddress -PolicyName $CallerId
        $Log += "SUCCESS: Assigned calling ID policy $CallerId to $EmailAddress.`n`n"
        Write-Warning "SUCCESS: Assigned calling ID policy $CallerId to $EmailAddress."
        $Result_CallingId = $true
    }
    catch {
        $Log += "ERROR: Failed to assign calling ID policy.`nERROR DETAILS: " + $_
        Write-Error "Failed to assign calling ID policy : $_"
		$Result_CallingId = $true
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