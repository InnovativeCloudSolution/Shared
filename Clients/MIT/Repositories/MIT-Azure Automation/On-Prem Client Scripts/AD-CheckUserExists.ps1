<#

Mangano IT - Active Directory - Check User Exists in AD
Created by: Gabriel Nugent
Version: 1.2

This runbook is designed to be used in conjunction with a Power Automate flow.

#>

param(
    [string]$GivenName,
	[string]$Surname,
    [string]$SamAccountName,
    [string]$DisplayName,
    [bool]$ReturnJson = $false
)

## SCRIPT VARIABLES ##

# Track status of automation
[string]$Log = ''

#Output the status variable content to a file
$Date = Get-Date -Format "dd-MM-yyyy HHmm"
$FilePath = "C:\Scripts\Logs\AD-CheckUserExists"
if ($DisplayName -eq '') {
    $FileName = "$GivenName $Surname-$Date.txt"
} else {
    $FileName = "$DisplayName-$Date.txt"
}

## CHECK IF USER EXISTS ##
try {
    $Log += "Attempting to locate user...`n"
    if ($SamAccountName -ne '') {
        $User = Get-ADUser -Filter * | Where-Object {$_.SamAccountName -eq $SamAccountName}
    }
    elseif ($GivenName -ne '' -and $Surname -ne '') {
        $User = Get-ADUser -Filter * | Where-Object {$_.GivenName -eq $GivenName -and $_.Surname -eq $Surname}
    } else {
        $User = Get-ADUser -Filter * | Where-Object {$_.Name -eq $DisplayName}
    }
    
    if ($null -ne $User) {
        $UserPrincipalName = $User.userPrincipalName
        $Log += "SUCCESS: $UserPrincipalName has been located."
        Write-Warning "SUCCESS: $UserPrincipalName has been located."
        $Result = $true 
    }
    else {
        $Log += "INFO: An account for $GivenName $Surname has not been located."
        Write-Warning "INFO: An account for $GivenName $Surname has not been located."
        $Result = $false
    }
} catch {
    $Log += "ERROR: Unable to run check for user.`nERROR DETAILS: " + $_
    Write-Error "Unable to run check for user : $_"
}

## SEND DETAILS BACK TO FLOW ##

$Output = @{
    Result = $Result
    User = $User
    Log = $Log
}

if ($ReturnJson) {
    Write-Output $Output | ConvertTo-Json -Depth 100
} else {
    Write-Output $Output
}

# Makes folder for logs and outputs logs
.\CreateLogFile.ps1 -Log $Log -FilePath $FilePath -FileName $FileName | Out-Null