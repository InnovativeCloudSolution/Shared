<#

Mangano IT - Delete WDAC Policies
Created by: Gabriel Nugent
Version: 1.1

Remember to unassign WDAC policies in Intune, and restart after!

#>

# Define environment variables
$MountPoint = 'C:\EFIMount'
$DestinationFolder = $env:windir+"\System32\CodeIntegrity\CiPolicies"
$EFIDestinationFolder = "$MountPoint\EFI\Microsoft\Boot\CiPolicies"
$LogsPath = "C:\Logs"

# Logs variable and date variable for log name
$Logs = ''
$Date = Get-Date -Format FileDateTime

# Checks to see if the policy path exists
if (!(Test-Path $DestinationFolder)) {
    try {
        Write-Host "`n$DestinationFolder does not exist, creating now."
        $Logs += "`n$DestinationFolder does not exist, creating now."
        New-Item -Path $DestinationFolder -ItemType "directory"
    } catch {
        Write-Host "ERROR: $DestinationFolder was not able to be created. Are you running this script as an elevated user?"
        $Logs += "ERROR: $DestinationFolder was not able to be created. Are you running this script as an elevated user?"
        exit
    }
}
else { $Logs += "`n$DestinationFolder exists." }

# Deletes the policy folders
try {
    Write-Host "`nDeleting policy folders..."
    $Logs += "`nDeleting policy folders..."
    Remove-Item "$DestinationFolder\Active" -Recurse
    Remove-Item "$DestinationFolder\Staged" -Recurse
} catch {
    Write-Host "ERROR: Unable to delete policy folders from $DestinationFolder. Are you running this script as an elevated user?"
    $Logs += "ERROR: Unable to delete policy folders from $DestinationFolder. Are you running this script as an elevated user?"
    exit
}

# Checks to make sure the target mounting folder exists
if (!(Test-Path $MountPoint)) {
    try {
        Write-Host "`n$MountPoint does not exist, creating now."
        $Logs += "`n$MountPoint does not exist, creating now."
        New-Item -Path $MountPoint -ItemType "directory"
    } catch {
        Write-Host "ERROR: $MountPoint was not able to be created. Are you running this script as an elevated user?"
        $Logs += "ERROR: $MountPoint was not able to be created. Are you running this script as an elevated user?"
        exit
    }
}

# Mounts EFI partition that holds policies, and copies policy in
try {
    Write-Host "Attempting to mount $MountPoint..."
    $Logs += "Attempting to mount $MountPoint..."
    $EFIPartition = (Get-Partition | Where-Object IsSystem).AccessPaths[0]
    mountvol $MountPoint $EFIPartition
} catch {
    Write-Host "ERROR: Unable to mount $MountPoint. Are you running this script as an elevated user?"
    $Logs += "ERROR: Unable to mount $MountPoint. Are you running this script as an elevated user?"
    exit
}

# Check to make sure destination folder exists
if (!(Test-Path $EFIDestinationFolder)) {
    try {
        Write-Host "`n$EFIDestinationFolder does not exist, creating now."
        $Logs += "`n$EFIDestinationFolder does not exist, creating now."
        New-Item -Path $EFIDestinationFolder -ItemType "directory"
    } catch {
        Write-Host "ERROR: $EFIDestinationFolder was not able to be created. Are you running this script as an elevated user?"
        $Logs += "ERROR: $EFIDestinationFolder was not able to be created. Are you running this script as an elevated user?"
        exit
    }
}
else { $Logs += "`n$EFIDestinationFolder exists." }

# Deletes the EFI policy folders
try {
    Write-Host "`nDeleting policy folders..."
    $Logs += "`nDeleting policy folders..."
    Remove-Item "$EFIDestinationFolder\Active" -Recurse
    Remove-Item "$EFIDestinationFolder\Staged" -Recurse
} catch {
    Write-Host "ERROR: Unable to delete policy folders from $EFIDestinationFolder. Are you running this script as an elevated user?"
    $Logs += "ERROR: Unable to delete policy folders from $EFIDestinationFolder. Are you running this script as an elevated user?"
    exit
}

# Unmount EFI directory
try {
    Write-Host "`nUnmounting $EFIPartition from $MountPoint, please wait..."
    $Logs += "`nUnmounting $EFIPartition from $MountPoint, please wait..."
    mountvol $MountPoint /d
    Write-Host "Unmounting complete. Please restart your device to finish removing WDAC policies."
    $Logs += "Unmounting complete. Please restart your device to finish removing WDAC policies."
} catch {
    Write-Host "ERROR: Unable to unmount partition. Check to see if the partition clears itself up post-restart."
    $Logs += "ERROR: Unable to unmount partition. Check to see if the partition clears itself up post-restart."
    exit
}

# Makes folder for logs and outputs logs
if (!(Test-Path -PathType container $LogsPath)) { New-Item -ItemType Directory -Path $LogsPath }
$Logs | Out-File -FilePath "$LogsPath\Remove WDAC $Date.txt" -Force -Confirm:$false