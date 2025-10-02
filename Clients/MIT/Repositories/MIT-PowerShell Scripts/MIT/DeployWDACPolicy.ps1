# Define standard policy paths
$BasePolicyPath = "C:\Policies\Base Policy\{963DC7A2-B4FA-44A9-A194-8A3617B2A8EF}.cip"
$MITAZU1MON1PolicyPath = "C:\Policies\MIT-AZU1-MON1\SiPolicy.p7b"
$MITAZU1RMM01PolicyPath = "C:\Policies\MIT-AZU1-RMM01\{380889E5-B199-408E-94DD-1379DFD839A1}.cip"
$MITAZU1RMM02PolicyPath = "C:\Policies\MIT-AZU1-RMM02\{CBCB7134-01C5-4C0B-8355-6752BEE76888}.cip"
$MITHMLN1AHW01PolicyPath = "C:\Policies\MIT-HMLN1-AHW01\SiPolicy.p7b"
$MITHMLN1APP01PolicyPath = "C:\Policies\MIT-HMLN1-APP01\SiPolicy.p7b"
$MITHMLN1HV01PolicyPath = "C:\Policies\MIT-HMLN1-HV01\SiPolicy.p7b"
$MITHMLN1HV02PolicyPath = "C:\Policies\MIT-HMLN1-HV02\SiPolicy.p7b"
$MITHMLN1MGT01PolicyPath = "C:\Policies\MIT-HMLN1-MGT01\MIT - Servers - MIT-HMLN1-MGT01.bin"

# Define policy refresh path, destination folder, and mount point
$RefreshPolicyTool = "C:\Policies\RefreshPolicy(AMD64).exe"
$DestinationFolder = $env:windir+"\System32\CodeIntegrity\CIPolicies\Active\"
$DestinationFolderLegacy = $env:windir+"\System32\CodeIntegrity\"
$DestinationBinaryLegacy = $env:windir+"\System32\CodeIntegrity\SiPolicy.p7b"
$MountPoint = 'C:\EFIMount'
$EFIDestinationFolder = "$MountPoint\EFI\Microsoft\Boot\CiPolicies\Active"

# Select policy to copy
$PolicyToApply = ''

## Program details - please remember to update the version number when making changes
Write-Host "`nMangano IT - Deploy WDAC Policies to Internal Servers" -ForegroundColor Yellow
Write-Host "Version: " -ForegroundColor yellow -NoNewLine; Write-Host "2.1"
Write-Host "Created by: " -ForegroundColor yellow -NoNewLine; Write-Host "Gabriel Nugent"

try {
    Write-Host "`nGrabbing PC name..."
    $ComputerInfo = Get-ComputerInfo
    Write-Host "PC name is" $ComputerInfo.CsName
} catch {
    Write-Error "Unable to get PC name."
    $ComputerInfo.CsName = $(Write-Host "Please provide PC name manually: " -ForegroundColor yellow -NoNewLine; Read-Host)
}

switch ($ComputerInfo.CsName) {
    "MIT-AZU1-MON1" { $PolicyToApply = $MITAZU1MON1PolicyPath }
    "MIT-AZU1-RMM01" { $PolicyToApply = $MITAZU1RMM01PolicyPath }
    "MIT-AZU1-RMM02" { $PolicyToApply = $MITAZU1RMM02PolicyPath }
    "MIT-HMLN1-AHW01" { $PolicyToApply = $MITHMLN1AHW01PolicyPath }
    "MIT-HMLN1-APP01" { $PolicyToApply = $MITHMLN1APP01PolicyPath }
    "MIT-HMLN1-HV01" { $PolicyToApply = $MITHMLN1HV01PolicyPath }
    "MIT-HMLN1-HV02" { $PolicyToApply = $MITHMLN1HV02PolicyPath }
    "MIT-HMLN1-MGT01" { $PolicyToApply = $MITHMLN1MGT01PolicyPath }
    default {
        Write-Host "`n"$ComputerInfo.CsName "was not found in the list of device names. Please modify the script to include this new device and its policy path."
        exit
    }
}

# Grabs the version of Windows to determine what method needs to be used for copying the policy
try {
    Write-Host "`nGrabbing version of Windows..."
    Write-Host "Windows version is" $ComputerInfo.WindowsVersion
} catch {
    Write-Error "Unable to get Windows version info."
    $ComputerInfo.WindowsVersion = $(Write-Host "Please provide four-digit version number (e.g. 1809) manually: " -ForegroundColor yellow -NoNewLine; Read-Host)
}

if ($ComputerInfo.WindowsVersion -ge 1903) {
    # Checks to see if the policy path exists
    if (!(Test-Path $DestinationFolder)) {
        try {
            Write-Host "`n$DestinationFolder does not exist, creating now."
            New-Item -Path $DestinationFolder -ItemType "directory"
        } catch {
            Write-Error "$DestinationFolder was not able to be created. Are you running this script as an elevated user?"
            exit
        }
    }
    else { Write-Host "`n$DestinationFolder exists." }

    # Copies the required policy to the policy folder, then refreshes applied policies
    try {
        Write-Host "`nCopying policy to $DestinationFolder and refreshing policy..."
        Copy-Item -Path $PolicyToApply -Destination $DestinationFolder -Force
        & $RefreshPolicyTool
    } catch {
        Write-Error "Unable to copy policy to $DestinationFolder or refresh policy. Are you running this script as an elevated user?"
        exit
    }
}

else {
    # Checks to see if the policy path exists
    if (!(Test-Path $DestinationFolderLegacy)) {
        try {
            Write-Host "`n$DestinationFolderLegacy does not exist, creating now."
            New-Item -Path $DestinationFolderLegacy -ItemType "directory"
        } catch {
            Write-Error "$DestinationFolderLegacy was not able to be created. Are you running this script as an elevated user?"
            exit
        }
    }
    else { Write-Host "`n$DestinationFolderLegacy exists." }

    # Copies the required policy to the policy folder
    try {
        Write-Host "`nCopying policy to $DestinationFolderLegacy..."
        Copy-Item -Path $PolicyToApply -Destination $DestinationFolderLegacy -Force
    } catch {
        Write-Error "Unable to copy policy to $DestinationFolderLegacy. Are you running this script as an elevated user?"
        exit
    }

    # Applies new policy
    try {
        Write-Host "`nAttempting to refresh Code Integrity policy..."
        $ApplyPolicy = Invoke-CimMethod -Namespace root\Microsoft\Windows\CI -ClassName PS_UpdateAndCompareCIPolicy -MethodName Update -Arguments @{FilePath = $DestinationBinaryLegacy}
        if ($ApplyPolicy.cmdletOutput -eq 1) { Write-Host "Code Integrity policy has been refreshed correctly." }
        else {
            Write-Host "Policy has not been applied correctly. Is the SiPolicy.p7b file in the right folder?"
            exit
        }
    } catch {
        Write-Error "Unable to refresh Code Integrity policy."
        exit
    }
}

# Checks to make sure the target mounting folder exists
if (!(Test-Path $MountPoint)) {
    try {
        Write-Host "`n$MountPoint does not exist, creating now."
        New-Item -Path $MountPoint -ItemType "directory"
    } catch {
        Write-Error "$MountPoint was not able to be created. Are you running this script as an elevated user?"
        exit
    }
}

# Mounts EFI partition that holds policies, and copies policy in
try {
    Write-Host "Attempting to mount $MountPoint..."
    $EFIPartition = (Get-Partition | Where-Object IsSystem).AccessPaths[0]
    mountvol $MountPoint $EFIPartition
} catch {
    Write-Error "Unable to mount $MountPoint. Are you running this script as an elevated user?"
    exit
}

# Check to make sure destination folder exists
if (!(Test-Path $EFIDestinationFolder)) {
    try {
        Write-Host "`n$EFIDestinationFolder does not exist, creating now."
        New-Item -Path $EFIDestinationFolder -ItemType "directory"
    } catch {
        Write-Error "$EFIDestinationFolder was not able to be created. Are you running this script as an elevated user?"
        exit
    }
}
else { Write-Host "`n$EFIDestinationFolder exists." }

# Copy policy to EFI directory
try {
    Write-Host "`nCopying policy to $EFIDestinationFolder..."
    Copy-Item -Path $PolicyToApply -Destination $EFIDestinationFolder -Force
    Write-Host "Policy update complete - pre-1903 devices will need to be restarted for it to take effect."
} catch {
    Write-Error "Unable to copy policy to $EFIDestinationFolder. Are you running this script as an elevated user?"
    exit
}

# Unmount EFI directory
try {
    Write-Host "`nUnmounting $EFIPartition from $MountPoint, please wait..."
    mountvol $MountPoint /d
    Write-Host "Unmounting complete."
} catch {
    Write-Error "Unable to unmount partition. Check to see if the partition clears itself up post-restart."
    exit
}