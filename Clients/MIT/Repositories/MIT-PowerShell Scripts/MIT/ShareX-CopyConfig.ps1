<#

Mangano IT - Copy ShareX Config to Documents
Created by: Gabriel Nugent & Liam Adair
Version: 1.1

#>

# Logging variables
$LogsPath = "C:\Logs"
$Logs = ''

# Get all user folders
$Logs += "Checking all user folders`n"
$UserFolders = Get-ChildItem -Path C:\Users
$Logs += "User folders:`n$UserFolders`n`n"

# For all non-public users, copy ShareX config file
foreach ($UserFolder in $UserFolders) {
    if ($UserFolder.Name -ne "Public") {
        $UserFolderName = $UserFolder.Name
        $Logs += "Copying config file to C:\Users\$UserFolderName\OneDrive - Mangano IT\Documents`n"
        Copy-Item ".\ShareX-InternalSystemsApprovedConfig.sxb" -Destination "C:\Users\$UserFolderName\OneDrive - Mangano IT\Documents" -Confirm:$false -Force
    }
}

# Makes folder for logs and outputs logs
If(!(Test-Path -PathType container $LogsPath))
{
    New-Item -ItemType Directory -Path $LogsPath
}
$Logs | Out-File -FilePath "$LogsPath\ShareX-CopyConfig.txt" -Force -Confirm:$false