# Logging variables
$LogsPath = "C:\Logs"
$Logs = ''
$Date = Get-Date -Format FileDateTime

# Make list of local admin accounts
$administrators = @( ([ADSI]"WinNT://./Administrators").psbase.Invoke('Members') | `
ForEach-Object { $_.GetType().InvokeMember('AdsPath','GetProperty',$null,$($_),$null) } ) -match '^WinNT';

# Fix usernames to match how they show in the localgroup
$administrators = $administrators -replace "WinNT://",""
$administrators = $administrators -replace "AzureAD/","AzureAD\"
$administrators = $administrators -replace "mits/","mits\"
$Logs += "$administrators `n"

# Remove all Azure AD users as local admin users
foreach ($administrator in $administrators) {
    $Logs += "`nChecking to see if $administrator matches criteria..."
    if ($administrator -like "AzureAD\*" -or $administrator -like "mits\*") { 
        Remove-LocalGroupMember -group "administrators" -member $administrator 
        $Logs += "`nRemoved $administrator from on-device administrators."
    }
}

# Makes folder for logs and outputs logs
If(!(Test-Path -PathType container $LogsPath))
{
    New-Item -ItemType Directory -Path $LogsPath
}
$Logs | Out-File -FilePath "$LogsPath\RemoveAdmins-$Date.txt" -Force -Confirm:$false