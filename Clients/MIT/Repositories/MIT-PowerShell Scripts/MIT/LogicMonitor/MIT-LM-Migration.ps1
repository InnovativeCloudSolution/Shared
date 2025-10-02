Param (
    [Parameter(Mandatory=$False)][boolean] $add = $True,
    [Parameter(Mandatory=$False)][string] $UserName = "aadds\svc.logicmonitor",
    [Parameter(Mandatory=$False)][string] $Password,
    [Parameter(Mandatory=$False)][string] $Path
)

$version = $PSVersionTable.PSVersion.Major
if ([int]$version -lt 5) {
    Write-Output "PowerShell version 5 is the minimum mandatory requirement to run this script"
    exit
}

if ($help -eq $true) {
    Write-Output "Usage : [-help|-add|-remove][-UserName userName][-Password password][-Path path]
                  -help         Help Prompt     - Show this message.
                  -add|-remove  Operation Flag  - Operation Flag is required '-add' for adding and '-remove' for reversal.
                  -UserName     Non-Admin User  - Name of Non-Admin User under which want to move Collector services. Mandatory when not '-help'.
                  -Password     Password        - Password of the Non-Admin User.
                  -Path         Install path    - Installation path of the Collector (default: `"C:\Program Files\LogicMonitor`")

                  Example 0 : .\Windows_NonAdmin_Config.ps1 -help
                  Example 1 : .\Windows_NonAdmin_Config.ps1 -add -UserName LOGICMONITOR\MyUserName
                  Example 2 : .\Windows_NonAdmin_Config.ps1 -remove -UserName LOGICMONITOR\MyUserName -Password MySecurePassword -Path `"C:\Program Files\Path\LogicMonitor`"
                  "
    exit
}

if (($help -eq $false) -and ($UserName -eq "" -or $UserName -eq $null)) {
    Write-Output "UserName is a mandatory argument, Kindly pass correct UserName. Check `".\Windows_NonAdmin_Config.ps1 -help`" for further help."
    exit
}

$ComputerName = $env:COMPUTERNAME

$isCollectorHost = $false

$agentService = Get-Service -Name "logicmonitor-agent" -ErrorAction SilentlyContinue
$watchdogService = Get-Service -Name "logicmonitor-watchdog" -ErrorAction SilentlyContinue
if ($agentService -and $watchdogService) {
    $isCollectorHost = $true
    Write-Output "LogicMonitor Agent and Watchdog Services found on this host $ComputerName"
}

Function Stop-Services {
    Write-Output "Stopping services..."
    sc.exe stop "logicmonitor-agent"
    sc.exe stop "logicmonitor-watchdog"

    Get-Service -Name "logicmonitor-agent" | Wait-ForServiceToStop
    Get-Service -Name "logicmonitor-watchdog" | Wait-ForServiceToStop
}

Function Wait-ForServiceToStop {
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [System.ServiceProcess.ServiceController] $service
    )
    
    while ($service.Status -eq 'StopPending') {
        Write-Output "Waiting for service $($service.DisplayName) to stop..."
        Start-Sleep -Seconds 5
        $service.Refresh()
    }
    
    if ($service.Status -eq 'Stopped') {
        Write-Output "Service $($service.DisplayName) stopped successfully."
    } else {
        Write-Output "Service $($service.DisplayName) did not stop successfully."
    }
}

Function Set-UserLocalGroup {
    [cmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,Position=0)][string] $user,
        [Parameter(Mandatory=$True,Position=1)][string] $operation
    )

    $LocalGroups = "Performance Monitor Users", "Event Log Readers", "Remote Management Users", "Distributed COM Users"

    ForEach ($group in $LocalGroups) {
        net localgroup $group $user /$operation 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) {
            Write-Output "User successfully processed for the $group, Operation = $operation"
        }
        else {
            Write-Output "WARNING - User not found in the $group"
        }
    }
}

Function DoCollectorHostConfig {
    [cmdletBinding()]
    Param (
        [parameter(Mandatory=$true, Position=0, ValueFromPipelineByPropertyName=$true)][string] $UserName
    )

    # Handle password input if not provided
    if ($Password -eq $null -or $Password -eq "") {
        $Credential = Get-Credential -UserName $UserName -Message "Enter the password for $UserName"
        $Password = $Credential.GetNetworkCredential().Password
    }

    # Set the agent path, defaulting if not provided
    $agentPath = if ($Path -eq $null -or $Path -eq "") { "C:\Program Files\LogicMonitor" } else { $Path }

    # Stop services before making changes
    Stop-Services

    # Assign rights to the specified user
    $Rights = "SeServiceLogonRight","SeChangeNotifyPrivilege"
    Write-Output "Granted rights to $UserName"

    # Start services after changes
    sc.exe start "logicmonitor-agent"
    sc.exe start "logicmonitor-watchdog"
    Write-Output "The collector services are now switched to run under user $UserName."
}

if ($add) {
    Write-Output "Starting configuration tasks for $ComputerName"

    if ($isCollectorHost -eq $true) {
        DoCollectorHostConfig -UserName "$UserName"
    }
    Write-Output "Configuration completed."
} else {
    Write-Output "Starting reversal tasks for $ComputerName"
    Set-UserLocalGroup $UserName "delete"
    Write-Output "Reversal completed."
}
