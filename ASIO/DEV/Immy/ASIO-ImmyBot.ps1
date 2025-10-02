$Global:TenantURL = "manganoit.com.au"
$Global:ClientID = "812408b3-6b87-418b-95ed-b036e2a67402"
$Global:ClientSecret = "Pmk8Q~FOA3Kd5QOR~9Bm.xiRBWNyJh1.TxjEcbIV"
$Global:BaseURL = "https://mit.immy.bot"
$ComputerName = ""
$UserEmail = "Brad.Swan@manganoit.com.au"
$SoftwareName = "PrinterLogic"

Function Get-ImmyBotApiAuthToken {
	param (
		$TenantURL,
		$ClientID,
		$ClientSecret,
		$BaseURL
	)

	$ContentType = 'application/x-www-form-urlencoded'
	$Scope = "$BaseURL/.default"
	$GrantType = "client_credentials"
	
	$ApiArguments = @{
		Uri             = "https://login.microsoftonline.com/$TenantURL/oauth2/v2.0/token"
		Method          = 'POST'
		Headers         = @{'Content-Type' = $ContentType }
		Body            = "client_id=$ClientID&client_secret=$ClientSecret&grant_type=$GrantType&scope=$Scope"
		UseBasicParsing = $true
	}
	try {
		return Invoke-RestMethod @ApiArguments
	}
	catch { throw }
}

Function Invoke-ImmyBotRestMethod {
	param(
		$BaseURL,
		$Endpoint,
		$Method,
		$BearerToken,
		$Body
	)

	$Endpoint = $Endpoint.TrimStart('/')
	$ContentType = "application/json"
	$ApiArguments = @{}

	$ApiArguments = @{
		Uri             = "$($BaseURL)/$($Endpoint)"
		Method          = $Method
		Headers         = @{
			'Authorization' = $BearerToken;
			'Content-Type'  = $ContentType
		}
		UseBasicParsing = $true
	}

	if ($Body) {
		if ($Body -is [Hashtable]) {
			$Body = $Body | ConvertTo-Json -Depth 100
		}
		$ApiArguments.Body = $body
	}
		
	try {
		return Invoke-RestMethod @ApiArguments
	}
	catch { throw }
}

$Global:BearerToken = "Bearer $((Get-ImmyBotApiAuthToken -TenantURL $TenantURL -ClientID $ClientID  -ClientSecret $ClientSecret -BaseURL $BaseURL).access_token)"

$Softwares = Invoke-ImmyBotRestMethod -BaseURL $BaseURL -Method "GET"  -Endpoint "/api/v1/software/local" -BearerToken $BearerToken
$SelectedSoftware = $Softwares | Where-Object { $_.name -eq "*$SoftwareName*" }
if (-not $SelectedSoftware){
	$Softwares = Invoke-ImmyBotRestMethod -BaseURL $BaseURL -Method "GET" -Endpoint "/api/v1/software/global" -BearerToken $BearerToken
	$SelectedSoftware = $Softwares | Where-Object { $_.name -like "*$SoftwareName*" }
}

if($UserEmail){
	$AllComputers = Invoke-ImmyBotRestMethod -BaseURL $BaseURL -Method "GET" -Endpoint "/api/v1/computers/dx" -BearerToken $BearerToken
	$SelectedComputers = $AllComputers.data | Where-Object { $_.primaryUserEmail -eq $UserEmail }
	$SelectedComputers | Select-Object computerName, primaryUserEmail
}

# Note: Many of these fields can likely be omitted but are included for completeness
Invoke-ImmyBotRestMethod -BaseURL $BaseURL -Method "POST" -Endpoint "/api/v1/run-immy-service" -BearerToken $BearerToken `
-Body @{
    fullMaintenance = $false # When true, the triggered session will have all software/tasks applied to the machine. If false, it will be limited to one. You must provide a maintenanceParams property to specify the one you want
    resolutionOnly = $false # When this is true, we "resolve" the desired state of the software against the deployments. This is is useful for determining if the user/computer should have the software installed. The computer does not need to be online for resolution to run.
    detectionOnly = $false # Detection just detects what version of the software exists on the machine, if any. Both resolution and detection are required to determine what action is necessary to acheive the desired state. The computer must be online for detection to run.
    inventoryOnly = $false # Session will end after the inventory scripts run
    runInventoryInDetection = $false # When this is true, all inventory scripts will be run during detection. When this is false, only the Software Inventory script is run
    cacheOnly = $false # Skips Software Inventory script and uses the most recent software inventory to determine the currently installed version
    useWinningDeployment = $false # When true, the desiredSoftwareState in the maintenanceParams below is ignored
    deploymentId = $null # If useWinningDeployment is false, you can specify a deployment here. When null,use maintenanceParams below, or if maintenanceParams are not specified resolution will determine the "winning" deployment
    deploymentType = $null # The deploymentId is in what database? 0 - Global Database (Recommended deployments), 1 - Local (Typical)
    maintenanceParams = @{
        maintenanceIdentifier = "$($SelectedSoftware.Id)"
        maintenanceType = 0
        repair = $false
        desiredSoftwareState = 5
        <#
            DesiredSoftwareState.NoAction => 0,
            DesiredSoftwareState.NotPresent => 1,
            DesiredSoftwareState.ThisVersion => 2,
            DesiredSoftwareState.OlderOrEqualVersion => 3,
            DesiredSoftwareState.LatestVersion => 4,
            DesiredSoftwareState.NewerOrEqualVersion => 5, # This is the default. It should be called "Newer or Equal to the _expected_ version". Sure you would think LatestVersion would be the default but LatestVersion refers to the latest version in our database (before dynamic versions) and NewerOrEqual was added to prevent the action from being marked as failed if the software self-updates during installation to a version newer than we expected.
            DesiredSoftwareState.AnyVersion => 6,
        #>
        maintenanceTaskMode = 0
    }
    skipBackgroundJob = $true # true bypasses the concurrent session limit. Careful, you can quickly overload your instance if you abuse this
    rebootPreference = 1 # Force = -1, Normal = 0, Suppress = 1
    scheduleExecutionAfterActiveHours = $false
    useComputersTimezoneForExecution = $false
    offlineBehavior = 2 #  Skip = 1, ApplyOnConnect = 2
    suppressRebootsDuringBusinessHours = $false
    sendDetectionEmail = $false
    sendDetectionEmailWhenAllActionsAreCompliant = $false
    sendFollowUpEmail = $false
    sendFollowUpOnlyIfActionNeeded = $false
    showRunNowButton = $false
    showPostponeButton = $false
    showMaintenanceActions = $false
    computers = @($SelectedComputers | ForEach-Object { @{ computerId = $_.id } })
    tenants = @() # If the specified maintenanceItem is a Cloud Task you would specify this instead of computers
}




Target.Type
All Computers + Desiredstate -ne Latest, Update if Found	
Azure Group Target.Target add to group then run maintenance on 1 item


Update if found = just Update
Latest = install and Update