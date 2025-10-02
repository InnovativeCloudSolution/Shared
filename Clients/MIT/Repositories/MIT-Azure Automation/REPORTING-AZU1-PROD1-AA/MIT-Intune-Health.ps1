# Load Dependencies
. .\Function-ALL-General.ps1
. .\Function-ALL-MSGraph.ps1
. .\Function-ALL-MSGraph-Intune.ps1
. .\Function-ALL-CWM.ps1
Write-MessageLog -Message "Dependencies loaded successfully."

# Main Execution
function ProcessMainFunction {
    Write-MessageLog -Message "Starting the Main Process."
    
    try {
        Initialize
    }
    catch {
        Write-ErrorLog -Message "Failed to initialize script configuration. Error: $_"
        return
    }

    try {
        Write-MessageLog -Message "Connecting to Microsoft Graph."
        ConnectMSGraph -AzKeyVaultName $AzKeyVaultName -TenantUrl $TenantUrl -ClientIdSecretName $ClientIdSecretName -ClientSecretSecretName $ClientSecretSecretName
        Write-MessageLog -Message "Connected to Microsoft Graph successfully."
    }
    catch {
        Write-ErrorLog -Message "Failed to connect to Microsoft Graph. Error: $_"
        return
    }

    try {
        Write-MessageLog -Message "Connecting to ConnectWise Manage."
        ConnectCWMTest
        ConnectCWM -AzKeyVaultName $AzKeyVaultName -CWMClientIdName $CWMClientIdName -CWMPublicKeyName $CWMPublicKeyName -CWMPrivateKeyName $CWMPrivateKeyName -CWMCompanyIdName $CWMCompanyIdName -CWMUrlName $CWMUrlName
        Write-MessageLog -Message "Connected to ConnectWise Manage successfully."
    }
    catch {
        Write-ErrorLog -Message "Failed to connect to ConnectWise Manage. Error: $_"
        return
    }

    $IntuneHealthResults = @()

    Write-MessageLog "Querying Intune Connector Health..."
    $IntuneHealthResults += Get-IntuneConnectorHealth
    
    Write-MessageLog "Querying Apple Push Notification Certificate..."
    $IntuneHealthResults += Get-ApplePushNotificationCertificate
    
    Write-MessageLog "Querying Apple VPP Token..."
    $IntuneHealthResults += Get-AppleVppTokens
    
    Write-MessageLog "Querying Apple DEP Token..."
    $IntuneHealthResults += Get-AppleDepTokens
    
    Write-MessageLog "Querying Managed Google Play settings..."
    $IntuneHealthResults += Get-ManagedGooglePlay
    
    Write-MessageLog "Querying Autopilot Connector..."
    $IntuneHealthResults += Get-Autopilot
    
    Write-MessageLog "Querying Mobile Threat Defense Connector..."
    $IntuneHealthResults += Get-MobileThreatDefenseConnectors
    
    if ($IntuneHealthResults.Count -gt 0) {
        $CWMCompany = Get-CWM-Company -ClientIdentifier $ClientIdentifier
        $CWMContact = Get-CWM-Contact -ClientIdentifier $ClientIdentifier -ContactName $ContactName
        $CWMBoard = Get-CWM-Board -ServiceBoardName $ServiceBoardName
        $CWMBoardStatus = Get-CWM-BoardStatus -ServiceBoardName $ServiceBoardName -ServiceBoardStatusName $ServiceBoardStatusName
        $CWMBoardType = Get-CWM-BoardType -ServiceBoardName $ServiceBoardName -ServiceBoardTypeName $ServiceBoardTypeName
        $CWMBoardSubType = Get-CWM-BoardSubType -ServiceBoardName $ServiceBoardName -ServiceBoardSubTypeName $ServiceBoardSubTypeName
        $CWMBoardItem = Get-CWM-BoardItem -ServiceBoardName $ServiceBoardName -ServiceBoardTypeName $ServiceBoardTypeName -ServiceBoardSubTypeName $ServiceBoardSubTypeName -ServiceBoardItemName $ServiceBoardItemName
        $CWMPriority = Get-CWM-Priority -PriorityName $PriorityName
    
        $CWMTicketSummary = "Intune Health Check: Issues Detected"
        $CWMTicketNoteInitial = "The following issues were identified during the Intune Health Check:`n`n"
    
        foreach ($issue in $IntuneHealthResults) {
            if ($null -ne $issue.state) {
                $CWMTicketNoteInitial += "Intune Connector Health: State is '$($issue.state)'.`n"
            }
            elseif ($null -ne $issue.expirationDateTime) {
                $CWMTicketNoteInitial += "Apple Push Notification Certificate: Expiration Date - $($issue.expirationDateTime).`n"
            }
            elseif ($null -ne $issue.lastSyncStatus -or $null -ne $issue.state) {
                $CWMTicketNoteInitial += "Apple VPP Token: Last Sync Status - $($issue.lastSyncStatus), State - $($issue.state).`n"
            }
            elseif ($null -ne $issue.lastSyncErrorCode -or $null -ne $issue.tokenExpirationDateTime) {
                $CWMTicketNoteInitial += "Apple DEP Token: Error Code - $($issue.lastSyncErrorCode), Expiration Date - $($issue.tokenExpirationDateTime).`n"
            }
            elseif ($null -ne $issue.bindStatus -or $null -ne $issue.lastAppSyncStatus) {
                $CWMTicketNoteInitial += "Managed Google Play: Bind Status - $($issue.bindStatus), App Sync Status - $($issue.lastAppSyncStatus).`n"
            }
            elseif ($null -ne $issue.syncStatus) {
                $CWMTicketNoteInitial += "Autopilot: Sync Status - $($issue.syncStatus).`n"
            }
            elseif ($null -ne $issue.partnerState) {
                $CWMTicketNoteInitial += "Mobile Threat Defense Connector: Partner State - $($issue.partnerState).`n"
            }
        }
    
        New-CWM-Ticket -CWMCompany $CWMCompany `
            -CWMContact $CWMContact `
            -CWMBoard $CWMBoard `
            -CWMTicketSummary $CWMTicketSummary `
            -CWMTicketNoteInitial $CWMTicketNoteInitial `
            -CWMBoardStatus $CWMBoardStatus `
            -CWMBoardType $CWMBoardType `
            -CWMBoardSubType $CWMBoardSubType `
            -CWMBoardItem $CWMBoardItem `
            -CWMPriority $CWMPriority
    }
    else {
        Write-MessageLog "All checks passed successfully. No issues found!"
    }
    

    try {
        Write-MessageLog -Message "Disconnecting from ConnectWise Manage."
        Disconnect-CWM
        Write-MessageLog -Message "Disconnected from ConnectWise Manage successfully."
    }
    catch {
        Write-ErrorLog -Message "Failed to disconnect from ConnectWise Manage. Error: $_"
    }

    Write-MessageLog -Message "Main Process completed successfully."
}

# Function to Initialize Variables and Configuration
function Initialize {
    Write-MessageLog -Message "Initializing script configuration."
    $Global:TenantUrl = 'manganoit.com.au'
    $Global:ClientIdSecretName = "MIT-AutomationApp-ClientID"
    $Global:ClientSecretSecretName = "MIT-AutomationApp-ClientSecret"
    try {
        Write-MessageLog -Message "Retrieving Azure Key Vault name from automation variables."
        $Global:AzKeyVaultName = Get-AutomationVariable -Name 'AzKeyVaultName'
    
        if ($null -eq $AzKeyVaultName) {
            Write-ErrorLog -Message "Azure Key Vault name is missing. Check your automation variables."
            throw "Azure Key Vault name is not defined."
        }
    
        Write-MessageLog -Message "Azure Key Vault name retrieved successfully: $AzKeyVaultName."
    }
    catch {
        Write-ErrorLog -Message "Failed to retrieve Azure Key Vault name. Error: $_"
        throw $_
    }
    $Global:AccessToken = Get-MSGraph-BearerToken -AzKeyVaultName $AzKeyVaultName -TenantUrl $TenantUrl -ClientIdSecretName $ClientIdSecretName -ClientSecretSecretName $ClientSecretSecretName
    $Global:CWMClientIdName = "MIT-CWMApi-ClientId"
    $Global:CWMPublicKeyName = "MIT-CWMApi-PubKey"
    $Global:CWMPrivateKeyName = "MIT-CWMApi-PrivateKey"
    $Global:CWMCompanyIdName = "MIT-CWMApi-CompanyId"
    $Global:CWMUrlName = "MIT-CWMApi-Url"
    $Global:ClientIdentifier = "MIT"
    $Global:ContactName = "Paul Mangano"
    $Global:ServiceBoardName = "HelpDesk (MS)"
    $Global:ServiceBoardStatusName = "Scheduling Required"
    $Global:ServiceBoardTypeName = "Incident"
    $Global:ServiceBoardSubTypeName = "Server"
    $Global:ServiceBoardItemName = "Azure/Cloud"
    $Global:PriorityName = "P2 Critical Response"
    Write-MessageLog -Message "Script configuration initialized."
}

# Run Main Process
ProcessMainFunction
