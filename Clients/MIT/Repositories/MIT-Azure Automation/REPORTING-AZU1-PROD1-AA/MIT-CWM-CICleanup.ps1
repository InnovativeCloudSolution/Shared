
. .\Function-ALL-General.ps1
. .\Function-ALL-MSGraph.ps1
. .\Function-ALL-CWM.ps1
Write-MessageLog -Message "Dependencies loaded successfully."

function Initialize {
    Write-MessageLog -Message "Initializing script configuration."
    $Global:TenantUrl = "manganoit.com.au"
    $Global:CWMClientIdName = "MIT-CWMApi-ClientId"
    $Global:CWMPublicKeyName = "MIT-CWMApi-PubKey"
    $Global:CWMPrivateKeyName = "MIT-CWMApi-PrivateKey"
    $Global:CWMCompanyIdName = "MIT-CWMApi-CompanyId"
    $Global:CWMUrlName = "MIT-CWMApi-Url"
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
    } catch {
        Write-ErrorLog -Message "Failed to retrieve Azure Key Vault name. Error: $_"
        throw $_
    }
    Write-MessageLog -Message "Script configuration initialized."
}

function Restore-CWMInactiveStatus {
    param (
        [string]$ClientIdentifier,
        [string]$ConfigurationType,
        [string]$StatusName
    )

    $ClientCompany = Get-CWMCompany -condition "identifier='$ClientIdentifier'"
    $ConfigurationAICondition = "company/identifier='$($ClientCompany.identifier)' AND status/name='$StatusName' AND type/name='$ConfigurationType'"
    $ConfigurationAIs = Get-CWMCompanyConfiguration -condition $ConfigurationAICondition -all

    Write-MessageLog -Message "Starting configuration status revert process..."
    
    foreach ($ConfigurationAI in $ConfigurationAIs) {
        Write-MessageLog -Message "Processing CI#$($ConfigurationAI.id) - Retrieving audit trail..."
        $ConfigurationAIAudit = Get-CWMAuditTrail -type Configuration -id $ConfigurationAI.id
        
        $previousStatus = $null
        if ($ConfigurationAIAudit -and $ConfigurationAIAudit[0].text -match '<b>(.*?)</b> to <b>Automate Inactive</b>') {
            $previousStatus = $matches[1]
            Write-MessageLog -Message "Previous status found for CI#$($ConfigurationAI.id): '$previousStatus'"
        } else {
            Write-MessageLog -Message "No previous status match found for CI#$($ConfigurationAI.id)."
        }

        if (-not $previousStatus) {
            Write-MessageLog -Message "Skipping CI#$($ConfigurationAI.id) - No previous status found"
            continue
        }

        Write-MessageLog -Message "Attempting to revert status for CI#$($ConfigurationAI.id) to '$previousStatus'..."
        try {
            Update-CWMCompanyConfiguration -id $ConfigurationAI.id -Operation Replace -Path "status" -Value @{ name = $previousStatus }
            Write-MessageLog -Message "Successfully updated CI#$($ConfigurationAI.id) - Reverted status to '$previousStatus'"
        }
        catch {
            Write-MessageLog -Message "Failed to update CI#$($ConfigurationAI.id) - Error: $_"
        }
    }
    Write-MessageLog -Message "Configuration status revert process completed."
}

function Set-CWMParentConfiguration {
    param (
        [string]$ClientIdentifier,
        [string]$ConfigurationType
    )

    Write-MessageLog -Message "Updating parent configuration for $ConfigurationType in $ClientIdentifier."
    
    $ClientCompany = Get-CWMCompany -condition "identifier='$ClientIdentifier'"
    $ConfigurationChildrenCondition = "company/identifier='$($ClientCompany.identifier)' AND type/name='$ConfigurationType' AND parentConfigurationId!=null"
    $ConfigurationChildren = Get-CWMCompanyConfiguration -condition $ConfigurationChildrenCondition -all

    foreach ($ConfigurationChild in $ConfigurationChildren) {
        $NewParent = $ConfigurationChild
        Update-CWMCompanyConfiguration -id $NewParent.id -Operation remove -Path parentConfigurationId -Value null
        Write-MessageLog -Message "Removed CI#$($NewParent.id)-$($NewParent.name) from its parent and set as new parent."
    }
}

function ProcessMainFunction {
    try {
        Initialize
        Write-MessageLog -Message "Initialization completed successfully."
    }
    catch {
        Write-ErrorLog -Message "Failed during initialization. Error: $_"
        return
    }

    try {
        Write-MessageLog -Message "Connecting to ConnectWise Manage."
        ConnectCWM -AzKeyVaultName $AzKeyVaultName -CWMClientIdName $CWMClientIdName -CWMPublicKeyName $CWMPublicKeyName -CWMPrivateKeyName $CWMPrivateKeyName -CWMCompanyIdName $CWMCompanyIdName -CWMUrlName $CWMUrlName
        Write-MessageLog -Message "Connected to ConnectWise Manage successfully."
    }
    catch {
        Write-ErrorLog -Message "Failed to connect to ConnectWise Manage. Error: $_"
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

    Restore-CWMInactiveStatus -ClientIdentifier "CQL" -ConfigurationType "PC/Notebook" -StatusName "Automate Inactive"
    Restore-CWMInactiveStatus -ClientIdentifier "CQL" -ConfigurationType "Virtual Machine" -StatusName "Automate Inactive"
    Set-CWMParentConfiguration -ClientIdentifier "CQL" -ConfigurationType "PC/Notebook"
    Set-CWMParentConfiguration -ClientIdentifier "CQL" -ConfigurationType "Virtual Machine"

    Write-MessageLog -Message "ProcessMainFunction execution completed."
}

ProcessMainFunction
