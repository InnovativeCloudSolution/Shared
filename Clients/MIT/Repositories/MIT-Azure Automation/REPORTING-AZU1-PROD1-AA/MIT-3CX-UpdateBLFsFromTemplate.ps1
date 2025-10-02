# 3CX Script to update all 9xxx numbers from a source, template user 7002.
# Created by Joshua Ceccato and Juan Moredo. 
# 2024/11/20

# Load Dependencies
. .\Function-ALL-General.ps1
. .\Function-ALL-MSGraph.ps1
. .\Function-ALL-3CX.ps1

Write-MessageLog -Message "Dependencies loaded successfully."

# Main Execution
function ProcessMainFunction {
    Write-MessageLog "Starting the 3CX Main Process."
    
    try {
        Initialize
    } catch {
        Write-ErrorLog -Message "Failed to initialize script configuration. Error: $_"
        return
    }

    try {
        Write-MessageLog -Message "Connecting to 3CX."
        Authenticate3CX
        #Authenticate3CXtest
        Write-MessageLog -Message "Connected to 3CX successfully."
    } catch {
        Write-ErrorLog -Message "Failed to connect to ConnectWise Manage. Error: $_"
        return
    }

    try {
        Update3CXNumbers
    } catch {
        Write-ErrorLog -Message "Failed to update 3CX BLF Keys. Error: $_"
        return
    }

    Write-MessageLog "3CX Main Process completed successfully."
}

# Function to Initialize Variables and Configuration
function Initialize {
    Write-MessageLog "Initializing script configuration for 3CX."
    $Global:Secret3CX = "MIT-3CX-7000"
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
    $Global:sourceDnNumber = "7002"
    $Global:dnPropertyName = "SHARED_BLFS"
    $Global:3CXURL = "https://voip.manganoit.com.au:5001"
    Write-MessageLog "Script configuration initialized."
}

# Function to Update 3CX Numbers
function Update3CXNumbers {
    Write-MessageLog "Updating 3CX numbers."

    $numbersToRemove = GetExcludedNumbers
    $dnNumbersArray = GetFiltered9xxxNumbers $numbersToRemove
    $sourceResponse = GetSourceDNProperty

    ProcessDNNumbers $dnNumbersArray $sourceResponse
    Write-MessageLog "3CX numbers update completed."
}

# Function to Get Excluded 9xxx Numbers
function GetExcludedNumbers {
    Write-MessageLog "Fetching list of numbers to exclude."
    $excludedNumbers = @(
        9000, 9028, 9050, 9064, 9066, 9077, 9088, 9099, 8100, 8101, 8099
    )
    Write-MessageLog "Excluded numbers: $($excludedNumbers -join ', ')"
    return $excludedNumbers
}

# Function to Get Filtered 9xxx Numbers
function GetFiltered9xxxNumbers {
    param (
        [array]$ExcludedNumbers
    )
    
    try {
        $string = "$3CXURL" + '/xapi/v1/Users?%24top=100&%24skip=0&%24orderby=Number&%24select=Number&%24filter=startswith(Number, ''9'')'
        $numbers = Invoke-RestMethod $string -Method 'GET' -Headers $headers
        $filteredNumbers = $numbers.value | Where-Object { $_.Number -notin $ExcludedNumbers } | ForEach-Object { $_.Number }
        return $filteredNumbers
    }
    catch {
        Write-ErrorLog "Failed to retrieve or filter 9xxx numbers: $_"
        throw
    }
}

# Function to Get Source DN Property Value
function GetSourceDNProperty {
    Write-MessageLog "Fetching source DN property value for dnNumber: $sourceDnNumber."
    
    try {
        $response = Invoke-RestMethod "$3CXURL/xapi/v1/DNProperties/Pbx.GetDNPropertyByName(dnNumber='$sourceDnNumber',name='$dnPropertyName')" -Method "GET" -Headers $headers
        Write-MessageLog "Source DN property value retrieved successfully."
        return $response
    }
    catch {
        Write-ErrorLog "Failed to retrieve source DN property value: $_"
        throw
    }
}

# Function to Process Each DN Number and Update Properties
function ProcessDNNumbers {
    param (
        [array]$DNNumbers,
        $SourceResponse
    )
    
    Write-MessageLog "Starting to process and update each DN number."
    
    foreach ($dnNumber in $DNNumbers) {
        Write-MessageLog "Processing dnNumber: $dnNumber"
        try {
            $targetResponse = GetTargetDNProperty $dnNumber
            UpdateDNProperty $dnNumber $targetResponse $SourceResponse
        }
        catch {
            Write-ErrorLog "Error processing dnNumber $($DNNumber): $_"
            continue
        }
    }
}

# Function to Get Target DN Property Value
function GetTargetDNProperty {
    param (
        [string]$DNNumber
    )
    
    Write-MessageLog "Fetching target DN property for dnNumber: $DNNumber."
    
    try {
        $response = Invoke-RestMethod "$3CXURL/xapi/v1/DNProperties/Pbx.GetDNPropertyByName(dnNumber='$DNNumber',name='$dnPropertyName')" -Method "GET" -Headers $headers
        Write-MessageLog "Target DN property value retrieved successfully for dnNumber: $DNNumber."
        return $response
    }
    catch {
        Write-ErrorLog "Failed to retrieve target DN property for dnNumber $($DNNumber): $_"
        throw
    }
}

# Function to Compare and Update DN Property Value
function UpdateDNProperty {
    param (
        [string]$DNNumber,
        $TargetResponse,
        $SourceResponse
    )
    
    if ($TargetResponse.Value -ne $SourceResponse.Value) {
        Write-MessageLog "Updating dnNumber: $DNNumber with new value from source dnNumber: $sourceDnNumber"

        # Prepare the body for the update
        $body = @{
            dnNumber = "$DNNumber"
            property = @{
                Description = $TargetResponse.Description
                Id          = $TargetResponse.Id
                Name        = $dnPropertyName
                Value       = $SourceResponse.Value
            }
        } | ConvertTo-Json

        try {
            # Update the DNProperty
            Invoke-RestMethod "$3CXURL/xapi/v1/DNProperties/Pbx.UpdateDNProperty" -Method "POST" -Headers $headers -Body $body
            Write-MessageLog "Successfully updated dnNumber: $DNNumber"
        }
        catch {
            Write-ErrorLog "Failed to update dnNumber $($DNNumber): $_"
        }
    } else {
        Write-MessageLog "No update needed for dnNumber: $DNNumber"
    }
}

# Run Main Process
ProcessMainFunction
