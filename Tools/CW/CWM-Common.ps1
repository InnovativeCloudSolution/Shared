# ConnectWise Manage Common Functions and Configuration
# This module contains shared functions and hardcoded parameters used across all CWM scripts

# ========================================
# HARDCODED PARAMETERS
# ========================================
$script:CWMUrl = "https://api-aus.myconnectwise.net"
$script:ApiVersion = "v4_6_release/apis/3.0"
$script:CompanyId = "dropbearit"
$script:PublicKey = "xAVcYWO20x5dRyG7"
$script:PrivateKey = "QUC1zTaMuUXbiJqX"
$script:ClientId = "1748c7f0-976c-4205-afa1-9bc9e1533565"

# ========================================
# LOGGING FUNCTIONS
# ========================================
function Write-Log {
    <#
    .SYNOPSIS
    Writes a log message to both console and log file
    
    .PARAMETER Message
    The message to log
    
    .PARAMETER Level
    The log level (INFO, SUCCESS, WARNING, ERROR)
    #>
    param(
        [string]$Message,
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    if ($script:LogPath) {
        Add-Content -Path $script:LogPath -Value $logEntry
    }
    
    switch ($Level) {
        "ERROR" {
            Write-Host $Message -ForegroundColor Red
            $script:ErrorCount++
        }
        "WARNING" {
            Write-Host $Message -ForegroundColor Yellow
            $script:WarningCount++
        }
        "SUCCESS" {
            Write-Host $Message -ForegroundColor Green
            $script:SuccessCount++
        }
        default {
            Write-Host $Message
        }
    }
}

function Initialize-Logging {
    <#
    .SYNOPSIS
    Initializes logging with a timestamped log file
    
    .PARAMETER LogName
    The base name for the log file (e.g., "CWM-Import")
    #>
    param(
        [string]$LogName = "CWM-Log"
    )
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $script:LogPath = ".logs\${LogName}_$timestamp.txt"
    
    $script:ErrorCount = 0
    $script:WarningCount = 0
    $script:SuccessCount = 0
    
    return $script:LogPath
}

function Write-LogSummary {
    <#
    .SYNOPSIS
    Writes a summary of the log statistics
    #>
    Write-Log "" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "SUMMARY:" -Level "INFO"
    Write-Log "  Total Successes: $script:SuccessCount" -Level "INFO"
    Write-Log "  Total Warnings:  $script:WarningCount" -Level "INFO"
    Write-Log "  Total Errors:    $script:ErrorCount" -Level "INFO"
    Write-Log "=========================================" -Level "INFO"
    Write-Log "" -Level "INFO"
    Write-Log "Log file saved to: $script:LogPath" -Level "SUCCESS"
}

# ========================================
# CONNECTION FUNCTIONS
# ========================================
function Connect-CWM {
    <#
    .SYNOPSIS
    Connects to ConnectWise Manage API and sets up headers
    #>
    $authString = "$script:CompanyId+$script:PublicKey`:$script:PrivateKey"
    $encodedAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($authString))
    
    $script:CWMHeaders = @{
        'Authorization' = "Basic $encodedAuth"
        'clientId' = $script:ClientId
        'Content-Type' = 'application/json'
    }
    
    $script:CWMBaseUrl = "$script:CWMUrl/$script:ApiVersion"
    
    Write-Log "Connected to ConnectWise Manage at $script:CWMUrl" -Level "SUCCESS"
}

# ========================================
# COMPANY FUNCTIONS
# ========================================
function Get-CWMCompanyByIdentifier {
    <#
    .SYNOPSIS
    Retrieves a company by its identifier
    
    .PARAMETER Identifier
    The company identifier to search for
    #>
    param(
        [string]$Identifier
    )
    
    $escapedIdentifier = $Identifier -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("identifier='$escapedIdentifier'")
    $uri = "$script:CWMBaseUrl/company/companies?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        
        if ($response.Count -gt 0) {
            return $response[0]
        } else {
            return $null
        }
    } catch {
        Write-Log "Failed to retrieve company '$Identifier': $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMCompanyStatus {
    <#
    .SYNOPSIS
    Retrieves a company status by name
    
    .PARAMETER StatusName
    The status name to search for
    #>
    param(
        [string]$StatusName
    )
    
    $escapedName = $StatusName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/company/companies/statuses?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) {
            return $response[0]
        }
        return $null
    } catch {
        Write-Log "Failed to retrieve company status: $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMCompanyType {
    <#
    .SYNOPSIS
    Retrieves a company type by name
    
    .PARAMETER TypeName
    The type name to search for
    #>
    param(
        [string]$TypeName
    )
    
    $escapedName = $TypeName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/company/companies/types?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) {
            return $response[0]
        }
        return $null
    } catch {
        Write-Log "Failed to retrieve company type: $_" -Level "ERROR"
        return $null
    }
}

function New-CWMCompany {
    <#
    .SYNOPSIS
    Creates a new company in ConnectWise Manage
    
    .PARAMETER CompanyData
    Hashtable containing the company data
    #>
    param(
        [hashtable]$CompanyData
    )
    
    $uri = "$script:CWMBaseUrl/company/companies"
    $body = $CompanyData | ConvertTo-Json -Depth 10
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Post -Body $body
        Write-Log "  Created company: $($response.name) (ID: $($response.id))" -Level "SUCCESS"
        return $response
    } catch {
        Write-Log "Failed to create company '$($CompanyData.name)': $_" -Level "ERROR"
        return $null
    }
}

# ========================================
# SITE FUNCTIONS
# ========================================
function Get-CWMSite {
    <#
    .SYNOPSIS
    Retrieves a site by company ID and site name
    
    .PARAMETER CompanyId
    The company ID
    
    .PARAMETER SiteName
    The site name to search for
    #>
    param(
        [int]$CompanyId,
        [string]$SiteName
    )
    
    $escapedName = $SiteName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/company/companies/$CompanyId/sites?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        
        if ($response.Count -gt 0) {
            return $response[0]
        } else {
            return $null
        }
    } catch {
        Write-Log "Failed to retrieve site: $_" -Level "ERROR"
        return $null
    }
}

function New-CWMSite {
    <#
    .SYNOPSIS
    Creates a new site for a company
    
    .PARAMETER CompanyId
    The company ID
    
    .PARAMETER SiteData
    Hashtable containing the site data
    #>
    param(
        [int]$CompanyId,
        [hashtable]$SiteData
    )
    
    $uri = "$script:CWMBaseUrl/company/companies/$CompanyId/sites"
    $body = $SiteData | ConvertTo-Json -Depth 10
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Post -Body $body
        Write-Log "    Created site: $($response.name) (ID: $($response.id))" -Level "SUCCESS"
        return $response
    } catch {
        Write-Log "  Failed to create site '$($SiteData.name)': $_" -Level "ERROR"
        return $null
    }
}

# ========================================
# CONTACT FUNCTIONS
# ========================================
function Get-CWMContact {
    <#
    .SYNOPSIS
    Retrieves a contact by company ID, first name, and last name
    
    .PARAMETER CompanyId
    The company ID
    
    .PARAMETER FirstName
    The contact's first name
    
    .PARAMETER LastName
    The contact's last name
    #>
    param(
        [int]$CompanyId,
        [string]$FirstName,
        [string]$LastName
    )
    
    $escapedFirst = $FirstName -replace "'", "''"
    $escapedLast = $LastName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("firstName='$escapedFirst' and lastName='$escapedLast'")
    $uri = "$script:CWMBaseUrl/company/contacts?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        
        if ($response.Count -gt 0) {
            foreach ($contact in $response) {
                if ($contact.company.id -eq $CompanyId) {
                    return $contact
                }
            }
            return $null
        } else {
            return $null
        }
    } catch {
        Write-Log "Failed to retrieve contact: $_" -Level "ERROR"
        return $null
    }
}

function New-CWMContact {
    <#
    .SYNOPSIS
    Creates a new contact in ConnectWise Manage
    
    .PARAMETER ContactData
    Hashtable containing the contact data
    #>
    param(
        [hashtable]$ContactData
    )
    
    $uri = "$script:CWMBaseUrl/company/contacts"
    $body = $ContactData | ConvertTo-Json -Depth 10
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Post -Body $body
        Write-Log "  Created contact: $($response.firstName) $($response.lastName) (ID: $($response.id))" -Level "SUCCESS"
        return $response
    } catch {
        Write-Log "Failed to create contact '$($ContactData.firstName) $($ContactData.lastName)': $_" -Level "ERROR"
        return $null
    }
}

# ========================================
# LOOKUP FUNCTIONS
# ========================================
function Get-CWMTerritory {
    <#
    .SYNOPSIS
    Retrieves a territory (location) by name
    
    .PARAMETER TerritoryName
    The territory name to search for
    #>
    param(
        [string]$TerritoryName
    )
    
    $escapedName = $TerritoryName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/system/locations?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) {
            return $response[0]
        }
        return $null
    } catch {
        Write-Log "Failed to retrieve territory: $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMMarket {
    <#
    .SYNOPSIS
    Retrieves a market by name
    
    .PARAMETER MarketName
    The market name to search for
    #>
    param(
        [string]$MarketName
    )
    
    $escapedName = $MarketName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/company/markets?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) {
            return $response[0]
        }
        return $null
    } catch {
        Write-Log "Failed to retrieve market: $_" -Level "ERROR"
        return $null
    }
}

# ========================================
# SERVICE BOARD FUNCTIONS
# ========================================
function Get-CWMServiceBoard {
    <#
    .SYNOPSIS
    Retrieves a service board by name
    
    .PARAMETER BoardName
    The board name to search for
    #>
    param(
        [string]$BoardName
    )
    
    $escapedName = $BoardName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/service/boards?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) {
            return $response[0]
        }
        return $null
    } catch {
        Write-Log "Failed to retrieve service board: $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMServiceType {
    <#
    .SYNOPSIS
    Retrieves a service type by board ID and type name
    
    .PARAMETER BoardId
    The board ID
    
    .PARAMETER TypeName
    The type name to search for
    #>
    param(
        [int]$BoardId,
        [string]$TypeName
    )
    
    $escapedName = $TypeName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/service/boards/$BoardId/types?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) {
            return $response[0]
        }
        return $null
    } catch {
        Write-Log "Failed to retrieve service type: $_" -Level "ERROR"
        return $null
    }
}

function New-CWMServiceType {
    <#
    .SYNOPSIS
    Creates a new service type
    
    .PARAMETER BoardId
    The board ID
    
    .PARAMETER TypeData
    Hashtable containing the type data
    #>
    param(
        [int]$BoardId,
        [hashtable]$TypeData
    )
    
    $uri = "$script:CWMBaseUrl/service/boards/$BoardId/types"
    $body = $TypeData | ConvertTo-Json -Depth 10
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Post -Body $body
        Write-Log "  Created type: $($response.name) (ID: $($response.id))" -Level "SUCCESS"
        return $response
    } catch {
        Write-Log "Failed to create type '$($TypeData.name)': $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMServiceSubType {
    <#
    .SYNOPSIS
    Retrieves a service subtype by board ID and subtype name
    
    .PARAMETER BoardId
    The board ID
    
    .PARAMETER SubTypeName
    The subtype name to search for
    #>
    param(
        [int]$BoardId,
        [string]$SubTypeName
    )
    
    $escapedName = $SubTypeName -replace "'", "''"
    $encodedCondition = [Uri]::EscapeDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/service/boards/$BoardId/subtypes?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) {
            return $response[0]
        }
        return $null
    } catch {
        Write-Log "Failed to retrieve service subtype: $_" -Level "ERROR"
        return $null
    }
}

function New-CWMServiceSubType {
    <#
    .SYNOPSIS
    Creates a new service subtype
    
    .PARAMETER BoardId
    The board ID
    
    .PARAMETER SubTypeData
    Hashtable containing the subtype data
    #>
    param(
        [int]$BoardId,
        [hashtable]$SubTypeData
    )
    
    $uri = "$script:CWMBaseUrl/service/boards/$BoardId/subtypes"
    $body = $SubTypeData | ConvertTo-Json -Depth 10
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Post -Body $body
        Write-Log "    Created subtype: $($response.name) (ID: $($response.id))" -Level "SUCCESS"
        return $response
    } catch {
        Write-Log "  Failed to create subtype '$($SubTypeData.name)': $_" -Level "ERROR"
        return $null
    }
}

function Get-CWMServiceItem {
    <#
    .SYNOPSIS
    Retrieves a service item by board ID and item name
    
    .PARAMETER BoardId
    The board ID
    
    .PARAMETER ItemName
    The item name to search for
    #>
    param(
        [int]$BoardId,
        [string]$ItemName
    )
    
    $escapedName = $ItemName -replace "'", "''"
    $encodedCondition = [Uri]::EscaceDataString("name='$escapedName'")
    $uri = "$script:CWMBaseUrl/service/boards/$BoardId/items?conditions=$encodedCondition"
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Get
        if ($response.Count -gt 0) {
            return $response[0]
        }
        return $null
    } catch {
        Write-Log "Failed to retrieve service item: $_" -Level "ERROR"
        return $null
    }
}

function New-CWMServiceItem {
    <#
    .SYNOPSIS
    Creates a new service item
    
    .PARAMETER BoardId
    The board ID
    
    .PARAMETER ItemData
    Hashtable containing the item data
    #>
    param(
        [int]$BoardId,
        [hashtable]$ItemData
    )
    
    $uri = "$script:CWMBaseUrl/service/boards/$BoardId/items"
    $body = $ItemData | ConvertTo-Json -Depth 10
    
    try {
        $response = Invoke-RestMethod -Uri $uri -Headers $script:CWMHeaders -Method Post -Body $body
        Write-Log "      Created item: $($response.name) (ID: $($response.id))" -Level "SUCCESS"
        return $response
    } catch {
        Write-Log "    Failed to create item '$($ItemData.name)': $_" -Level "ERROR"
        return $null
    }
}

# Export functions for use in other scripts
Export-ModuleMember -Function *
