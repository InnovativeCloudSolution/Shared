# Example script showing how to use CWM-Common.ps1

# Import the common module
. "$PSScriptRoot\CWM-Common.ps1"

# Initialize logging
$logPath = Initialize-Logging -LogName "CWM-Example"

Write-Log "=========================================" -Level "INFO"
Write-Log "Example Script Using Common Functions" -Level "INFO"
Write-Log "Log File: $logPath" -Level "INFO"
Write-Log "=========================================" -Level "INFO"
Write-Log "" -Level "INFO"

# Connect to CWM (uses hardcoded credentials from CWM-Common.ps1)
Connect-CWM

# Example: Get a company
$company = Get-CWMCompanyByIdentifier -Identifier "DBIT"
if ($company) {
    Write-Log "Found company: $($company.name) (ID: $($company.id))" -Level "SUCCESS"
} else {
    Write-Log "Company not found" -Level "WARNING"
}

# Write summary
Write-LogSummary
