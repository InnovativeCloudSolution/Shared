function Write-MessageLog {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Output "$timestamp [$Level] $Message"
}

function Write-ErrorLog {
    param (
        [string]$Message
    )
    Write-MessageLog $Message "ERROR"
    throw $Message
}
