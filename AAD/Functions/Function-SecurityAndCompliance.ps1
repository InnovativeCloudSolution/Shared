
# Security: Access and manage security alerts, secure scores, and other security-related information
# Compliance: Interact with compliance-related data, such as eDiscovery cases

function Get-GraphSecurityAlerts {
    param (
        [int]$top = 10
    )
    
    $url = "https://graph.microsoft.com/v1.0/security/alerts?`$top=$top"
    Invoke-GraphRequest -method GET -url $url
}
