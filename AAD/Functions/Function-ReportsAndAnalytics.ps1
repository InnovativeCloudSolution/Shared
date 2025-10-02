
# Usage Reports: Retrieve various usage reports across Microsoft 365 services
function Get-GraphUsageReports {
    param (
        [string]$reportType,
        [string]$period = "D7"  # Default to last 7 days
    )
    
    $url = "https://graph.microsoft.com/v1.0/reports/get${reportType}Usage(period='$period')"
    Invoke-GraphRequest -method GET -url $url
}

# Audit Logs: Access audit logs to track user and admin activities
function Get-GraphAuditLogs {
    param (
        [string]$activityType,
        [datetime]$startTime,
        [datetime]$endTime
    )
    
    $url = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=activityType eq '$activityType' and createdDateTime ge $($startTime.ToString('yyyy-MM-ddTHH:mm:ssZ')) and createdDateTime le $($endTime.ToString('yyyy-MM-ddTHH:mm:ssZ'))"
    Invoke-GraphRequest -method GET -url $url
}
