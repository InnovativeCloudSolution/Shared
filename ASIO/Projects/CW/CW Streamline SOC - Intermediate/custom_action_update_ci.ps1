param(
    [int]$CIID,
    [string]$CVEDataJSON,
    [string]$KBNumber
)

$apiUrl = "$env:CW_URL/v4_6_release/apis/3.0/company/configurations/$CIID"
$headers = @{
    "Authorization" = "Basic $env:CW_AUTH"
    "Content-Type" = "application/json"
}

$cveData = $CVEDataJSON | ConvertFrom-Json

$existingCI = Invoke-RestMethod -Uri $apiUrl -Headers $headers -Method Get
$existingCVEs = if ($existingCI.customFields.Pending_CVEs) { 
    $existingCI.customFields.Pending_CVEs | ConvertFrom-Json 
} else { 
    @() 
}

$existingKBs = if ($existingCI.customFields.Pending_KB_Patches) {
    $existingCI.customFields.Pending_KB_Patches -split ',' | ForEach-Object { $_.Trim() }
} else {
    @()
}

$mergedCVEs = $existingCVEs + $cveData | Sort-Object -Property cve_id -Unique
if ($KBNumber -notin $existingKBs) {
    $existingKBs += $KBNumber
}

$criticalCount = ($mergedCVEs | Where-Object { $_.severity -eq "CRITICAL" }).Count
$highCount = ($mergedCVEs | Where-Object { $_.severity -eq "HIGH" }).Count
$mediumCount = ($mergedCVEs | Where-Object { $_.severity -eq "MEDIUM" }).Count
$lowCount = ($mergedCVEs | Where-Object { $_.severity -eq "LOW" }).Count

$summary = "Total Pending: $($mergedCVEs.Count) CVEs (Critical: $criticalCount, High: $highCount, Medium: $mediumCount, Low: $lowCount)`n"
$summary += "Patches: $($existingKBs -join ', ')`n"
$summary += "Last Updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"

$operations = @(
    @{ op = "replace"; path = "customFields/Pending_CVEs"; value = ($mergedCVEs | ConvertTo-Json -Compress) }
    @{ op = "replace"; path = "customFields/Pending_KB_Patches"; value = ($existingKBs -join ',') }
    @{ op = "replace"; path = "customFields/Vulnerability_Summary"; value = $summary }
    @{ op = "replace"; path = "customFields/Last_Vulnerability_Scan"; value = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss') }
    @{ op = "replace"; path = "customFields/Patch_Status_$KBNumber"; value = "Pending" }
)

try {
    $result = Invoke-RestMethod -Uri $apiUrl -Headers $headers -Method Patch -Body ($operations | ConvertTo-Json)
    @{ success = $true; message = "CI updated successfully" } | ConvertTo-Json
}
catch {
    @{ success = $false; message = $_.Exception.Message } | ConvertTo-Json
}

