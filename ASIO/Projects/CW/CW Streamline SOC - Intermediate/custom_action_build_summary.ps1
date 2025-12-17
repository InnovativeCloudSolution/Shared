param(
    [string]$CVEDataJSON
)

$cveList = $CVEDataJSON | ConvertFrom-Json

$summary = "## Vulnerability Summary`n`n"
$summary += "**Total CVEs**: $($cveList.Count)`n`n"

$critical = $cveList | Where-Object { $_.severity -eq "CRITICAL" }
$high = $cveList | Where-Object { $_.severity -eq "HIGH" }
$medium = $cveList | Where-Object { $_.severity -eq "MEDIUM" }
$low = $cveList | Where-Object { $_.severity -eq "LOW" }

$summary += "**Severity Breakdown**:`n"
$summary += "- Critical: $($critical.Count)`n"
$summary += "- High: $($high.Count)`n"
$summary += "- Medium: $($medium.Count)`n"
$summary += "- Low: $($low.Count)`n`n"

$summary += "## CVE Details`n`n"

foreach ($cve in $cveList | Sort-Object -Property cvss_score -Descending) {
    $summary += "### $($cve.cve_id) - $($cve.severity)`n"
    $summary += "**CVSS Score**: $($cve.cvss_score)`n"
    $summary += "**Description**: $($cve.description)`n"
    $summary += "**Link**: $($cve.url)`n`n"
}

$summary

