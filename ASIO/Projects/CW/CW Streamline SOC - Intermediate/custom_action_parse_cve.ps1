param(
    [int]$TicketID
)

$apiUrl = "$env:CW_URL/v4_6_release/apis/3.0/service/tickets/$TicketID/notes"
$headers = @{
    "Authorization" = "Basic $env:CW_AUTH"
    "Content-Type" = "application/json"
}

$notes = Invoke-RestMethod -Uri $apiUrl -Headers $headers -Method Get
$allNotes = ($notes | ForEach-Object { $_.text }) -join "`n"

$cvePattern = 'CVE-\d{4}-\d{4,7}'
$cveMatches = [regex]::Matches($allNotes, $cvePattern)
$cveList = $cveMatches | ForEach-Object { $_.Value } | Select-Object -Unique

$result = @{
    "cve_ids" = $cveList
    "count" = $cveList.Count
}

$result | ConvertTo-Json

