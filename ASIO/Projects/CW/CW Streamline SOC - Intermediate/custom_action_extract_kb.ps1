param(
    [string]$NotesText
)

$kbPattern = 'Patch to be installed\s*:\s*(KB\d+)'
$kbMatch = [regex]::Match($NotesText, $kbPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

if ($kbMatch.Success) {
    @{ kb_number = $kbMatch.Groups[1].Value } | ConvertTo-Json
}
else {
    @{ kb_number = $null } | ConvertTo-Json
}

